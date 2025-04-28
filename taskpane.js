(function () {
    "use strict";

    let currentItem = null;
    let userDomain = null;
    let lastClientDomain = null;
    let isValidating = false;
    let autoValidationInterval = null;
    
    // Email cache to avoid unnecessary revalidation
    const emailValidationCache = new Map();
    
    // Common domains for typo detection
    const COMMON_DOMAINS = [
        'gmail.com', 'yahoo.com', 'outlook.com', 'hotmail.com', 'aol.com', 
        'icloud.com', 'protonmail.com', 'msn.com', 'live.com', 'me.com'
    ];
    
    // Default client domain
    const DEFAULT_CLIENT_DOMAIN = "xyz.com";

    Office.onReady(function () {
        if (Office.context.mailbox) {
            // Initialize the UI
            initializeUI();
            
            // Get the user's email domain
            const userEmail = Office.context.mailbox.userProfile.emailAddress;
            userDomain = userEmail.split('@')[1];

            // Store the current item
            currentItem = Office.context.mailbox.item;
            
            // Check for saved settings
            loadSavedSettings();
            
            // Set up event handlers
            setupEventHandlers();
            
            // Start background monitoring
            startBackgroundMonitoring();
        }
    });
    
    function initializeUI() {
        document.getElementById("app-body").style.display = "block";
        document.getElementById("sideload-msg").style.display = "none";
        
        // Initially hide warning UI
        document.getElementById("warning-ui").style.display = "none";
    }
    
    function loadSavedSettings() {
        // Try to load saved client domain
        try {
            const savedDomain = Office.context.roamingSettings.get("lastClientDomain");
            if (savedDomain) {
                lastClientDomain = savedDomain;
            } else {
                lastClientDomain = DEFAULT_CLIENT_DOMAIN;
            }
        } catch (error) {
            // If settings aren't available, use default
            lastClientDomain = DEFAULT_CLIENT_DOMAIN;
            console.log("Could not load settings, using defaults");
        }
    }
    
    function saveSettings() {
        try {
            Office.context.roamingSettings.set("lastClientDomain", lastClientDomain);
            Office.context.roamingSettings.saveAsync();
        } catch (error) {
            console.log("Could not save settings: " + error.message);
        }
    }
    
    function setupEventHandlers() {
        // Event handlers for warning UI buttons
        document.getElementById("go-back-button").addEventListener("click", hideWarningUI);
        document.getElementById("proceed-button").addEventListener("click", proceedAnyway);
        
        // Client selection event handler
        document.getElementById("client-select").addEventListener("change", function() {
            lastClientDomain = this.value;
            saveSettings();
            validateEmailsBackground();
        });
    }

    function startBackgroundMonitoring() {
        // Listen for recipient changes
        listenForRecipientChanges();
        
        // Set up regular validation checks (every 3 seconds to be less resource-intensive)
        autoValidationInterval = setInterval(validateEmailsBackground, 3000);
        
        // Show minimal background UI
        updateStatusIndicator("success", "Monitoring for advanced issues");
    }
    
    function listenForRecipientChanges() {
        // Use event-based validation when recipients change
        currentItem.to.addHandlerAsync(Office.EventType.RecipientsChanged, debounce(validateEmailsBackground, 500));
        currentItem.cc.addHandlerAsync(Office.EventType.RecipientsChanged, debounce(validateEmailsBackground, 500));
        currentItem.bcc.addHandlerAsync(Office.EventType.RecipientsChanged, debounce(validateEmailsBackground, 500));
    }
    
    // Debounce helper to improve performance
    function debounce(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                timeout = null;
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    }

    function updateStatusIndicator(status, message) {
        const indicator = document.getElementById("minimal-status-indicator");
        const text = document.getElementById("minimal-status-text");
        
        // Update indicator color
        indicator.className = "status-indicator status-" + status;
        
        // Update text
        text.textContent = message;
    }

    async function validateEmailsBackground() {
        // Prevent concurrent validations
        if (isValidating) return;
        isValidating = true;

        try {
            const clientDomain = lastClientDomain;
            if (!clientDomain) {
                updateStatusIndicator("warning", "No client selected");
                isValidating = false;
                return;
            }

            // Get all recipients
            const recipients = await getAllRecipients();
            
            // Perform advanced validation (beyond what Outlook already checks)
            const validationResults = performAdvancedValidation(recipients, clientDomain);
            
            // Check if we need to show warnings for issues Outlook doesn't detect
            if (validationResults.shouldShowWarning) {
                // Show warning UI only if we detected issues that Outlook doesn't catch
                showWarningUI(validationResults.results, clientDomain);
                updateStatusIndicator("warning", "Advanced security warning");
            } else {
                // Everything looks good
                hideWarningUI();
                updateStatusIndicator("success", "No advanced issues detected");
            }
        } catch (error) {
            console.error("Validation error:", error);
            updateStatusIndicator("error", "Validation error");
        } finally {
            isValidating = false;
        }
    }

    async function getAllRecipients() {
        // Get all recipients from To, CC, and BCC fields
        const allRecipients = [];
        
        // Get To recipients
        return new Promise((resolve, reject) => {
            currentItem.to.getAsync(function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    reject(new Error("Failed to get recipients"));
                    return;
                }

                const toRecipients = asyncResult.value.map(recipient => recipient.emailAddress);
                allRecipients.push(...toRecipients);

                // Get CC recipients
                currentItem.cc.getAsync(function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        reject(new Error("Failed to get CC recipients"));
                        return;
                    }

                    const ccRecipients = asyncResult.value.map(recipient => recipient.emailAddress);
                    allRecipients.push(...ccRecipients);

                    // Get BCC recipients
                    currentItem.bcc.getAsync(function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                            reject(new Error("Failed to get BCC recipients"));
                            return;
                        }

                        const bccRecipients = asyncResult.value.map(recipient => recipient.emailAddress);
                        allRecipients.push(...bccRecipients);
                        resolve(allRecipients);
                    });
                });
            });
        });
    }

    function performAdvancedValidation(recipients, clientDomain) {
        const results = {
            domainTypos: [],        // Potential typos in domains
            multipleDomains: false, // Multiple external domains
            allExternalDomains: new Set()
        };
        
        // Count external domains (to check for multiple external domains)
        const externalDomains = new Set();
        
        recipients.forEach(email => {
            // Skip empty emails
            if (!email || email.trim() === '') return;
            
            // Check cache first
            if (emailValidationCache.has(email)) {
                const cachedResult = emailValidationCache.get(email);
                if (cachedResult.isDomainTypo) {
                    results.domainTypos.push(email);
                }
                if (cachedResult.domain && cachedResult.domain !== userDomain) {
                    externalDomains.add(cachedResult.domain);
                    results.allExternalDomains.add(cachedResult.domain);
                }
                return;
            }
            
            const domain = email.split('@')[1]?.toLowerCase();
            if (!domain) return;
            
            // Store validation result in cache
            const validationResult = {
                domain: domain,
                isDomainTypo: false
            };
            
            // Check for domain typos
            if (domain !== userDomain) {
                externalDomains.add(domain);
                results.allExternalDomains.add(domain);
                
                // Check for potential typos in common domains
                const possibleTypo = checkForDomainTypos(domain);
                if (possibleTypo) {
                    results.domainTypos.push(email);
                    validationResult.isDomainTypo = true;
                }
            }
            
            // Cache the result
            emailValidationCache.set(email, validationResult);
        });
        
        // Check for multiple external domains
        results.multipleDomains = externalDomains.size > 1;
        
        // Determine if we should show warnings
        // Only show warning if we have advanced issues that Outlook doesn't catch
        const shouldShowWarning = results.domainTypos.length > 0 || results.multipleDomains;
        
        return {
            results: results,
            shouldShowWarning: shouldShowWarning
        };
    }
    
    function checkForDomainTypos(domain) {
        // Don't check internal domain
        if (domain === userDomain) return false;
        
        // For common domains, check for potential typos
        for (const commonDomain of COMMON_DOMAINS) {
            // Skip exact matches
            if (domain === commonDomain) return false;
            
            // Check for similar domain (potential typo)
            if (isTypoLikely(domain, commonDomain)) {
                return true;
            }
        }
        
        return false;
    }
    
    function isTypoLikely(domain, knownDomain) {
        // Simple Levenshtein distance check for similar domains
        if (levenshteinDistance(domain, knownDomain) <= 2) {
            return true;
        }
        
        // Check for common typos like transposed letters
        if (domain.length === knownDomain.length) {
            let differences = 0;
            for (let i = 0; i < domain.length; i++) {
                if (domain[i] !== knownDomain[i]) {
                    differences++;
                }
            }
            if (differences <= 2) return true;
        }
        
        // Check for missing/extra character
        if (Math.abs(domain.length - knownDomain.length) === 1) {
            if (domain.includes(knownDomain) || knownDomain.includes(domain)) {
                return true;
            }
        }
        
        return false;
    }
    
    function levenshteinDistance(a, b) {
        // Simple implementation of Levenshtein distance for string similarity
        if (a.length === 0) return b.length;
        if (b.length === 0) return a.length;
        
        const matrix = [];
        
        // Initialize matrix
        for (let i = 0; i <= b.length; i++) {
            matrix[i] = [i];
        }
        for (let j = 0; j <= a.length; j++) {
            matrix[0][j] = j;
        }
        
        // Fill matrix
        for (let i = 1; i <= b.length; i++) {
            for (let j = 1; j <= a.length; j++) {
                if (b.charAt(i - 1) === a.charAt(j - 1)) {
                    matrix[i][j] = matrix[i - 1][j - 1];
                } else {
                    matrix[i][j] = Math.min(
                        matrix[i - 1][j - 1] + 1, // substitution
                        matrix[i][j - 1] + 1,     // insertion
                        matrix[i - 1][j] + 1      // deletion
                    );
                }
            }
        }
        
        return matrix[b.length][a.length];
    }

    function showWarningUI(results, clientDomain) {
        // Hide minimal UI
        document.getElementById("minimal-ui").style.display = "none";
        
        // Show warning UI
        document.getElementById("warning-ui").style.display = "block";
        
        // Populate results
        displayAdvancedWarnings(results, clientDomain);
    }
    
    function hideWarningUI() {
        // Hide warning UI
        document.getElementById("warning-ui").style.display = "none";
        
        // Show minimal UI
        document.getElementById("minimal-ui").style.display = "flex";
        
        // Clear results
        document.getElementById("validation-results").innerHTML = "";
    }
    
    function proceedAnyway() {
        // Log the action
        logAction("proceeded", lastClientDomain, []);
        
        // Hide warning UI
        hideWarningUI();
        
        // Update status
        updateStatusIndicator("info", "Proceeding with send");
        
        // After a delay, go back to monitoring
        setTimeout(() => {
            updateStatusIndicator("success", "Monitoring for advanced issues");
        }, 2000);
    }

    function displayAdvancedWarnings(results, clientDomain) {
        const resultsElement = document.getElementById("validation-results");
        resultsElement.innerHTML = "";

        const warningDiv = document.createElement("div");
        warningDiv.className = "validation-warning";
        
        const warningList = document.createElement("div");
        warningList.className = "warning-list";
        
        // Show domain typo warnings
        if (results.domainTypos.length > 0) {
            const typoHeader = document.createElement("div");
            typoHeader.className = "warning-header";
            typoHeader.innerHTML = "⚠️ <b>Possible typos in email domains:</b>";
            warningList.appendChild(typoHeader);
            
            results.domainTypos.forEach(email => {
                const emailSpan = document.createElement("div");
                emailSpan.className = "email-warning";
                
                const domain = email.split('@')[1];
                const username = email.split('@')[0];
                
                emailSpan.innerHTML = `${username}@<span class="highlight-warning">${domain}</span>`;
                warningList.appendChild(emailSpan);
            });
        }
        
        // Show multiple external domains warning
        if (results.multipleDomains) {
            const domainHeader = document.createElement("div");
            domainHeader.className = "warning-header";
            domainHeader.innerHTML = "⚠️ <b>Multiple external domains detected:</b>";
            warningList.appendChild(domainHeader);
            
            const domainList = document.createElement("div");
            domainList.className = "domain-list";
            
            // List each external domain
            results.allExternalDomains.forEach(domain => {
                if (domain !== userDomain) {
                    const domainSpan = document.createElement("div");
                    domainSpan.className = "domain-warning";
                    domainSpan.textContent = domain;
                    domainList.appendChild(domainSpan);
                }
            });
            
            warningList.appendChild(domainList);
        }
        
        // Add explanation
        const explanationText = document.createElement("div");
        explanationText.className = "explanation-text";
        explanationText.innerHTML = "<i>These are advanced issues that Outlook's native protection might not detect.</i>";
        warningList.appendChild(explanationText);
        
        warningDiv.appendChild(warningList);
        resultsElement.appendChild(warningDiv);
    }

    function logAction(action, clientDomain, issues) {
        const logEntry = {
            timestamp: new Date().toISOString(),
            user: Office.context.mailbox.userProfile.emailAddress,
            action: action,
            clientDomain: clientDomain,
            issues: issues,
            subject: currentItem.subject.slice(0, 100)
        };

        console.log("MailVolt Advanced Security Log:", logEntry);
    }
})(); 