(function () {
    "use strict";

    let currentItem = null;
    let userDomain = null;

    Office.onReady(function () {
        if (Office.context.mailbox) {
            document.getElementById("app-body").style.display = "block";
            document.getElementById("sideload-msg").style.display = "none";

            // Get the user's email domain
            const userEmail = Office.context.mailbox.userProfile.emailAddress;
            userDomain = userEmail.split('@')[1];

            // Add event handler for the validate button
            document.getElementById("validate-button").addEventListener("click", validateEmails);

            // Store the current item
            currentItem = Office.context.mailbox.item;

            // Add event listener for client selection
            document.getElementById("client-select").addEventListener("change", function() {
                document.getElementById("validation-status").textContent = "";
                document.getElementById("validation-results").innerHTML = "";
            });

            // Monitor email input changes
            monitorEmailInput();
        }
    });

    function monitorEmailInput() {
        // Get the current item's body
        currentItem.body.getAsync("html", function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const body = asyncResult.value;
                // Add event listener for body changes
                currentItem.addHandlerAsync(Office.EventType.ItemChanged, function() {
                    validateEmails();
                });
            }
        });
    }

    function validateEmails() {
        const clientDomain = document.getElementById("client-select").value;
        
        if (!clientDomain) {
            showStatus("Select a client organization", "warning");
            return;
        }

        // Get all recipients
        const recipients = [];
        
        // Get To recipients
        currentItem.to.getAsync(function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                showStatus("Failed to get recipients", "error");
                return;
            }

            const toRecipients = asyncResult.value.map(recipient => recipient.emailAddress);
            recipients.push(...toRecipients);

            // Get CC recipients
            currentItem.cc.getAsync(function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showStatus("Failed to get CC recipients", "error");
                    return;
                }

                const ccRecipients = asyncResult.value.map(recipient => recipient.emailAddress);
                recipients.push(...ccRecipients);

                // Get BCC recipients
                currentItem.bcc.getAsync(function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        showStatus("Failed to get BCC recipients", "error");
                        return;
                    }

                    const bccRecipients = asyncResult.value.map(recipient => recipient.emailAddress);
                    recipients.push(...bccRecipients);
                    validateRecipientsList(recipients, clientDomain);
                });
            });
        });
    }

    function validateRecipientsList(recipients, clientDomain) {
        const results = {
            internal: [], // Same domain as sender
            valid: [],    // Valid external domains
            invalid: []   // Potentially incorrect domains
        };

        recipients.forEach(email => {
            const domain = email.split('@')[1]?.toLowerCase();
            if (!domain) return;

            if (domain === userDomain) {
                results.internal.push(email);
            } else if (isValidDomain(domain)) {
                results.valid.push(email);
            } else {
                results.invalid.push(email);
            }
        });

        const resultsElement = document.getElementById("validation-results");
        resultsElement.innerHTML = "";

        if (results.invalid.length > 0) {
            showStatus(`Found ${results.invalid.length} potentially incorrect email(s)`, "warning");

            const warningDiv = document.createElement("div");
            warningDiv.className = "validation-warning";
            
            const recipientsList = document.createElement("div");
            recipientsList.className = "recipients-list";
            
            // Add internal emails (grey)
            results.internal.forEach(email => {
                const emailSpan = document.createElement("span");
                emailSpan.className = "email-internal";
                emailSpan.textContent = email;
                recipientsList.appendChild(emailSpan);
            });

            // Add valid external emails (green)
            results.valid.forEach(email => {
                const emailSpan = document.createElement("span");
                emailSpan.className = "email-valid";
                emailSpan.textContent = email;
                recipientsList.appendChild(emailSpan);
            });

            // Add invalid emails (red)
            results.invalid.forEach(email => {
                const emailSpan = document.createElement("span");
                emailSpan.className = "email-invalid";
                emailSpan.textContent = email;
                recipientsList.appendChild(emailSpan);
            });
            
            const buttonContainer = document.createElement("div");
            buttonContainer.className = "button-container";

            const goBackButton = document.createElement("button");
            goBackButton.className = "ms-Button ms-Button--default";
            goBackButton.innerHTML = "Go Back";
            goBackButton.onclick = () => {
                logAction("went_back", clientDomain, results.invalid);
                resultsElement.innerHTML = "";
                showStatus("", "none");
            };

            const proceedButton = document.createElement("button");
            proceedButton.className = "ms-Button ms-Button--primary";
            proceedButton.innerHTML = "Proceed Anyway";
            proceedButton.onclick = () => {
                if (confirm(`Warning: The following email(s) appear to be incorrect:\n\n${results.invalid.join('\n')}\n\nDo you want to proceed anyway?`)) {
                    logAction("proceeded", clientDomain, results.invalid);
                    resultsElement.innerHTML = "";
                    showStatus("Proceeding with send", "info");
                }
            };

            warningDiv.appendChild(recipientsList);
            buttonContainer.appendChild(goBackButton);
            buttonContainer.appendChild(proceedButton);
            warningDiv.appendChild(buttonContainer);
            resultsElement.appendChild(warningDiv);
        } else {
            showStatus("All recipients verified", "success");
        }
    }

    function isValidDomain(domain) {
        // Basic domain validation - you can enhance this with more sophisticated checks
        const commonDomains = ['gmail.com', 'yahoo.com', 'outlook.com', 'hotmail.com', 'aol.com'];
        return commonDomains.includes(domain) || domain.includes('.');
    }

    function showStatus(message, type) {
        const statusElement = document.getElementById("validation-status");
        statusElement.textContent = message;
        statusElement.className = "status-text status-" + type;
    }

    function logAction(action, clientDomain, invalidRecipients) {
        const logEntry = {
            timestamp: new Date().toISOString(),
            user: Office.context.mailbox.userProfile.emailAddress,
            action: action,
            clientDomain: clientDomain,
            invalidRecipients: invalidRecipients,
            subject: currentItem.subject.slice(0, 100)
        };

        console.log("MailVolt Log:", logEntry);
    }
})(); 