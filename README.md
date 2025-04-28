# MailVolt

An Outlook add-in for streamlining email distribution list security.

## Features

- Secure email distribution list management
- Project-based access control
- Simple and intuitive interface

## Setup

1. Clone the repository:
```bash
git clone https://github.com/szakacs-media/mailvolt.git
cd mailvolt
```

2. Install dependencies:
```bash
npm install
```

3. Start the development server:
```bash
npm start
```

4. Load the add-in in Outlook:
- Open Outlook on the web
- Go to Settings > Manage Add-ins
- Click "Add a custom add-in" > "Add from file"
- Select the manifest.xml file
- Allow insecure content (for development)

## Development

The add-in consists of:
- `taskpane.html`: The main interface
- `manifest.xml`: Add-in configuration
- `assets/`: Icons and other static files

## License

ISC 