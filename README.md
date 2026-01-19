# Outlook Mail Merge Add-in

A mail merge add-in for Outlook on macOS that lets you create and send personalized bulk emails from templates, similar to the Thunderbird Mail Merge extension.

## Features

- **Template Editor**: Create email templates with variable placeholders
- **Multiple Data Sources**: Support for CSV, Excel (XLSX), and manual entry
- **Variable Substitution**: Replace {{FirstName}}, {{Email}}, etc. in subject and body
- **Advanced Variables**: Support for conditionals and transformations
- **Live Preview**: See exactly how each email will look before sending
- **Draft Creation**: Generate personalized drafts in Outlook for review
- **Batch Operations**: Create multiple drafts at once

## Getting Started

### Prerequisites

- macOS with Outlook 2016 or newer
- Node.js 16+ and npm
- Microsoft Office Add-in development tools

### Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/outlook-mail-merge.git
cd outlook-mail-merge
```

2. Install dependencies:
```bash
npm install
```

3. Start the development server:
```bash
npm run dev
```

4. In a new terminal, sideload the add-in:
```bash
npx office-addin-debugging start manifest.xml
```

## Development

### Project Structure

```
src/
├── App.tsx                 # Main application component
├── index.tsx              # Entry point
├── index.css              # Global styles
├── components/            # React components
│   ├── TemplateEditor.tsx
│   ├── DataSourceSelector.tsx
│   ├── PreviewPane.tsx
│   └── SendPane.tsx
├── lib/
│   └── TemplateEngine.ts  # Variable substitution engine
└── styles/               # Component-specific styles

public/
└── taskpane.html        # HTML entry point for Office Add-in

manifest.xml            # Office Add-in manifest
package.json           # Dependencies and scripts
webpack.config.js      # Webpack build configuration
tsconfig.json         # TypeScript configuration
```

### Build

```bash
npm run build
```

### Testing

```bash
npm test
```

## Usage

### Step 1: Create Template

1. Write your email subject line and body
2. Use {{Variable}} placeholders for dynamic content
3. Common variables: {{FirstName}}, {{LastName}}, {{Email}}, {{Company}}, {{Title}}

### Step 2: Add Recipients

Choose from:
- **CSV File**: Upload a CSV with column headers matching your variables
- **Excel File**: Upload an XLSX spreadsheet
- **Manual Entry**: Paste JSON objects directly

Example CSV:
```
FirstName,LastName,Email,Company
John,Doe,john@example.com,Acme Corp
Jane,Smith,jane@example.com,Tech Inc
```

### Step 3: Preview

Navigate through recipients to preview how variables are substituted in each email.

### Step 4: Send

Click "Create Email Drafts" to:
1. Generate individual personalized emails
2. Create drafts in your Outlook Drafts folder
3. Review each one before sending

## Template Variables

### Basic Variables
```
{{FirstName}}  - Replace with First Name value
{{Email}}      - Replace with Email value
{{Company}}    - Replace with Company value
```

### Conditional Variables

**Equals:**
```
{{Status|active|Is Active|}}
If Status = "active", show "Is Active", otherwise empty
```

**Equals with Else:**
```
{{Status|active|Is Active|Is Inactive}}
If Status = "active", show "Is Active", else "Is Inactive"
```

**Contains:**
```
{{Title|*|Manager|Is Manager|Is Not Manager}}
If Title contains "Manager", show "Is Manager", else "Is Not Manager"
```

**Starts With:**
```
{{Company|^|Tech|Tech Company|Other Company}}
If Company starts with "Tech", show "Tech Company", else "Other Company"
```

## Data Source Formats

### CSV
- First row must contain column headers
- Headers are used as variable names
- Supports different character encodings
- Supports field delimiters: comma, semicolon, tab

### Excel (XLSX)
- First sheet is used by default
- First row must contain headers
- Supports formulas and formatting

### Manual Entry
JSON object format, one per line:
```json
{"FirstName":"John","LastName":"Doe","Email":"john@example.com"}
{"FirstName":"Jane","LastName":"Smith","Email":"jane@example.com"}
```

## Security & Privacy

- All processing happens locally on your computer
- No data is sent to external servers
- Template engine runs entirely in Outlook
- Works with Outlook's native security model

## Limitations

- Requires Outlook 16.16+ (2016 or newer)
- macOS only (currently)
- Maximum 5000 recipients per batch recommended
- Attachments not yet supported in v1.0

## Future Enhancements

- [ ] Attachment support with per-recipient customization
- [ ] Scheduled send support
- [ ] Contact/Address book integration
- [ ] Windows Outlook support
- [ ] Advanced filtering and sorting
- [ ] Template library and saving
- [ ] Send tracking and analytics

## Comparison with Thunderbird Mail Merge

| Feature | Outlook | Thunderbird |
|---------|---------|------------|
| CSV Support | ✓ | ✓ |
| Excel Support | ✓ | ✓ |
| JSON Support | ✓ | ✓ |
| Contacts Integration | ⏳ | ✓ |
| Variable Conditionals | ✓ | ✓ |
| Draft Review | ✓ | ✓ |
| Platform | macOS | Cross-platform |

## License

GPL-3.0 (Compatible with Thunderbird Mail Merge)

## Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Test thoroughly on macOS
4. Submit a pull request

## Support

For issues, questions, or feature requests, please open an issue on GitHub.

## Acknowledgments

Inspired by [Alexander Bergmann's Mail Merge extension for Thunderbird](https://addons.thunderbird.net/addon/mail-merge/)
