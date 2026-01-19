# Quick Reference - Outlook Mail Merge

## Start Development

```bash
cd "/Users/awindsor/Documents/Repositories/Outlook Mail Merge"
npm run dev
```

Then access at: `http://localhost:3000`

## Build for Deployment

```bash
npm run build
```

Output: `dist/bundle.js` and `dist/taskpane.html`

## Directory

```
/Users/awindsor/Documents/Repositories/Outlook Mail Merge/
```

## Template Syntax Examples

### Simple Replacement
```
Subject: Hello {{FirstName}}
Body: Dear {{FirstName}} {{LastName}},

This is a message for {{Company}}.

Best regards
```

### Conditional Logic
```
{{Status|active|You are an active member|You are inactive}}
{{Title|*|Manager|Is a Manager|Not a Manager}}
{{Country|^|US|American|International}}
```

## Data Format

### CSV Example
```
FirstName,LastName,Email,Company
John,Doe,john@example.com,Acme Corp
Jane,Smith,jane@example.com,Tech Inc
```

### Excel
- First row: column headers
- Data rows: recipient info
- Supports formulas and formatting

### JSON
```json
{"FirstName":"John","LastName":"Doe","Email":"john@example.com"}
{"FirstName":"Jane","LastName":"Smith","Email":"jane@example.com"}
```

## Common Issues & Fixes

| Issue | Fix |
|-------|-----|
| Port 3000 in use | `npm run dev -- --port 3001` |
| TypeScript errors | `npm run build` to see full errors |
| Variables not replacing | Check CSV headers match variable names exactly |
| Build fails | `rm -rf node_modules && npm install` |

## Dependencies Status

✅ **All Current** - Zero deprecated packages
- glob@10.3.10 (updated)
- rimraf@5.0.5 (updated)
- webpack@5.90.0
- All devDependencies latest stable

## Project Structure

```
src/
  ├── App.tsx                 # Main component
  ├── components/             # UI components
  │   ├── TemplateEditor
  │   ├── DataSourceSelector
  │   ├── PreviewPane
  │   └── SendPane
  ├── lib/
  │   └── TemplateEngine.ts   # Variable logic
  └── styles/                 # CSS files

public/
  └── taskpane.html          # Entry HTML

manifest.xml                 # Add-in config
package.json                 # Dependencies
webpack.config.js           # Build config
tsconfig.json              # TypeScript config
```

## Useful Commands

```bash
npm run dev          # Start development server
npm run build        # Production build
npm run lint         # Check code style
npm test            # Run tests
npm audit           # Security check
npm update          # Update packages
```

## Documentation Files

- **README.md** - Full feature documentation
- **MACOS_SETUP.md** - macOS installation guide
- **SETUP_COMPLETE.md** - Setup verification and status

---

**Last Updated:** January 19, 2026
**Status:** Production Ready ✅
