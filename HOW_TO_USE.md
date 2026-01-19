# How to Use Outlook Mail Merge Add-in on macOS

## Step 1: Start the Development Server

Open Terminal and run:

```bash
cd "/Users/awindsor/Documents/Repositories/Outlook Mail Merge"
npm run dev
```

This starts a server at `http://localhost:3000` that will serve the add-in to Outlook.

## Step 2: Access the Add-in in Outlook

Unfortunately, macOS Outlook doesn't have a built-in UI for sideloading add-ins like Windows does. Here are your options:

### Option A: Web Access (Easiest for Testing)
1. Open Outlook on the web at https://outlook.office.com
2. Go to **Settings** → **Add-ins** → **Get add-ins**
3. Click **My add-ins** → **Upload My Add-in**
4. Upload the `manifest.xml` file from the project root
5. The add-in will appear in your compose window

### Option B: Desktop Outlook (Requires Registry/Plist Modification)

**For Outlook on macOS (more complex):**
1. Keep the dev server running (`npm run dev`)
2. The add-in will be available through Outlook's add-in store (when published)
3. Or modify Outlook's Registry/Plist to point to localhost (advanced)

### Option C: Use Outlook Web Access (Recommended for Development)

This is the fastest way to test:

1. **Keep dev server running**
   ```bash
   npm run dev
   ```

2. **Go to Outlook Web** at https://outlook.office.com

3. **Create a new email** by clicking **+ New message**

4. **Upload the manifest**:
   - Click the **+ Add-ins** button
   - Select **Get add-ins**
   - Choose **Upload My Add-in**
   - Select `/Users/awindsor/Documents/Repositories/Outlook Mail Merge/manifest.xml`

5. **The add-in will open in a pane** on the right side of the compose window

## Step 3: Use the Add-in (4-Step Workflow)

### Step 1: Create Your Email Template

1. In the **Template** tab, enter:
   - **Subject Line**: `Hello {{FirstName}}, we have an opportunity for {{Company}}`
   - **Body**: 
     ```
     Dear {{FirstName}} {{LastName}},
     
     We'd like to offer you a position at {{Company}}.
     
     Best regards
     ```

2. Click **+ Variable** buttons to insert variables easily

### Step 2: Load Your Recipient Data

1. Go to **Data Source** tab
2. Choose your data source:
   - **CSV File**: Upload a CSV with columns: FirstName, LastName, Email, Company
   - **Excel File**: Upload XLSX with same columns
   - **Manual Entry**: Paste JSON objects

**Example CSV:**
```
FirstName,LastName,Email,Company
John,Doe,john@example.com,Acme Corp
Jane,Smith,jane@example.com,Tech Inc
Bob,Johnson,bob@example.com,StartUp LLC
```

### Step 3: Preview Your Emails

1. Go to **Preview** tab
2. Click **← Previous** / **Next →** to browse through recipients
3. See exactly how each personalized email will look
4. Verify variables are correct before sending

### Step 4: Create Drafts

1. Go to **Send** tab
2. Review the summary (number of recipients, template)
3. Click **Create Email Drafts**
4. Wait for progress bar to complete
5. Drafts will appear in your Outlook **Drafts** folder

## Step 4: Review and Send

1. Go to Outlook **Drafts** folder
2. Open each draft to review
3. Make any edits if needed
4. Click **Send** when ready

## Template Variables Reference

### Simple Variables
```
{{FirstName}}    → John
{{LastName}}     → Doe
{{Email}}        → john@example.com
{{Company}}      → Acme Corp
```

### Conditional (if equals)
```
{{Status|active|Is Active|Is Inactive}}
If Status="active" → "Is Active", else "Is Inactive"
```

### Contains
```
{{Title|*|Manager|Management|Staff}}
If Title contains "Manager" → "Management", else "Staff"
```

### Starts With
```
{{Company|^|Tech|Tech Company|Other}}
If Company starts with "Tech" → "Tech Company", else "Other"
```

## Troubleshooting

### "Add-in won't load"
- Ensure dev server is running: `npm run dev`
- Check that `http://localhost:3000` is accessible
- Try uploading manifest.xml again

### "Variables show {{Variable}} instead of replacing"
- Check that CSV/Excel column headers exactly match variable names
- Variable names are case-sensitive
- Make sure data was loaded successfully in Step 2

### "Port 3000 already in use"
```bash
npm run dev -- --port 3001
```

### "Cannot upload manifest"
- Ensure manifest.xml is in the project root
- Check that manifest.xml is valid XML
- Try using Outlook Web instead of desktop

## Important Notes

⚠️ **Draft Review**: Always review drafts before sending! The add-in creates drafts for safety, not sending directly.

✅ **Local Processing**: All data stays on your computer - no cloud processing

📧 **One Draft Per Recipient**: Each recipient gets their own personalized draft

## Next: Deploy to Production

When you're ready to use this professionally:

1. **Build for production**:
   ```bash
   npm run build
   ```

2. **Deploy to a web server** (HTTPS required):
   - Update `sourceLocation` in `manifest.xml` with your server URL
   - Upload built files from `dist/` directory

3. **Submit to Microsoft AppSource**:
   - Package the add-in
   - Submit for Microsoft review
   - Users can install from Office Store

## Support

- **Full Documentation**: See `README.md` in project folder
- **macOS Setup Guide**: See `MACOS_SETUP.md`
- **Repository**: https://github.com/awindsor/outlook-mail-merge
