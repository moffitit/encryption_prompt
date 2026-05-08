# Encryption Reminder – Outlook Add-in

Intercepts every Send action and asks:
**"Does this email need to be encrypted?"**

The user can choose:
- **Cancel** – stops the send so they can apply encryption first
- **Send Anyway** – proceeds normally

---

## Files

| File | Purpose |
|------|---------|
| `manifest.xml` | Add-in manifest (sideload this in Outlook / M365 Admin) |
| `commands.html` | Hidden function-file page loaded by Outlook |
| `commands.js` | Intercepts the ItemSend event, opens the dialog |
| `dialog.html` | The visible prompt shown to the user |

---

## Deployment steps

### 1 – Host the files
Upload `commands.html`, `commands.js`, and `dialog.html` to a web server
that is reachable by Outlook (must be **HTTPS**).

Example: `https://addin.yourcompany.com/`

### 2 – Update manifest.xml
Replace every occurrence of `https://YOUR-DOMAIN.com` with your actual URL.

```xml
<!-- three places to update: -->
<SourceLocation DefaultValue="https://addin.yourcompany.com/taskpane.html"/>
<bt:Url id="Commands.Url" DefaultValue="https://addin.yourcompany.com/commands.html"/>
<bt:Url id="Dialog.Url"   DefaultValue="https://addin.yourcompany.com/dialog.html"/>
```

### 3 – Sideload / deploy the manifest

**Individual testing (Outlook desktop)**
1. Open Outlook → File → Manage Add-ins (opens OWA)
2. Click **My add-ins** → **Add a custom add-in** → **Add from file…**
3. Select your updated `manifest.xml`

**Organisation-wide (Microsoft 365 Admin)**
1. Go to [admin.microsoft.com](https://admin.microsoft.com)
2. Settings → Integrated apps → Upload custom apps
3. Upload `manifest.xml` and assign to users/groups

---

## Requirements

- Microsoft 365 / Exchange Online  
- Outlook desktop (Windows/Mac) or Outlook on the Web  
- Mailbox requirement set **1.3** or later  
- The add-in host must be served over **HTTPS**

---

## Customisation tips

- Edit the prompt text in `dialog.html` (`<h1>` and `<p>` tags)
- Change button colours via the `.btn-cancel` / `.btn-send` CSS classes
- To auto-detect sensitive keywords and skip the prompt for normal mail,
  add subject/body inspection logic in `commands.js` before calling
  `displayDialogAsync`
