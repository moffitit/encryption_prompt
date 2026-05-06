# Encrypt Email – Outlook Add-in
## Microsoft Purview OME Prompt

A minimal Outlook add-in (Classic · New · Web) that shows a simple **Yes / No** prompt before sending:

> "Does this email need to be encrypted?"

If the user clicks **Yes**, the add-in applies Microsoft Purview OME encryption via a sensitivity label (New Outlook / OWA) or an Exchange transport rule trigger (Classic Outlook fallback).

---

## File structure

```
encrypt-prompt-addin/
├── manifest.xml       ← Office add-in manifest (sideload or deploy via M365 admin)
├── taskpane.html      ← The Yes/No prompt UI
├── commands.html      ← Background function file (OnSend event handler)
├── assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-80.png
│   └── icon-128.png
└── README.md
```

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Microsoft 365 tenant | E3/E5 or OME add-on licence |
| Exchange Online | Transport rules available |
| Microsoft Purview | Information Protection configured |
| HTTPS web host | Azure Static Web Apps, SharePoint, or any HTTPS host |

---

## Step 1 – Create your sensitivity label (for New Outlook / OWA)

1. Go to **Microsoft Purview compliance portal** → **Information protection** → **Labels**
2. Create (or identify) a label with **Encryption** set to **Apply encryption** using **Microsoft Purview Message Encryption** (OME)
3. Copy the label **GUID** from the label properties page
4. Open `taskpane.html` and replace:
   ```js
   const ENCRYPT_LABEL_GUID = "YOUR-PURVIEW-LABEL-GUID-HERE";
   ```
   with your actual GUID, e.g.:
   ```js
   const ENCRYPT_LABEL_GUID = "ab12cd34-ef56-7890-abcd-ef1234567890";
   ```

---

## Step 2 – Create the Exchange transport rule (Classic Outlook fallback)

The Classic Outlook path sets a custom internet header `X-Encrypt-This: true`.  
You need an Exchange Online transport rule to detect it and apply OME.

Run in **Exchange Online PowerShell**:

```powershell
Connect-ExchangeOnline

New-TransportRule `
  -Name "Apply OME when X-Encrypt-This header is set" `
  -HeaderContainsMessageHeader "X-Encrypt-This" `
  -HeaderContainsWords "true" `
  -ApplyOME $true `
  -Comments "Applied by Encrypt Email Outlook add-in"
```

> **Note:** `ApplyOME $true` applies the default OME template.  
> To use a custom Purview template, use `-ApplyRightsProtectionTemplate` instead.

---

## Step 3 – Host the add-in files

Host `taskpane.html` and `commands.html` on an **HTTPS** server.

**Option A – Azure Static Web Apps (recommended)**
```bash
az staticwebapp create --name encrypt-addin --resource-group <rg> --source ./encrypt-prompt-addin
```

**Option B – SharePoint App Catalog**  
Upload files to a SharePoint document library that is served over HTTPS.

**Option C – Any web host**  
Upload files to your server. Ensure CORS is open for `https://appsforoffice.microsoft.com`.

Update **all three occurrences** of `https://YOUR-DOMAIN` in `manifest.xml` with your actual hosting URL.

---

## Step 4 – Deploy the manifest

### Option A – Microsoft 365 Admin Center (recommended for org-wide rollout)
1. Go to **M365 Admin Center** → **Settings** → **Integrated apps**
2. Click **Upload custom app** → upload `manifest.xml`
3. Assign to users or groups

### Option B – Sideload for testing
1. In **Outlook on the web**: Settings → Manage add-ins → upload `manifest.xml`
2. In **Classic Outlook** (Windows): File → Manage Add-ins → Add from file

---

## How it works

```
User opens compose window
         │
         ▼
   Clicks "Encrypt Email" ribbon button
         │
         ▼
   Task pane opens: "Does this email need to be encrypted?"
         │
    ┌────┴────┐
    │         │
   YES        NO
    │         │
    ▼         ▼
Try sensitivity   Close pane,
label API (1.7+)  send normally
    │
    ├─ Success → label applied, Exchange enforces OME
    │
    └─ Fallback → set X-Encrypt-This: true header
                  Exchange transport rule applies OME
```

---

## Customisation

### Block send unless the user answers the prompt
In `commands.html`, uncomment the block inside `onItemSend()`. This will prevent send if neither Yes nor No was clicked (i.e. the task pane was never opened).

### Change the OME template
In the transport rule, replace `-ApplyOME $true` with:
```powershell
-ApplyRightsProtectionTemplate "Encrypt"   # or your custom template name
```

### Change the prompt text
Edit the `<h1>` and `<p>` in `taskpane.html`.

---

## Supported clients

| Client | Mechanism |
|---|---|
| **Outlook on the Web (OWA)** | Sensitivity label API (req. set 1.7) |
| **New Outlook (Windows/Mac)** | Sensitivity label API (req. set 1.7) |
| **Classic Outlook (Windows)** | Internet header + Exchange transport rule |
| **Classic Outlook (Mac)** | Internet header + Exchange transport rule |

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| Ribbon button doesn't appear | Ensure manifest is deployed and `VersionOverrides` `AppDomain` matches your host |
| "Could not apply encryption" error | Check browser console; verify label GUID or header API availability |
| OME not applied on delivery | Verify the Exchange transport rule is active and the header value matches |
| Task pane blank | Ensure your host serves files over HTTPS with a valid TLS certificate |

---

## Licence
MIT – free to use and modify within your organisation.
