# Ryno-Tec Business Hub (Microsoft 365)

The hub for Ryno-Tec Consulting and Lesego Laka Projects: dashboard, Outlook mail, mail archive,
tenders, clients, pipeline, projects, tasks, timesheets, quotes, invoices, contracts and the
compliance pack. It runs at **https://hub.ryno-tec.co.za**.

- **Sign-in:** Microsoft 365 (Entra ID app *Ryno-Tec Hub*). Only people assigned to the app in
  Entra > Enterprise applications > Ryno-Tec Hub > Users and groups can sign in.
- **Data:** SharePoint lists `Hub_<section>` on https://rynotecconsulting.sharepoint.com/sites/rynotechub.
  Each record is one list item; the full record is the `Data` column (JSON).
- **Documents:** the site's *Compliance* library.
- **Mail:** the signed-in user's Outlook mailbox through Microsoft Graph. The Exchange transport rule
  adds the email signature, so the hub doesn't.

Nothing in this repository is secret. A browser app has no client secret; access is controlled by
Entra sign-in and SharePoint permissions.

## Files

| File | What it is |
|---|---|
| `index.html` | The hub page. **Generated**, don't edit by hand. |
| `src/hub-artifact.html` | The hub as developed (Claude artifact version). |
| `tools/build.py` | Builds `index.html` from `src/hub-artifact.html` with the Microsoft 365 changes. |
| `m365.js` | Microsoft 365 runtime: sign-in, SharePoint lists, Outlook mail, document library. |
| `hub-config.js` | Tenant, app and SharePoint IDs from `Setup-RynoTecHub.ps1` / `Import-RynoTecHub.ps1`. |
| `vendor/msal-browser-3.28.1.min.js` | Microsoft's sign-in library (MSAL), kept here so the hub doesn't depend on a CDN. |
| `CNAME` | GitHub Pages custom domain. |

## Updating the hub

```
# after saving the new artifact HTML as src/hub-artifact.html
python3 tools/build.py
git add -A && git commit -m "Update hub" && git push
```

## Setup (done once)

1. `Setup-RynoTecHub.ps1`: app registration, SharePoint site, lists, Compliance library.
2. `Import-RynoTecHub.ps1`: copies records and documents into SharePoint, adds missing lists.
3. GitHub: repository **public**, Settings > Pages > Deploy from branch `main`, folder `/ (root)`,
   custom domain `hub.ryno-tec.co.za`, Enforce HTTPS.
4. DNS at Afrihost: `CNAME hub -> ryno-tec.github.io.`
