# Southcoast Daily Update – Outlook Add-in

A one-click button in **new Outlook** that fills a compose email with:
- **To:** trobinson@basesix.com
- **Cc:** rrowland@basesix.com, chriscervantes@btsdata.com
- **Subject:** `Southcoast - Daily Update M/D` (today's date auto-inserted)
- **Body:** the daily-update template, **prepended above your existing signature**

Because it works on the email Outlook already created, your signature is preserved automatically. No Graph permissions or admin consent are required — the add-in only touches the message you're composing.

## Files
| File | Purpose |
|------|---------|
| `manifest.xml` | The add-in definition (sideload / deploy this). |
| `commands.html` | Loads Office.js and the command script. |
| `commands.js` | The button logic (recipients, subject, template). |
| `assets/icon-*.png` | Ribbon button icons (16/32/80 px). |

## Hosting (GitHub Pages)
1. Push this folder to a **public** GitHub repo.
2. Repo **Settings → Pages → Build and deployment → Deploy from a branch → `main` / `(root)`** → Save.
3. Wait ~1 minute; files are served at `https://<username>.github.io/<repo>/`.
4. In `manifest.xml`, every `__BASEURL__` must equal that base URL **without** a trailing slash
   (e.g. `https://joshp.github.io/daily-update-addin`).

## Install for yourself (sideload, new Outlook)
1. New Outlook → **Settings (gear) → General → Manage add-ins** (opens the add-in dialog), or
   **More apps (···) → Add-ins → My add-ins → Add a custom add-in → Add from URL**.
2. Paste the URL to your hosted `manifest.xml`
   (`https://<username>.github.io/<repo>/manifest.xml`) and confirm.
3. Open **New mail** → look for the **Daily Update** button (BTS group) on the ribbon.

## Deploy to the whole team (you're tenant admin)
**Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps →**
provide the manifest URL, choose the users/groups, finish. The button then appears
for everyone you assigned.

## Changing recipients or the template
Edit `commands.js` (recipients near the top, template HTML below), commit, and push.
GitHub Pages updates within a minute. No need to reinstall the add-in.
