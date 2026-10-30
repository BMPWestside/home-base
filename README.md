# Home Base

Family shopping lists, to-dos and trackers — a home-screen web app (PWA) on GitHub Pages, synced between phones through a small Google Apps Script web app that saves to your Google Drive.

## Sync backend (one-time, ~5 minutes)
1. <https://script.google.com> → **New project** → name it `Home Base sync`.
2. Replace `Code.gs` with [`apps-script/Code.gs`](apps-script/Code.gs) and set your own `HOUSE_KEY`.
3. **Deploy → New deployment → Web app** · Execute as **Me** · Who has access **Anyone** → Deploy → authorise.
4. Copy the **Web app URL** (ends in `/exec`).
5. Open the app → welcome screen (or ⚙︎ → Sync) → paste the URL and key → **Connect**.
6. **Send invite link** to family. They open it → Share → **Add to Home Screen**.

The key is never stored in this repo — only in your Apps Script and on your phones.
