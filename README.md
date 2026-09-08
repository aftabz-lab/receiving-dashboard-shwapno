# Receiving Dashboard-Shwapno

A branded GitHub Pages wrapper for the live Shwapno Receiving Power BI report.

## What it provides

- Full-screen access to the existing two-page Power BI report.
- The latest Over Receiving and Under Receiving views remain interactive.
- Automatic report reload every 15 minutes.
- Immediate manual refresh, full-screen mode and direct Power BI access.
- Desktop and mobile-friendly dashboard shell.
- No copied data, API key, password or Power BI credential is stored in this repository.

## Publish on GitHub Pages

1. Sign in to the `aftabz-lab` GitHub account.
2. Create a new **Public** repository named exactly `receiving-dashboard-shwapno`.
3. Upload `index.html`, `styles.css`, `app.js` and `README.md` to the repository root.
4. Commit the files to the `main` branch.
5. Open **Settings → Pages**.
6. Under **Build and deployment**, choose **Deploy from a branch**.
7. Select branch **main**, folder **/(root)**, then click **Save**.
8. After GitHub Pages finishes publishing, open:
   `https://aftabz-lab.github.io/receiving-dashboard-shwapno/`

## Automatic updates

The site embeds the supplied Power BI Publish-to-web URL. When the Power BI dataset refreshes or the report is updated under the same published link, the GitHub dashboard shows the new version without uploading new website files. The page also reloads the embedded report every 15 minutes and when the browser returns after a missed refresh.

If Power BI generates a completely new Publish-to-web URL, replace the URL in the `src` and `data-report-url` attributes inside `index.html`.

## Add it to the main portal

Use the separately supplied portal `index.html` update in the `aftabz-lab.github.io` repository. It adds the dashboard name and link without changing the existing portal logic or styling.
