# Challan → Invoice Portal

Upload your RTO-style challan upload sheet (.xlsx) and generate editable, print-ready tax invoices.

## Deploy on GitHub Pages
1. Push this repo to GitHub (branch: `main`).
2. Go to **Settings → Pages** in the repo.
3. Under **Build and deployment → Source**, choose **GitHub Actions**.
4. Push any commit to `main` (or re-run the workflow) — the included workflow at
   `.github/workflows/deploy.yml` will build and publish the site automatically.
5. Your live link will be: `https://<your-username>.github.io/<repo-name>/`

## Local development
```
npm install
npm run dev
```
