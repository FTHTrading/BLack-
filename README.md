# UNYKORN BLACK

Premium Event Mobility launch site and lender packet for Atlanta 2026.

## Live Product Intent
UNYKORN BLACK combines:
- Premium black-car fleet operations
- Sponsor route inventory
- Merchant referral monetization
- Event demand-driven dispatch
- TROPTIONS reporting alignment

## Included In This Repo
- `index.html`: full branded funder landing page
- `brand/`: logo and shield assets
- `fleet/`: SUV, sedan, van, shuttle concept visuals
- `downloads/`: lender-ready files (PDF, DOCX, CSV, ZIP)
- `docs/generate_docs.py`: script to regenerate downloadable docs

## Downloads
- `downloads/unykorn-black-lender-proposal.pdf`
- `downloads/unykorn-black-lender-proposal.docx`
- `downloads/unykorn-black-sponsor-rate-card.pdf`
- `downloads/unykorn-black-event-demand-calendar.csv`
- `downloads/unykorn-black-downloads.zip`

## Deploy Target
- Domain: `black.unykorn.org`
- Platform: Cloudflare Pages

## Auto Deploy (GitHub Actions)
This repo includes `.github/workflows/deploy-cloudflare-pages.yml`.

Set these GitHub repository secrets:
- `CLOUDFLARE_API_TOKEN`
- `CLOUDFLARE_ACCOUNT_ID`

Then push to `main` and the site will deploy to Cloudflare Pages project `black-unykorn`.

## Cloudflare Domain Binding
In Cloudflare Pages:
1. Open project `black-unykorn`
2. Go to Custom domains
3. Add `black.unykorn.org`
4. Confirm DNS record in zone `unykorn.org`

## Local Preview
Open `index.html` directly in browser, or serve via any static file server.

## Brand + Legal
Independent service. No official event affiliation implied without written authorization.
