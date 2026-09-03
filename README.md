# UNYKORN BLACK

Autonomous CyberTaxi enterprise architecture for Atlanta — fleet operations from 5655 Peachtree Pkwy, Norcross, GA, commercial insurance underwriting, UUPS revenue splits, and $UBK RWA tokenization.

## Live
- Production: https://black.unykorn.org/
- Alias: https://black.unykorn.ai/
- Cloudflare Pages project: `black-unykorn`

## Product
- Norcross depot staging, live fleet radar, CyberCab / Model Y / CyberVan tiers
- Interactive insurance, split-ledger, and $UBK subscription calculators
- BitGo custody narrative + Atlanta commercial lender directory
- Voice tours (Atlas, Orion, Eve) and institutional document vault

## Deploy
Cloudflare Pages project `black-unykorn` (direct upload / wrangler).

```bash
npx wrangler pages deploy . --project-name black-unykorn --branch main
```

Required: `CLOUDFLARE_API_TOKEN`, `CLOUDFLARE_ACCOUNT_ID`

GitHub Actions workflow in `.github/workflows/deploy-cloudflare-pages.yml` also publishes on push to `main` when those secrets are set.

## Local preview
Serve the repo root as a static site (any static file server). Open `index.html`.

## Brand + Legal
UnyKorn LLC. Independent service. No official event affiliation implied without written authorization.
