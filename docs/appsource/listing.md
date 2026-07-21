# NovoMCP for Word — AppSource Listing Copy

All fields keyed to Partner Center "Offer setup → Properties / Offer listing" form.
Source of truth for the public marketplace submission. When AppSource verification
clears, paste these into the form as-is.

---

## Publisher

- **Publisher display name:** NovoMCP
- **Provider name (manifest):** NovoMCP
- **Support URL:** https://novomcp.com/support
- **Privacy URL:** https://novomcp.com/privacy
- **Terms URL:** https://novomcp.com/terms

## Add-in name

NovoMCP — Computational Chemistry for Word

(60 char limit; this is 42)

## Short description (100 char limit)

Look up ADMET, FAVES compliance, and quantum properties for any SMILES in your manuscript.

(96 chars)

## Long description (4,000 char limit; no markdown — Partner Center renders plain text)

NovoMCP brings the computational chemistry engine behind drug discovery and materials science into Microsoft Word. Highlight any SMILES string in your manuscript and the NovoMCP taskpane returns a full molecular profile — chemical properties, 31 ADMET predictions, FAVES regulatory compliance, structural alerts, similarity search across 122 million molecules, and on-demand quantum calculations — without leaving the document.

WHO IT IS FOR

Researchers, medicinal chemists, regulatory writers, and patent attorneys who draft documents that reference small molecules. If your workflow involves typing SMILES into a Word document and then opening five other tools to characterize each one, NovoMCP collapses the loop.

WHAT YOU GET

— Profile tab: molecular weight, logP, TPSA, rotatable bonds, H-bond donors and acceptors, QED v2, Lipinski compliance, BOILED-Egg classification.

— ADMET tab: 31 machine-learning models including CYP3A4/2D6/2C9 substrates and inhibitors, hERG, AMES mutagenicity, hepatotoxicity, clearance, half-life, solubility, plasma protein binding, P-glycoprotein, and blood-brain barrier permeability. Five endpoints hold state-of-the-art results on the TDC benchmark (NovoExpert-2).

— Compliance tab: FAVES v4.1 regulatory screening. Pass/fail verdict plus the specific structural alerts and toxicophores triggered, with PAINS filter, Brenk filter, and 24 additional medicinal chemistry filters.

— Similar tab: vector search across 122 million PubChem molecules. Tanimoto threshold adjustable from 0.6 to 0.95. Returns scored hits with one-click profile lookups.

— Advanced tab (compute tier only): on-demand quantum chemistry — pKa prediction, aqueous solubility with temperature dependence, bond dissociation energies, HOMO/LUMO frontier orbital energies. Backed by xTB, CREST, Psi4, and AIMNet2 NNP services.

— Document scan: one-click table scanner finds every SMILES in your tables and verifies the property columns against fresh predictions. Catches transcription errors and stale values before submission.

— Insert results back into Word: every result panel has an "Insert as Word table" button. The reference and the prediction live in the same document.

CLINICAL OUTCOMES (PREVIEW)

NovoExpert-3 predicts Phase II clinical success probability for compounds in cardiovascular, GI, and mainstream therapeutic areas. The model is validated for these domains and the taskpane flags it explicitly when used outside its training distribution.

HOW IT CONNECTS

You bring your own NovoMCP API key. Free Core keys (prefix nmcp_) unlock profiles, ADMET, compliance, and similarity. Compute keys (prefix ncmcp_) on Scale or Enterprise tiers unlock the Advanced tab. Both are issued at https://novomcp.com after account creation. No document content is sent to NovoMCP servers — only the SMILES strings you choose to analyze, namespaced with the X-Novo-Surface: word-addin-v1 header for audit isolation.

The same engine powers NovoMCP's Chrome extension, our MCP server in Claude and Gemini, and the Novo dashboard. Funnel state propagates across surfaces — a profile you ran in Word is visible in Claude immediately, and vice versa.

PRIVACY

API keys are stored in the add-in's sandboxed localStorage per Office's add-in security model. No document content beyond explicit user-selected SMILES leaves your machine. No PII beyond the account email registered with NovoMCP. No cross-document tracking. Full policy at https://novomcp.com/privacy.

GETTING STARTED

After install, click NovoMCP on the Home ribbon → paste your API key → highlight any SMILES in your document and the Profile tab populates immediately.

Pricing and account creation: https://novomcp.com/pricing

(approx 3,400 chars — fits the 4,000 limit with room for minor edits)

## Search keywords (max 7)

1. SMILES
2. ADMET
3. cheminformatics
4. drug discovery
5. molecular properties
6. computational chemistry
7. medicinal chemistry

## Category (Partner Center taxonomy)

- **Primary:** Productivity
- **Secondary:** Education

(No "Science" or "Research" category exists in Office Store taxonomy — Productivity is the closest fit; many science add-ins ship under it.)

## Industries (max 2)

1. Healthcare and life sciences
2. Education

## Office host

- Word (web, Windows desktop, Mac desktop, iPad)

## Office versions supported (from manifest)

- Office on the web
- Office 2016 and later (desktop)
- Office on iPad

## Languages

- English (United States) — primary
- Future locales added per-release; no en-GB / fr-FR / de-DE forks at launch.

## Logos and icons

| Asset | Size | Path in repo | Status |
|---|---|---|---|
| Store logo (square) | 300 × 300 PNG, transparent or solid background | `assets/store/logo-300.png` | ✅ shipping |
| Store hero (optional) | 815 × 290 PNG | `assets/store/hero-815x290.png` | TODO — marketing to draft |
| Ribbon icon 16 | 16 × 16 PNG | `assets/icon-16.png` | ✅ shipping |
| Ribbon icon 32 | 32 × 32 PNG | `assets/icon-32.png` | ✅ shipping |
| Ribbon icon 80 | 80 × 80 PNG | `assets/icon-80.png` | ✅ shipping |

## Screenshots (1366 × 768 PNG, 1–10 allowed, recommend 5)

Capture in Word web at 100% browser zoom; the Office UI chrome must be visible so reviewers can see the add-in is genuinely running in Word, not a mockup.

1. **`01-profile-tab.png`** — Word document open with a paragraph containing "aspirin (CC(=O)OC1=CC=CC=C1C(=O)O) shows…", SMILES highlighted, NovoMCP taskpane on the right showing the Profile tab fully populated (MW, logP, TPSA, QED, Lipinski badges).
2. **`02-admet-tab.png`** — Same document, taskpane scrolled to ADMET tab, with at least 8 ADMET endpoints visible including CYP3A4 substrate, hERG, hepatotoxicity. Highlight at least one SOTA badge if rendered.
3. **`03-compliance-tab.png`** — A compound that triggers a non-clean FAVES verdict (e.g., a known PAINS-flagged scaffold). Taskpane shows the Compliance tab with the verdict, the specific alerts triggered, and the recommendation list.
4. **`04-similar-tab.png`** — Similar tab showing 5–10 Tanimoto hits for a parent SMILES, with similarity scores visible. Threshold slider at 0.85.
5. **`05-scan-document.png`** — A two-table Word document where the scanner has found SMILES in both tables. Scan panel shows ✓ green for matched property columns and ⚠ yellow for one mismatched value.

Optional 6–10 if we want them:

6. **`06-advanced-pka.png`** — Advanced tab with a successful pKa prediction inserted as a Word table.
7. **`07-frontier-orbitals.png`** — Advanced tab with HOMO/LUMO energies for a chromophore plus the orbital images.

Screenshot capture script: `scripts/capture-screenshots.md` (to be written).

## Test accounts for Microsoft reviewers

Two demo keys minted specifically for AppSource review. Both must remain valid for the full 12-month listing lifecycle; rotate via dashboard-aggregator if compromised. Document them in Partner Center → Test accounts. Do NOT commit the actual key values to git.

| Field | Value |
|---|---|
| Demo Core account email | appsource-review-core@novomcp.com |
| Demo Core API key | nmcp_… (mint at https://app.novomcp.com → Admin → API Keys → "AppSource Review" label) |
| Demo Compute account email | appsource-review-compute@novomcp.com |
| Demo Compute API key | ncmcp_… (mint at https://app.novomcp.com → Admin → API Keys → "AppSource Review Compute" label, Scale tier) |

Both keys get an unmetered credit allowance (10,000 cr/mo) flagged `is_review_key=true` in the auth DB so they bypass billing throttles.

## Submission notes (free text → Partner Center "Notes for certification")

This add-in is read/write to the active Word document only. It does not read external documents, does not transmit document content beyond explicit user-selected SMILES strings, and does not require any Office account beyond standard add-in install permissions.

The Advanced tab requires a separate "Compute" API key (prefix ncmcp_). The reviewer-test compute key is included in the test accounts above. If the reviewer pastes only the Core key, the Advanced tab will correctly display the gate state ("Compute tier required") rather than failing — that is by design and matches the public user experience for Core-tier customers.

The add-in calls https://api.novomcp.com endpoints documented at https://novomcp.com/docs/api. All endpoints require Bearer token authentication; no anonymous calls are made.

Backend services are deployed on AWS us-east-1; data residency is U.S. only at launch. No EU / EEA personal data is processed.
