# NovoMCP for Word — Microsoft AppSource Reviewer Test Plan

Submitted to Partner Center → Offer → Properties → "Notes for certification"
or uploaded as a PDF if the field char limit is exceeded.

The goal of this document: a Microsoft reviewer with no prior context can install
this add-in, exercise its primary features, confirm it behaves as described in the
listing, and complete certification in a single sitting (≈ 20 minutes).

---

## 1. About the add-in

NovoMCP brings a computational chemistry engine into Word. The user highlights a
SMILES string in their document and the taskpane returns molecular properties,
ADMET predictions, regulatory compliance verdicts, and similarity hits across
122 million molecules. An Advanced tab gates quantum chemistry tools (pKa,
solubility, bond dissociation energies, frontier orbitals) behind a paid Compute
tier.

The add-in is read/write to the active Word document. It transmits only the
SMILES strings the user explicitly selects, plus an `X-Novo-Surface: word-addin-v1`
header for audit isolation. No document content is exfiltrated.

## 2. Reviewer credentials

**The add-in has no sign-in.** It authenticates solely by a NovoMCP API key that
you paste into the taskpane (sent as `Authorization: Bearer <key>`). There is no
username/password and no email login — you only need the two keys below.

| Tier | Unlocks |
|---|---|
| **Core** (`nmcp_`) | Profile, ADMET, Compliance, Similar tabs |
| **Compute** (`ncmcp_`) | The above + Advanced tab (pKa, solubility, BDE, frontier orbitals) |

Paste the **Core** key in §4 to unlock the main tabs; when you reach the Advanced
tab (§8), paste the **Compute** key as well. Copy the exact values below:

**Core key** (`nmcp_`):

```
{{CORE_KEY}}
```

**Compute key** (`ncmcp_`):

```
{{COMPUTE_KEY}}
```

Both keys are active on an Enterprise plan (50,000 credits/month) — far more than
a review session consumes, so there is no throttling — and remain valid for the
certification lifecycle. If a key ever fails authentication during review, contact
support@novomcp.com (response SLA < 4 hours).

<!-- The {{CORE_KEY}} / {{COMPUTE_KEY}} tokens are substituted with the live key
values ONLY in the PDF/notes delivered to Microsoft via Partner Center. Do NOT
commit live key values to this file. -->
_(Reviewers: the live key values appear in this document as delivered in Partner
Center — "Additional Certification Info" PDF and/or "Additional notes for
certification." They are visible only to Microsoft certification, not publicly.)_

## 3. Setup (5 min)

1. Open Word on the web at https://www.office.com → New blank document.
2. Insert menu → Add-ins → "Upload My Add-in".
3. Install the manifest using any of the three methods below (all equivalent):

   **Method A — paste the URL directly (fastest):**
   In the "Upload My Add-in" dialog, paste:
   ```
   https://addin.novomcp.com/manifest.xml
   ```

   **Method B — download, then drag-and-drop (proves the hosted version works):**
   ```bash
   curl -o ~/Downloads/novomcp-manifest.xml https://addin.novomcp.com/manifest.xml
   ```
   Then drag `~/Downloads/novomcp-manifest.xml` into the "Upload My Add-in" dialog.

   **Method C — Partner Center attachment:**
   Drag-and-drop the manifest XML attached in the Partner Center submission.

4. Confirm the "NovoMCP" button appears on the Home ribbon.
5. Click the NovoMCP button → the taskpane opens on the right of the document.

Expected: a "Connect" panel with an API-key input field. No spinners or error
banners. Taskpane width is 350 px default.

## 4. Authentication (2 min)

1. In the taskpane, paste the Core API key from §2.
2. Click "Connect".
3. The taskpane transitions to the main view with five tabs visible:
   Profile, ADMET, Compliance, Similar, Advanced. The Advanced tab has a
   lock icon (🔒) since the Core key does not unlock Compute.
4. The user-info area shows the account email and tier (Core).

Expected: key is accepted within 2 seconds. Key persists across taskpane
reopen — close the pane, reopen via the ribbon button, the user remains
authenticated.

## 5. Profile tab — the core flow (3 min)

1. In the Word document, type the following sentence:

   `Aspirin (CC(=O)OC1=CC=CC=C1C(=O)O) is a non-steroidal anti-inflammatory.`

2. Double-click to select the SMILES `CC(=O)OC1=CC=CC=C1C(=O)O`.
3. The taskpane Profile tab populates automatically within 2 seconds.

Expected values for aspirin (these are deterministic — the reviewer can verify
against PubChem CID 2244):

| Field | Expected value | Tolerance |
|---|---|---|
| Molecular weight | 180.16 g/mol | ± 0.01 |
| logP (XLogP3) | 1.2 | ± 0.2 |
| TPSA | 63.6 Å² | ± 0.5 |
| H-bond donors | 1 | exact |
| H-bond acceptors | 3 | exact |
| Rotatable bonds | 3 | exact |
| QED v2 | 0.55 | ± 0.05 |
| Lipinski | Pass | exact |
| BOILED-Egg | HIA-Yes / BBB-No | exact |

If any value falls outside tolerance, treat as a defect — do not approve.

## 6. ADMET tab (2 min)

1. With the aspirin SMILES still selected, click the ADMET tab.
2. The tab displays a "loading…" state for ≤ 4 seconds, then populates with
   31 ADMET endpoints grouped by category (CYP, toxicity, distribution,
   metabolism, excretion).

Expected:
- CYP3A4 substrate: low probability (≤ 0.3)
- hERG: low risk (≤ 0.3)
- AMES mutagenicity: negative
- Hepatotoxicity (DILI): low
- Five endpoints display a small "SOTA" badge — CYP2D6, CYP3A4 (both
  inhibitor models), CYP3A4 Substrate, Clearance Hepatocyte, DILI. These are
  the NovoExpert-2 wins on the TDC ADMET benchmark.

## 7. Compliance tab (2 min)

1. Still on aspirin → click Compliance.
2. Expected verdict: "Pass" (aspirin is clean against FAVES v4.1).

Now test the negative path:

3. In the document, type a known PAINS-flagged compound:
   `Quercetin (O=C1c2c(O)cc(O)cc2OC(c2ccc(O)c(O)c2)=C1O)`
4. Select the SMILES → Compliance tab.
5. Expected verdict: "Review" or "Fail". The panel lists the specific
   PAINS / Brenk alerts triggered (catechol, ene_six_het, ortho-quinone or
   similar — exact set depends on FAVES filter set version, but at least one
   alert MUST fire). A "Recommendations" list explains why the alerts
   matter and what scaffold modifications would clear them.

If quercetin returns "Pass," treat as a defect.

## 8. Similar tab (2 min)

1. Select the aspirin SMILES.
2. Click the Similar tab → adjust threshold slider to 0.85 → click "Run".
3. Expected: a results list within 5 seconds with at least 3 hits, each
   showing a similarity score ≥ 0.85, the molecule name (or PubChem CID),
   and a "View profile" button.
4. Click "View profile" on any hit → the Profile tab opens with that
   molecule's data.

## 9. Document scan (3 min)

1. Create a new Word document with a 2-column table:

   | Compound | SMILES |
   |---|---|
   | Aspirin | CC(=O)OC1=CC=CC=C1C(=O)O |
   | Ibuprofen | CC(C)Cc1ccc(C(C)C(=O)O)cc1 |
   | Caffeine | Cn1cnc2c1c(=O)n(C)c(=O)n2C |

2. In the taskpane → click "Scan Document" (top of taskpane, above the tabs).
3. The scan completes in ≤ 8 seconds and shows a per-table summary:
   - Found 3 compounds, all validated.
   - Each row gets a ✓ green check.

Expected: no console errors, no document modifications unless the reviewer
explicitly clicks "Insert as Word table" inside a result panel.

## 10. Advanced tab gating (1 min)

1. With the Core API key active, click the Advanced tab.
2. Expected: a clear gate panel — "Compute tier required" — with copy
   explaining that pKa / solubility / BDE / frontier-orbital lookups require
   an `ncmcp_` Compute key, and a link to https://novomcp.com/pricing.
3. The tab does NOT throw an error or attempt the call.

Expected behavior: graceful gate, no spinner-of-death.

## 11. Compute tier — Advanced tab (3 min)

1. In the taskpane → click the user avatar / settings icon → "Sign out".
2. Paste the Compute API key from §2 → Connect.
3. The Advanced tab unlocks (lock icon disappears).
4. Select the aspirin SMILES → Advanced tab → click "Run" on the pKa card.
5. Expected: a pKa value of 3.5 ± 0.3 returns within 30 seconds.
6. Click "Insert as Word table" → a small 2-row table appears in the Word
   document with the SMILES and the predicted pKa.

Repeat for one additional tool (e.g., HOMO/LUMO frontier orbitals) — return
within 60 seconds, insertable into the document.

## 12. Privacy and permissions (verification, no clicks)

The manifest declares `ReadWriteDocument` permission. This is required because
the "Insert as Word table" buttons write result tables back into the active
document. The add-in does NOT:

- Read documents other than the active one.
- Modify the document unless the user clicks an explicit "Insert" button.
- Transmit document content beyond user-selected SMILES strings.
- Access Office account data beyond what Office.js exposes by default
  (no Graph API calls, no Microsoft 365 account claims read).

Network traffic from the taskpane is limited to:
- `https://api.novomcp.com/*` — NovoMCP MCP server
- `https://app.novomcp.com/*` — token exchange / user info
- `https://addin.novomcp.com/*` — static assets (the manifest origin)

All three are declared in `<AppDomains>` in the manifest.

## 13. Uninstall test (1 min)

1. Insert menu → Add-ins → My Add-ins → find NovoMCP → "..." → Remove.
2. Confirm the ribbon button disappears.
3. Confirm any inserted Word tables remain in the document (uninstall must
   not modify document content).

## 14. Pass/fail criteria summary

For Microsoft certification, all of the following must hold:

- [ ] Add-in installs cleanly via the manifest URL.
- [ ] Both test API keys authenticate.
- [ ] Profile tab returns expected values for aspirin within tolerance.
- [ ] ADMET tab returns 31 endpoints, no missing values.
- [ ] Compliance tab returns the correct verdict for aspirin (Pass) and
      for quercetin (Review or Fail with PAINS alerts visible).
- [ ] Similar tab returns ≥ 3 hits for aspirin at threshold 0.85.
- [ ] Document scan finds all 3 SMILES in the test table.
- [ ] Advanced tab gates cleanly on the Core key and unlocks on the Compute key.
- [ ] pKa and HOMO/LUMO calculations return values within published tolerance.
- [ ] No document content is exfiltrated beyond user-selected SMILES.
- [ ] Uninstall is clean.

## 15. Reviewer support contact

If any step fails or behaves unexpectedly, contact:

- **Email:** ari@novomcp.com
- **Response SLA during certification:** < 4 business hours

We can replay the exact reviewer session from server-side audit logs using
the `X-Novo-Surface: word-addin-v1` header — please include the approximate
UTC timestamp of the failed call and we will diagnose within the SLA window.
