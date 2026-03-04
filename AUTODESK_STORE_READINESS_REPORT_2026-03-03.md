# Autodesk Store Readiness Report

**Date:** 2026-03-03  
**Product candidate:** Spectiv Inventor Automation Suite 2026 (AssemblyClonerAddIn track)  
**Workspace:** INVENTOR_AUTOMATION_SUITE_2026

## Executive Summary

Current state: **Partially ready**.

You are technically close on core add-in functionality, but still missing several submission-critical items:
- Production installer package (MSI) with validated install/uninstall path
- Store/compliance collateral (privacy policy, support policy, listing assets)
- Signing/release governance evidence

**Readiness estimate:** ~55-65% for first submission package.

---

## Evidence Reviewed

- `app-store-getting-started-guide.docx`
- `AUTODESK_STORE_COMPLIANCE.md`
- `AUTODESK_APP_STORE_COMMERCIALIZATION_PLAN.md`
- `INSTALLER_CREATION_GUIDE.md`
- `BUILD_INSTRUCTIONS.md`
- `InventorAddIn/AssemblyClonerAddIn/AssemblyClonerAddIn.vbproj`
- `InventorAddIn/AssemblyClonerAddIn/AssemblyClonerAddIn.addin`
- `InventorAddIn/INVENTOR_DEPLOYMENT_LOCATIONS.md`

---

## Store Requirements Status Matrix

## A) Core technical packaging

1. **Native desktop add-in architecture** — **PASS**
- VB.NET class library targeting .NET Framework 4.8 and Inventor interop.

2. **Installer package (MSI/EXE)** — **GAP (HIGH)**
- No confirmed maintained installer project/output in active add-in track.
- Required before submission hardening.

3. **Install/uninstall reliability** — **GAP (HIGH)**
- Need clean-machine validation evidence (install, load, uninstall, upgrade).

4. **Compatibility declaration** — **PARTIAL**
- Add-in manifest supports broad software version range.
- Store listing still needs explicit tested versions (recommend explicit matrix).

5. **Digital signing** — **GAP (MED-HIGH)**
- Need signed DLL + signed MSI/EXE for trust/review quality.

## B) Policy/compliance

6. **Privacy policy URL and in-app availability** — **GAP (HIGH)**
- DOCX explicitly requires both listing URL and privacy text accessible in app.

7. **Support contact + SLA statement** — **GAP (MED)**
- Must provide support channel and response commitment.

8. **EULA/license terms** — **GAP (MED)**
- Required for commercial clarity; strongly expected for paid/trial models.

9. **Data collection disclosure** — **PARTIAL**
- If telemetry/log upload exists or is added, policy must describe collection/use/retention/deletion/revocation.

## C) Commercial/listing operations

10. **Listing assets (icon/screenshots/description/video)** — **GAP (HIGH)**
- Must prepare production-quality listing assets.

11. **Pricing/payment setup (PayPal for paid)** — **GAP (MED)**
- Required if paid listing.

12. **Entitlement/IPN integration (optional by business model)** — **OPTIONAL / DECIDE**
- Needed only if you choose Autodesk entitlement workflow / IPN callbacks.

---

## Prerequisite Decision (Inventor Apprentice)

## Conclusion
- **Inventor Apprentice is NOT a required prerequisite** for this in-process Inventor add-in submission baseline.

## Required runtime prerequisites to publish and support
- Autodesk Inventor (supported version set, e.g. 2026).
- .NET Framework 4.8 runtime.

## When Apprentice becomes required
- Only if you ship separate headless document-processing utilities that instantiate Apprentice outside Inventor.

---

## Priority Action Plan (submission-critical)

## P0 (must close first)
1. Build and validate MSI installer (per-user AppData add-ins path).
2. Produce privacy policy and in-app privacy access.
3. Prepare listing assets and store metadata.
4. Complete clean-machine install/uninstall/upgrade validation report.

## P1 (strongly recommended before submit)
5. Code-sign DLL and MSI/EXE with timestamp.
6. Add explicit compatibility matrix (Inventor versions + OS versions).
7. Finalize support process (email, hours, SLA, issue triage flow).

## P2 (business optimization)
8. Decide on free/trial/paid model and entitlement/IPN strategy.
9. Add crash-safe telemetry only if policy-compliant and useful.

---

## Submission Artifact Checklist (what you should have in hand)

## Technical package
- MSI (and optional setup bootstrapper)
- Release DLL + `.addin` payload map
- SHA256 hashes for released artifacts
- Signed binaries/packages

## Documentation
- Install guide
- User guide / quickstart
- Release notes
- Privacy policy
- EULA/license terms
- Support policy/contact

## Listing assets
- App icon
- 5-10 screenshots
- Short + long descriptions
- Category/tags
- Compatibility statement
- Optional short demo video

## Validation evidence
- Clean machine install log
- Uninstall log
- Upgrade test log
- Core workflow smoke test evidence

---

## Risks if submitted now

- Rejection/delay due to missing production installer evidence.
- Compliance feedback loop due to missing privacy/EULA/support artifacts.
- Lower trust and conversion without signing and polished listing assets.

---

## Suggested Timeline (fast-track)

- **Day 1-2:** MSI project creation + packaging pipeline
- **Day 3:** Clean-machine install/uninstall/upgrade validation
- **Day 4:** Privacy policy + support + EULA finalization
- **Day 5:** Listing assets and metadata completion
- **Day 6:** Signing + final release candidate bundle
- **Day 7:** Submission

---

## Immediate Next Best Step

Use the implementation document:
- `MSI_IMPLEMENTATION_PLAN_AssemblyClonerAddIn.md`

Then execute P0 items in order and track completion evidence in a single release folder.
