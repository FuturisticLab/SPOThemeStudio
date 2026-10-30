# SharePoint Online Theme Studio & Interactive Visualizer

A professional, interactive theme design and deployment utility for **SharePoint Online**. It enables administrators, architects, and designers to generate harmonious Fluent UI color combinations, visually inspect which hex code corresponds to every specific component on a SharePoint page, and generate copy-paste ready `Add-SPOTheme` / `Set-SPOWebTheme` PowerShell scripts.

---

## 🌟 Key Features

1. **Bidirectional Interactive Visualizer**:
   - **Click or hover any element on the realistic SharePoint page mockup** (Suite Bar, Site Header, Nav tabs, Hero banner, Primary buttons, Quick Links cards, Document Library, News articles, Footer) to immediately highlight its mapped theme slot, hex code, and functional description.
   - **Hover or select any slot in the form** to see all corresponding page areas illuminate with high-contrast outlines!
2. **Complete 28-Slot Support (All 22 Required + 6 Semantic Slots)**:
   - Includes all 9 brand accent slots (`themePrimary` to `themeDarker`).
   - Includes all 13 neutral surface & text slots (`neutralLighterAlt` to `white`).
   - Includes extended semantic slots (`primaryBackground`, `bodyText`, etc.).
3. **Automated Combination & Scheme Generators**:
   - **Instant Presets**: Pre-configured Microsoft Fluent palettes including *Microsoft Modern Blue*, *Cyberdyne Crimson HUD (Dark)*, *Emerald Forest*, *Teal Horizon*, *Midnight Amethyst*, *Sunset Amber*, and *Obsidian Dark Executive*.
   - **🎲 Randomize Scheme**: Generates fresh, accessible color combinations dynamically using HSL color harmony algorithms.
   - **⚡ Single-Base Auto-Generator**: Enter any single corporate hex code (e.g. `#107C41` or `#0078D4`), and the utility auto-calculates all tints, shades, and neutrals matching Microsoft Fluent UI lightness steps.
   - **🌓 Light / Inverted Dark Mode Toggle**: Automatically inverts neutrals and tints for dark mode compliance.
4. **Native Windows GUI (WPF) + Browser Companion**:
   - **`SPOThemeStudio.ps1`**: Full desktop application with native Windows color picker dialogs and zero external dependencies.
   - **`SPO-Theme-Studio-Interactive.html`**: Zero-dependency standalone HTML5/JS app for modern web browsers.
5. **Deployment Ready**:
   - One-click copy for complete `Add-SPOTheme` & `Set-SPOWebTheme` PowerShell scripts.
   - Direct tenant deployment trigger via `Connect-SPOService`.
   - Export to `.ps1` file or JSON palette.

---

## 🚀 Quickstart

### Option 1: Native PowerShell GUI (Recommended)

Run directly in PowerShell (PowerShell 5.1 or PowerShell 7+):

```powershell
powershell -ExecutionPolicy Bypass -File .\SharePoint-Theme-Studio\SPOThemeStudio.ps1
```

Or from the `SharePoint WebPart` folder:
```powershell
.\Launch-SPOThemeStudio.ps1
```

### Option 2: Browser Companion

Double-click or open in Microsoft Edge / Chrome:
```
SharePoint-Theme-Studio\SPO-Theme-Studio-Interactive.html
```

---

## 🎨 SharePoint Theme Slot Reference Guide

| Slot Key | Type | Typical Light Value | Inverted Dark Value | Where It Appears in SharePoint Online |
| :--- | :--- | :--- | :--- | :--- |
| `themePrimary` | Brand Accent | `#0078d4` | `#ff1e3c` | Suite Bar accent, primary buttons (`+ New`, `Share`), active navigation underlines, hyperlinks, checkboxes, hero CTA buttons. |
| `themeLighterAlt` | Tint | `#eff6fc` | `#0a0103` | Selected list row background highlight, subtle command bar hover state. |
| `themeLighter` | Tint | `#deecf9` | `#29050a` | Hover highlight for secondary pills, progress bar background tracks. |
| `themeLight` | Tint | `#c7e0f4` | `#4d0912` | Command bar button hover fill, notification alert cards. |
| `themeTertiary` | Mid-Tone | `#71afe5` | `#991224` | Hyperlink hover color, secondary progress bars, tag outlines. |
| `themeSecondary` | Shade | `#2b88d8` | `#e01a35` | Secondary button outlines, primary button hover tint, active folder icons. |
| `themeDarkAlt` | Shade | `#106ebe` | `#ff3550` | Primary button hover background, hero banner gradient / background fill. |
| `themeDark` | Shade | `#005a9e` | `#ff546b` | Primary button pressed state, focused text input borders, active nav text. |
| `themeDarker` | Shade | `#004578` | `#ff8293` | Pressed hyperlink state, highest contrast brand text on light surfaces. |
| `neutralLighterAlt` | Neutral | `#faf9f8` | `#111319` | Card container surfaces, web part canvas containers, alternating table rows. |
| `neutralLighter` | Neutral | `#f3f2f1` | `#161922` | Default page background canvas (light theme), secondary button fills, footer bar. |
| `neutralLight` | Neutral | `#edebe9` | `#1f2430` | Section divider lines, subtle card borders, table column separators. |
| `neutralQuaternaryAlt`| Neutral | `#e1dfdd` | `#262b38` | Table borders, list row separator lines, disabled checkbox borders. |
| `neutralQuaternary` | Neutral | `#d2d0ce` | `#2c3242` | Input borders (search box, text inputs), secondary card outlines. |
| `neutralTertiaryAlt` | Neutral | `#c8c6c4` | `#444c63` | Disabled icons, placeholder glyphs, muted tags. |
| `neutralTertiary` | Neutral Text | `#a19f9d` | `#c8c8c8` | Disabled text on buttons, search box placeholder text, metadata timestamps. |
| `neutralSecondary` | Neutral Text | `#605e5c` | `#d0d0d0` | Secondary metadata text ("Modified 2h ago"), subtitles, inactive navigation tabs. |
| `neutralPrimaryAlt` | Neutral Text | `#3b3a39` | `#dadada` | Hover text for dark titles, sub-headings. |
| `neutralPrimary` | Neutral Text | `#323130` | `#ffffff` | Main body paragraph text, primary list labels, document titles, web part headings. |
| `neutralDark` | Neutral Text | `#201f1e` | `#f4f4f4` | Maximum contrast headings, hero titles, page H1 headings. |
| `black` | Neutral Core | `#000000` | `#f8f8f8` | Deepest black text. In inverted themes, mapped to high-contrast white. |
| `white` | Neutral Core | `#ffffff` | `#090a0d` | Standard page canvas surface, card surfaces, primary button text. *In inverted themes, mapped to deep dark canvas.* |

> [!NOTE]
> **Inverted / Dark Themes (`-IsInverted $true`)**:
> In SharePoint Online, when `-IsInverted $true` is set, Microsoft reverses the roles of `white` and `black`. The `white` slot defines the deep dark page background (e.g., `#090A0D` or `#0F1117`), while `black` and `neutralPrimary` define the bright readable text (e.g., `#FFFFFF`).

---

## 📋 Generated PowerShell Command Reference

```powershell
# Connect to SharePoint Online Admin Center
Connect-SPOService -Url "https://YOURTENANT-admin.sharepoint.com"

# Register theme in tenant catalog
Add-SPOTheme -Identity "Contoso Enterprise Blue" `
             -Palette $ThemePalette `
             -IsInverted $false `
             -Overwrite

# Apply theme to target site collection
Set-SPOWebTheme -Theme "Contoso Enterprise Blue" `
                -Web "https://YOURTENANT.sharepoint.com/sites/portal"
```

---

## 📚 Technical Documentation & Enterprise Guide

For complete technical specifications, review checklists, CI/CD pipeline definitions, enterprise DevOps guides, user workflows, and risk analysis, see the comprehensive [DOCUMENTATION.md](file:///c:/Users/jmarc/Desktop/FirebaseMigration/SharePoint-Theme-Studio/DOCUMENTATION.md):
- **Section 9**: Setup for Code Review & Static Analysis (`PSScriptAnalyzer` config & checklist)
- **Section 10**: Continuous Integration & Delivery (`.github/workflows/ci-cd.yml` & Azure DevOps)
- **Section 11**: Updates, Semantic Versioning & Dependency Maintenance
- **Section 12**: DevOps, GitOps, Drift Detection & Infrastructure as Code (IaC)
- **Section 13**: End-to-End User Guide & Step-by-Step Walkthroughs
- **Section 14**: Enterprise Risk Review, Blast-Radius Analysis & Emergency Rollback Procedures
- **Section 15**: Estimated Product Value, Commercial Valuation & Business Implementation ROI
  - **§15.3**: IT Professional Practitioner Profile & 5W1H Operational Framework (What, Why, When, How, Where)
  - **§15.4**: Enterprise Business Use Cases (M&A, MSPs, Viva Connections, Brand Governance)
- **Section 16**: Project Leadership, Authorship & AI Collaboration Model (J.M. + Google DeepMind Antigravity / Gemini)

---

## 👥 Authorship & Engineering Credits

- **Visionary, Solutions Architect & Orchestrator**: **J.M.** — The Brains, Creative Mind & Enterprise Architect driving product vision, UX/UI paradigm, deep root-cause diagnostics, and security governance.
- **AI Co-Author & Engineering Partner**: **Antigravity** (Powered by **Google DeepMind's Gemini** advanced agentic coding models) — Rapid code synthesis, mathematical color algorithms, AST testing harnesses, and CI/CD automation.
- **Total Delivery Velocity**: Conceived, architected, built, tested, security-hardened, and fully documented in **under 4.5 hours**.

---

## 🛠️ Testing & Validation

To run the automated verification test on the utility:
```powershell
powershell -ExecutionPolicy Bypass -File .\SharePoint-Theme-Studio\Test-SPOThemeStudio.ps1
```
