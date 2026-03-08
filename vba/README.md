# SDTD – VBA Legacy Add-in

## Setup Instructions

### Option A: Word Template (.dotm) — No Visual Studio needed
1. Open Word → **Alt+F11** (VBA Editor)
2. Insert → Module → Paste each `.bas` file into a separate module
3. Insert → UserForm → Import `frmSDTD` (see UserForm section)
4. File → Save as → **Word Macro-Enabled Template (.dotm)**
5. Save to: `%appdata%\Microsoft\Word\STARTUP\`
6. Restart Word — SDTD appears in the Add-ins tab

### Option B: VSTO Add-in (.dll) — Requires Visual Studio
1. Open Visual Studio → New Project → **Word VSTO Add-in**
2. Copy the VBA logic into C# equivalent classes
3. Use the Ribbon XML from `ribbon/CustomRibbon.xml`
4. Build → produces `.dll` + `.vsto` installer
5. Run the installer on target machines

---

## File Structure

```
vba/
├── README.md                  ← This file
├── modMain.bas                ← Entry point, toolbar/menu setup
├── modLayout.bas              ← Copy Page Layout, Copy Styles
├── modFormatting.bas           ← Format Helper, Clean MEPS
├── modADP.bas                 ← Replace Key Phrases, Transfer DP Comments, Detect Missing Sections
├── modBatch.bas               ← Batch processing across multiple files
├── modUtils.bas               ← Shared utilities (document picker, progress bar, etc.)
├── frmSDTD.frm.md             ← UserForm design instructions
└── ribbon/
    └── CustomRibbon.xml       ← For VSTO: Ribbon UI definition
```

## References Required (Tools → References in VBA Editor)
- **Microsoft Word Object Library** (auto-included)
- **Microsoft Excel Object Library** (for ADP Excel dictionary)
- **Microsoft Office Object Library** (for FileDialog)
