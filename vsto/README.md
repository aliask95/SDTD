# SDTD — VSTO Word Add-in

## Setup in Visual Studio

1. Open Visual Studio → **Create a new project** → Search **"Word VSTO Add-in"**  
   (Template: *Word Add-in* under Office/SharePoint → VSTO Add-ins)
2. Name it `SDTD`
3. Copy the files from this folder into the project:
   - `ThisAddIn.cs` → Replace the auto-generated one
   - `SdtdRibbon.xml` → Add as existing item, set Build Action to **Embedded Resource**
   - `SdtdRibbon.cs` → Add as existing item
   - `Tools/*.cs` → Create a `Tools` folder and add all files
4. Add NuGet reference: `Microsoft.Office.Interop.Word` (usually auto-included)
5. Add NuGet reference: `EPPlus` (for Excel reading in Replace Key Phrases)
6. Add reference to `Microsoft.Office.Tools.Common`
7. Build → Debug (F5) — Word will open with the add-in loaded

## Features

- **Copy Page Layout** — page size, headers, footers, content
- **Copy Styles** — paragraph styles from source to target
- **Format Helper** — detects Bold/Italic in source, adds review comments in target
- **Clean MEPS Up** — removes text with MEPS Bookman fonts from target
- **Replace Key Phrases** — Excel dictionary (Col A → Col B) replacement
- **Transfer DP Comments** — copies "Digital Publishing" comments between docs
- **Detect Missing Sections** — compares headings between source and target
- **Batch Mode** — process multiple files with any tool
