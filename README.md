# SDTD – Source to Target Document Word Add-in

SDTD (Source to Target Document) is a **Microsoft Word Add-in** designed to automate and accelerate workflows that involve transferring content, formatting, styles, and comments between two documents.

The add-in allows the user to define a **Source document** and a **Target document** and apply structured operations between them.

It is especially useful for **translation, digital publishing workflows, and structured document processing**.

---

# Core Concept

Many publishing workflows require working with **two documents simultaneously**:

- A **Source document** containing the original content.
- A **Target document** containing the translated or processed content.

SDTD simplifies this process by allowing the user to select both documents and run **automation tools** that synchronize elements between them.

---

# User Interface

The add-in opens a **Task Pane** inside Microsoft Word.

At the top of the pane the user selects:

```

Source Document: [Dropdown]
Target Document: [Dropdown]

```

These dropdown menus dynamically list **all currently opened Word documents**.

The add-in then applies operations **from Source → Target**.

---

# Task Pane Structure

The pane is divided into three main functional sections.

```

Layout
Formatting
ADP (Advanced Document Processing)

```

Each section contains tools that perform specific operations.

---

# Layout

Tools that transfer **document structure and styles**.

## Copy Page Layout

Copies page layout settings from the Source document to the Target document.

Options:

- Page Size
- Header
- Footer
- Copy Header/Footer Content

Example behavior:

- Copy margins and page size
- Replace header/footer structure
- Optionally copy header/footer content

---

## Copy Styles

Copies styles from the Source document and updates them in the Target document.

Supported styles:

- Paragraph styles
- Character styles

Behavior:

- If the style does not exist in Target → it is created.
- If the style exists → it is updated.

Copied properties include:

- Base style
- Font settings
- Paragraph formatting

---

# Formatting

Tools related to **text formatting detection and validation**.

## Format Helper

Detects formatting in the Source document and creates comments in the Target document to verify whether the formatting is correct.

Supported detection:

- Italic
- Bold
- Bold + Italic

Example comment inserted:

```

Author: Format Assist

Is the formatting (Italic/Bold/Bold+Italic) correct here?

```

This helps translators and editors verify formatting consistency.

---

# ADP – Advanced Document Processing

Automation tools for advanced document workflows.

---

## Transfer DP Comments

Transfers comments written by **Digital Publishing** from the Source document to the corresponding paragraph in the Target document.

Process:

1. Scan comments in Source
2. Filter comments where

```

Author = Digital Publishing

```

3. Insert the comment at the **end of the corresponding paragraph** in the Target document.

---

## Language Identifier Replacement

DP comments often contain language identifiers.

Example:

```

=E

```

The add-in allows replacing these identifiers.

User settings:

```

Find: =E
Replace: =BL

```

Additional option:

```

☑ Exclude comments containing "it"

```

This skips comments containing the identifier `it`.

---

## Detect Missing Sections

Compares Source and Target documents to detect missing sections.

Possible detection methods:

- Compare heading structure
- Identify headings present in Source but missing in Target

Results can be displayed as:

- Comments
- Task pane report

---

## Replace Key Phrases

Replaces phrases in the Target document using a **user-defined Excel dictionary**.

Excel structure:

| Column A | Column B |
|----------|----------|
| English phrase | Target phrase |

Example:

| Source Phrase | Replacement |
|--------------|------------|
| Kingdom Hall | Зала на Царството |

Features:

- Select an Excel file containing the dictionary
- Automatically replace phrases in the Target document
- Remember the **last used Excel file**

---

## Text Cleanup

Automated cleanup tools for common publishing corrections.

Available operations:

- Replace `par.` and `pars.` with Bulgarian equivalent `абз.`
- Replace spaces before numbered references with **en spaces**

Example transformation:

```

1. →  1.

```

---

## Document Metadata

Automates metadata generation from document content.

Features:

- Remove text before the first Cyrillic character in the first paragraph
- Convert the first paragraph to **sentence case**
- Set it as the document **Title property**

---

# Smart Sync (Planned)

A combined automation tool that executes multiple operations in one step.

Example:

```

Smart Sync
☑ Copy Styles
☑ Transfer DP Comments
☑ Replace Language Identifiers
☑ Apply Text Cleanup

[Run Smart Sync]

```

This replicates common publishing workflows with one action.

---

# Batch Processing

Batch mode allows applying selected operations to multiple documents.

Options:

```

Enable Batch Mode
Select Files
Select Folder

```

Possible use cases:

- Process multiple translated documents
- Apply formatting corrections to entire projects
- Run dictionary replacements across many files

---

# Technology Stack

Planned technologies:

- Office.js
- JavaScript / TypeScript
- HTML / CSS
- Excel file parsing for dictionaries

---

# Project Status

Early development.

The current goal is to build the **core automation tools for Source → Target document workflows**.

---

# License

This project is currently a personal productivity tool and may be expanded in the future.
