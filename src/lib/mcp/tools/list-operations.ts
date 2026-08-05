import { defineTool } from "@lovable.dev/mcp-js";

const OPERATIONS = [
  {
    id: "apply_layout",
    tab: "Layout",
    name: "Copy Page Layout",
    description:
      "Copies page size, headers, footers and their content from the Source document to the Target document.",
    options: ["pageSize", "header", "footer", "copyHeaderFooterContent", "deleteEmptyParagraphs"],
  },
  {
    id: "copy_styles",
    tab: "Layout",
    name: "Copy Styles",
    description: "Copies paragraph styles from Source to Target.",
    options: [],
  },
  {
    id: "format_helper",
    tab: "Format",
    name: "Format Helper",
    description:
      "Detects Italic/Bold/Bold+Italic text in the Source and adds review comments in the Target (author: Format Assist).",
    options: [],
  },
  {
    id: "clean_meps",
    tab: "Format",
    name: "Clean the MEPS Up",
    description: "Removes text using the selected MEPS fonts from the Target document.",
    options: ["mepsBookmanWTS", "mepsBookmanUniversal"],
  },
  {
    id: "replace_key_phrases",
    tab: "ADP",
    name: "Replace Key Phrases",
    description:
      "Uses an Excel dictionary (Column A to Column B) to replace phrases in the Target document.",
    options: ["excludeComments", "includeHeaderFooter"],
  },
  {
    id: "transfer_dp_comments",
    tab: "ADP",
    name: "Transfer DP Comments",
    description:
      'Copies comments by author "Digital Publishing" from Source to the matching paragraph in Target.',
    options: ["findIdentifier", "replaceWith", "excludeIt"],
  },
  {
    id: "detect_missing_sections",
    tab: "ADP",
    name: "Detect Missing Sections",
    description: "Reports headings present in Source but missing from Target.",
    options: [],
  },
] as const;

export default defineTool({
  name: "list_operations",
  title: "List SDTD operations",
  description:
    "List the document operations SDTD supports, with their tab, description and configurable option keys.",
  inputSchema: {},
  annotations: { readOnlyHint: true, idempotentHint: true, openWorldHint: false },
  handler: () => ({
    content: [{ type: "text", text: JSON.stringify(OPERATIONS, null, 2) }],
    structuredContent: { operations: OPERATIONS },
  }),
});
