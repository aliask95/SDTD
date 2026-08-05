import { auth, defineMcp } from "@lovable.dev/mcp-js";
import listOperationsTool from "./tools/list-operations";
import listPresetsTool from "./tools/list-presets";
import savePresetTool from "./tools/save-preset";
import deletePresetTool from "./tools/delete-preset";

const projectRef = import.meta.env.VITE_SUPABASE_PROJECT_ID ?? "project-ref-unset";

export default defineMcp({
  name: "sdtd",
  title: "SDTD",
  version: "0.1.0",
  instructions:
    "Tools for SDTD, a Word add-in that syncs a Source document to a Target document. Use `list_operations` to see the supported document operations and their option keys, and `list_presets` / `save_preset` / `delete_preset` to manage the signed-in user's saved option presets.",
  auth: auth.oauth.issuer({
    issuer: `https://${projectRef}.supabase.co/auth/v1`,
    acceptedAudiences: "authenticated",
  }),
  tools: [listOperationsTool, listPresetsTool, savePresetTool, deletePresetTool],
});
