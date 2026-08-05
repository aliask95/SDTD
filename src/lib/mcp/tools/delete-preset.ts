import { defineTool } from "@lovable.dev/mcp-js";
import { z } from "zod";
import { supabaseForUser } from "../supabase";

export default defineTool({
  name: "delete_preset",
  title: "Delete a preset",
  description: "Permanently delete one of the signed-in user's saved SDTD presets by id.",
  inputSchema: {
    id: z.string().trim().min(1).describe("Preset id returned by list_presets."),
  },
  annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false },
  handler: async ({ id }, ctx) => {
    if (!ctx.isAuthenticated()) {
      return { content: [{ type: "text", text: "Not authenticated" }], isError: true };
    }
    const supabase = supabaseForUser(ctx);
    const { data, error } = await supabase.from("presets").delete().eq("id", id).select("id");
    if (error) return { content: [{ type: "text", text: error.message }], isError: true };
    if (!data?.length) {
      return { content: [{ type: "text", text: `No preset found with id ${id}.` }], isError: true };
    }
    return { content: [{ type: "text", text: `Deleted preset ${id}.` }] };
  },
});
