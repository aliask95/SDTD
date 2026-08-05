import { defineTool } from "@lovable.dev/mcp-js";
import { z } from "zod";
import { supabaseForUser } from "../supabase";

export default defineTool({
  name: "list_presets",
  title: "List saved presets",
  description: "List the signed-in user's saved SDTD presets, optionally filtered by operation id.",
  inputSchema: {
    tool: z.string().optional().describe("Optional operation id, e.g. apply_layout."),
  },
  annotations: { readOnlyHint: true, idempotentHint: true, openWorldHint: false },
  handler: async ({ tool }, ctx) => {
    if (!ctx.isAuthenticated()) {
      return { content: [{ type: "text", text: "Not authenticated" }], isError: true };
    }
    const supabase = supabaseForUser(ctx);
    let query = supabase
      .from("presets")
      .select("id, name, tool, options, notes, created_at, updated_at")
      .order("updated_at", { ascending: false });
    if (tool) query = query.eq("tool", tool);
    const { data, error } = await query;
    if (error) return { content: [{ type: "text", text: error.message }], isError: true };
    return {
      content: [{ type: "text", text: JSON.stringify(data ?? [], null, 2) }],
      structuredContent: { presets: data ?? [] },
    };
  },
});
