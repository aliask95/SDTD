import { defineTool, ToolError } from "@lovable.dev/mcp-js";
import { z } from "zod";
import { supabaseForUser } from "../supabase";

export default defineTool({
  name: "save_preset",
  title: "Save a preset",
  description:
    "Create a saved SDTD preset for the signed-in user: a named set of options for one operation.",
  inputSchema: {
    name: z.string().trim().min(1).describe("Preset name."),
    tool: z.string().trim().min(1).describe("Operation id from list_operations."),
    options: z
      .record(z.union([z.string(), z.boolean(), z.number()]))
      .optional()
      .describe("Option keys and values for the operation."),
    notes: z.string().optional().describe("Optional free-text notes."),
  },
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
  handler: async ({ name, tool, options, notes }, ctx) => {
    if (!ctx.isAuthenticated()) {
      return { content: [{ type: "text", text: "Not authenticated" }], isError: true };
    }
    const userId = ctx.getUserId();
    if (!userId) throw new ToolError("Could not resolve the signed-in user.");
    const supabase = supabaseForUser(ctx);
    const { data, error } = await supabase
      .from("presets")
      .insert({ user_id: userId, name, tool, options: options ?? {}, notes: notes ?? null })
      .select()
      .single();
    if (error) return { content: [{ type: "text", text: error.message }], isError: true };
    return {
      content: [{ type: "text", text: JSON.stringify(data) }],
      structuredContent: { preset: data },
    };
  },
});
