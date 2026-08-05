import { useEffect, useState } from "react";
import { useNavigate, useSearchParams } from "react-router-dom";
import { supabase } from "@/integrations/supabase/client";
import { toast } from "sonner";
import sdtdLogo from "@/assets/sdtd-logo.png";

function safeNext(value: string | null): string {
  if (!value) return "/";
  if (!value.startsWith("/") || value.startsWith("//")) return "/";
  return value;
}

const Auth = () => {
  const [params] = useSearchParams();
  const navigate = useNavigate();
  const next = safeNext(params.get("next"));
  const [mode, setMode] = useState<"signin" | "signup">("signin");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      if (data.session) window.location.replace(next);
    });
  }, [next]);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setBusy(true);
    if (mode === "signin") {
      const { error } = await supabase.auth.signInWithPassword({ email, password });
      setBusy(false);
      if (error) return toast.error(error.message);
      window.location.replace(next);
    } else {
      const { error } = await supabase.auth.signUp({
        email,
        password,
        options: { emailRedirectTo: `${window.location.origin}${next}` },
      });
      setBusy(false);
      if (error) return toast.error(error.message);
      toast.success("Account created. You can sign in now.");
      setMode("signin");
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-4 bg-background">
      <div className="w-full max-w-[360px] rounded-xl border border-border bg-card shadow-2xl overflow-hidden">
        <div
          className="flex items-center gap-3 px-4 py-3"
          style={{ background: "linear-gradient(135deg, hsl(190 79% 35%), hsl(190 79% 28%))" }}
        >
          <img src={sdtdLogo} alt="SDTD logo" className="w-8 h-8 rounded-lg" />
          <div className="flex flex-col">
            <span className="text-sm font-bold tracking-wider text-primary-foreground">SDTD</span>
            <span className="text-[10px] text-primary-foreground/60 font-medium">
              {mode === "signin" ? "Sign in to your account" : "Create your account"}
            </span>
          </div>
        </div>
        <form onSubmit={submit} className="p-4 space-y-3">
          <div>
            <label className="text-[11px] font-bold uppercase tracking-wider text-muted-foreground">Email</label>
            <input
              type="email"
              required
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              className="w-full mt-1 px-2 py-1.5 text-xs rounded border border-input bg-background text-foreground"
            />
          </div>
          <div>
            <label className="text-[11px] font-bold uppercase tracking-wider text-muted-foreground">Password</label>
            <input
              type="password"
              required
              minLength={6}
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full mt-1 px-2 py-1.5 text-xs rounded border border-input bg-background text-foreground"
            />
          </div>
          <button
            type="submit"
            disabled={busy}
            className="w-full px-3 py-1.5 text-[11px] font-bold uppercase tracking-wider rounded bg-primary text-primary-foreground hover:brightness-110 disabled:opacity-60"
          >
            {mode === "signin" ? "Sign in" : "Create account"}
          </button>
          <button
            type="button"
            onClick={() => setMode(mode === "signin" ? "signup" : "signin")}
            className="w-full text-[11px] text-muted-foreground hover:text-foreground"
          >
            {mode === "signin" ? "Need an account? Sign up" : "Already have an account? Sign in"}
          </button>
        </form>
      </div>
    </div>
  );
};

export default Auth;
