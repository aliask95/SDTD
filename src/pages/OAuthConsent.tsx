import { useEffect, useState } from "react";
import { useSearchParams } from "react-router-dom";
import { supabase } from "@/integrations/supabase/client";

type OAuthClient = { name?: string; redirect_uri?: string };
type AuthorizationDetails = {
  client?: OAuthClient;
  scope?: string;
  redirect_url?: string;
  redirect_to?: string;
};
type OAuthApi = {
  getAuthorizationDetails: (id: string) => Promise<{ data: AuthorizationDetails | null; error: { message: string } | null }>;
  approveAuthorization: (id: string) => Promise<{ data: AuthorizationDetails | null; error: { message: string } | null }>;
  denyAuthorization: (id: string) => Promise<{ data: AuthorizationDetails | null; error: { message: string } | null }>;
};
const oauth = () => (supabase.auth as unknown as { oauth: OAuthApi }).oauth;

const SCOPE_LABELS: Record<string, string> = {
  openid: "Confirm who you are",
  email: "Share your email address",
  profile: "Share your basic profile",
};

const OAuthConsent = () => {
  const [params] = useSearchParams();
  const authorizationId = params.get("authorization_id") ?? "";
  const [details, setDetails] = useState<AuthorizationDetails | null>(null);
  const [account, setAccount] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    let active = true;
    (async () => {
      if (!authorizationId) {
        setError("Missing authorization_id.");
        return;
      }
      const { data: sess } = await supabase.auth.getSession();
      if (!sess.session) {
        const next = window.location.pathname + window.location.search;
        window.location.href = "/auth?next=" + encodeURIComponent(next);
        return;
      }
      if (!active) return;
      setAccount(sess.session.user.email ?? null);
      const { data, error: err } = await oauth().getAuthorizationDetails(authorizationId);
      if (!active) return;
      if (err) {
        setError(err.message);
        return;
      }
      const immediate = data?.redirect_url ?? data?.redirect_to;
      if (immediate && !data?.client) {
        window.location.href = immediate;
        return;
      }
      setDetails(data);
    })();
    return () => {
      active = false;
    };
  }, [authorizationId]);

  const decide = async (approve: boolean) => {
    setBusy(true);
    const { data, error: err } = approve
      ? await oauth().approveAuthorization(authorizationId)
      : await oauth().denyAuthorization(authorizationId);
    if (err) {
      setBusy(false);
      setError(err.message);
      return;
    }
    const target = data?.redirect_url ?? data?.redirect_to;
    if (!target) {
      setBusy(false);
      setError("No redirect returned by the authorization server.");
      return;
    }
    window.location.href = target;
  };

  const clientName = details?.client?.name ?? "this app";
  const scopes = (details?.scope ?? "").split(/\s+/).filter(Boolean);

  return (
    <main className="min-h-screen flex items-center justify-center p-4 bg-background">
      <div className="w-full max-w-[400px] rounded-xl border border-border bg-card shadow-2xl p-5 space-y-4">
        {error ? (
          <>
            <h1 className="text-base font-bold text-foreground">Could not load this request</h1>
            <p className="text-xs text-muted-foreground">{error}</p>
          </>
        ) : !details ? (
          <p className="text-xs text-muted-foreground">Loading…</p>
        ) : (
          <>
            <h1 className="text-base font-bold text-foreground">Connect {clientName} to SDTD</h1>
            {account && (
              <p className="text-xs text-muted-foreground">
                Signed in as <span className="text-foreground">{account}</span>
              </p>
            )}
            <p className="text-xs text-foreground">
              {clientName} will be able to call SDTD's enabled tools as you.
            </p>
            {details.client?.redirect_uri && (
              <p className="text-[11px] text-muted-foreground break-all">
                Redirects to {details.client.redirect_uri}
              </p>
            )}
            {scopes.length > 0 && (
              <ul className="space-y-1">
                {scopes.map((s) => (
                  <li key={s} className="text-xs text-muted-foreground">
                    • {SCOPE_LABELS[s] ?? `Additional permission requested: ${s}`}
                  </li>
                ))}
              </ul>
            )}
            <p className="text-[11px] text-muted-foreground">
              This does not bypass SDTD's permissions or backend policies.
            </p>
            <div className="flex gap-2 pt-1">
              <button
                disabled={busy}
                onClick={() => decide(true)}
                className="flex-1 px-3 py-1.5 text-[11px] font-bold uppercase tracking-wider rounded bg-primary text-primary-foreground hover:brightness-110 disabled:opacity-60"
              >
                Approve
              </button>
              <button
                disabled={busy}
                onClick={() => decide(false)}
                className="flex-1 px-3 py-1.5 text-[11px] font-bold uppercase tracking-wider rounded border border-input bg-secondary text-secondary-foreground hover:opacity-80 disabled:opacity-60"
              >
                Cancel connection
              </button>
            </div>
          </>
        )}
      </div>
    </main>
  );
};

export default OAuthConsent;
