import { useEffect, useRef, useState, type FormEvent } from "react";
import { useDb } from "../data/store";
import { deactivateUser, login } from "../data/api";
import { navigate } from "../app/router";
import { Brand } from "../app/Shell";
import { Button, Checkbox, Field, Notice, Select, TextInput } from "../ui/kit";

type State = "idle" | "failed" | "active" | "ended";
const SAVED = "tas-hr-saved-login";

export function Login() {
  const db = useDb();
  const saved = (() => {
    try {
      return JSON.parse(localStorage.getItem(SAVED) ?? "null") as {
        agency: string;
        username: string;
        password: string;
      } | null;
    } catch {
      return null;
    }
  })();
  const [agency, setAgency] = useState(saved?.agency ?? "");
  const [username, setUsername] = useState(saved?.username ?? "");
  const [password, setPassword] = useState(saved?.password ?? "");
  const [remember, setRemember] = useState(!!saved);
  const [state, setState] = useState<State>("idle");
  const [failures, setFailures] = useState(0);
  const [busy, setBusy] = useState<"ok" | "deactivate" | null>(null);
  const [error, setError] = useState<string | null>(null);
  const userRef = useRef<HTMLInputElement>(null);

  const chosen = db.agencies.find((a) => a.code === agency);
  const noDb = !!chosen && !chosen.database;
  const noAgencies = db.agencies.length === 0;
  const ended = state === "ended";
  const blocked = !agency || noDb || noAgencies || ended;

  useEffect(() => {
    userRef.current?.focus();
  }, []);

  const finish = () => {
    if (remember) localStorage.setItem(SAVED, JSON.stringify({ agency, username, password }));
    else localStorage.removeItem(SAVED);
    navigate("home");
  };

  const submit = async (e?: FormEvent) => {
    e?.preventDefault();
    if (blocked || busy) return;
    setError(null);
    if (!username.trim()) {
      setError("Enter a username.");
      userRef.current?.focus();
      return;
    }
    setBusy("ok");
    try {
      const res = await login(agency, username, password);
      if (res.status === "ok") return finish();
      if (res.status === "failed") {
        setState("failed");
        setFailures(res.failures);
        setPassword("");
        userRef.current?.focus();
        userRef.current?.select();
      } else if (res.status === "ended") setState("ended");
      else if (res.status === "active") setState("active");
    } catch (err) {
      setError(err instanceof Error ? err.message : "The payroll database did not answer.");
    } finally {
      setBusy(null);
    }
  };

  const onDeactivate = async () => {
    setBusy("deactivate");
    try {
      const res = await deactivateUser(agency, username, password);
      if (res.status === "ok") return finish();
      setState("active");
    } catch (err) {
      setError(err instanceof Error ? err.message : "The payroll database did not answer.");
    } finally {
      setBusy(null);
    }
  };

  const onCancel = () => {
    setAgency("");
    setUsername("");
    setPassword("");
    setRemember(false);
    setError(null);
    if (state !== "ended") setState("idle");
    userRef.current?.focus();
  };

  return (
    <div className="flex min-h-screen flex-col bg-grey-50">
      <div className="flex h-14 items-center border-b border-grey-100 bg-white px-4 sm:px-6">
        <Brand />
      </div>

      <main className="flex flex-1 items-start justify-center px-4 pt-[10vh]">
        <div className="w-full max-w-[400px]">
          <form
            onSubmit={submit}
            noValidate
            className="rounded-panel border border-grey-100 bg-white shadow-pop"
            aria-labelledby="login-title"
          >
            <div className="border-b border-grey-100 px-4 py-4 sm:px-6">
              <h1 id="login-title" className="text-lg font-semibold text-grey-900">
                Log in
              </h1>
              <p className="mt-0.5 text-sm text-grey-600">
                Use the account issued by your administrator.
              </p>
            </div>

            <div className="space-y-4 px-4 py-5 sm:px-6">
              {state === "failed" && (
                <Notice tone="error" title="Login failed.">
                  Check the username and password.{" "}
                  {failures >= 3
                    ? "One more failure ends this session."
                    : `${failures} of 3 attempts used.`}
                </Notice>
              )}
              {state === "ended" && (
                <Notice tone="error" title="Too many Invalid Logon Attempts">
                  This session has ended. Close this window and open HR and Payroll again to retry.
                </Notice>
              )}
              {state === "active" && (
                <Notice tone="info" title={`Someone is still logged in as ${username}.`}>
                  This account is active in another session. Deactivate user clears that login lock
                  so you can continue here.
                </Notice>
              )}
              {(noDb || noAgencies) && (
                <Notice tone="error" title="Database setup needed">
                  {noAgencies
                    ? "No agency is configured."
                    : `${chosen?.name} has no database set up.`}{" "}
                  Ask your system administrator to set up the database for this agency, then log in
                  again.
                </Notice>
              )}

              <Field
                label="Agency"
                required
                hint={chosen && !noDb ? undefined : "Choose the company you are working for."}
              >
                <Select
                  value={agency}
                  disabled={ended || !!busy}
                  onChange={(e) => setAgency(e.target.value)}
                >
                  <option value="">Choose an agency</option>
                  {db.agencies.map((a) => (
                    <option key={a.code} value={a.code}>
                      {a.code}
                    </option>
                  ))}
                </Select>
              </Field>
              {chosen && !noDb && (
                <div className="-mt-2 flex items-center justify-between rounded-ctl bg-blue-50 px-2.5 py-1.5 text-sm">
                  <span className="font-medium text-blue-900">{chosen.name}</span>
                  <span className="text-xs text-blue-700">{chosen.database}</span>
                </div>
              )}

              <Field label="Username" required error={error}>
                <TextInput
                  ref={userRef}
                  autoComplete="username"
                  value={username}
                  disabled={ended || !!busy}
                  onChange={(e) => setUsername(e.target.value)}
                />
              </Field>
              <Field label="Password" required>
                <TextInput
                  type="password"
                  autoComplete="current-password"
                  value={password}
                  disabled={ended || !!busy}
                  onChange={(e) => setPassword(e.target.value)}
                />
              </Field>
              <Checkbox
                checked={remember}
                disabled={ended}
                onChange={setRemember}
                label="Save password"
                hint="Only on a computer that is yours alone."
              />
            </div>

            <div className="flex flex-wrap items-center gap-2 rounded-b-panel border-t border-grey-100 bg-grey-50 px-4 py-3 sm:px-6">
              {state === "active" && (
                <Button onClick={onDeactivate} busy={busy === "deactivate"} disabled={!!busy}>
                  Deactivate user
                </Button>
              )}
              <div className="ml-auto flex gap-2">
                <Button onClick={onCancel} disabled={!!busy}>
                  Cancel
                </Button>
                <Button
                  type="submit"
                  variant="primary"
                  className="min-w-20"
                  busy={busy === "ok"}
                  disabled={blocked || !!busy}
                >
                  Ok
                </Button>
              </div>
            </div>
          </form>

          <p className="mt-4 px-1 text-xs leading-5 text-grey-500">
            Prototype accounts, password <span className="font-medium text-grey-700">tas123</span>:
            admin (superuser), staffpay, fmspay, lpanganiban (Accounting), jbernardo (already logged
            in elsewhere). Agency TAS03 has no database.
          </p>
        </div>
      </main>
    </div>
  );
}
