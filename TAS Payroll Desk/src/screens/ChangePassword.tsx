import { useState, type FormEvent } from "react";
import { ApiError, changePassword, setFlash } from "../data/api";
import { navigate, previousRoute } from "../app/router";
import {
  Button,
  Field,
  PageHeader,
  Panel,
  PanelHeader,
  StatusLine,
  TextInput,
  useStatus,
} from "../ui/kit";

export function ChangePassword() {
  const [current, setCurrent] = useState("");
  const [next, setNext] = useState("");
  const [confirm, setConfirm] = useState("");
  const [field, setField] = useState<string | undefined>();
  const [busy, setBusy] = useState(false);
  const { status, setStatus, clear } = useStatus();

  const submit = async (e: FormEvent) => {
    e.preventDefault();
    setBusy(true);
    setField(undefined);
    try {
      await changePassword(current, next, confirm);
      setFlash({ tone: "success", title: "Password changed." });
      navigate(previousRoute());
    } catch (err) {
      const f = err instanceof ApiError ? err.field : undefined;
      setField(f);
      setStatus({ tone: "error", title: "Password not changed", body: (err as Error).message });
      if (f === "current") setCurrent("");
    } finally {
      setBusy(false);
    }
  };
  const msg = status?.body as string | undefined;

  return (
    <>
      <PageHeader
        title="Change password"
        description="Your new password takes effect the next time you log in."
      />
      <form onSubmit={submit} className="max-w-[440px]">
        <Panel>
          <PanelHeader
            title="Password"
            actions={
              <>
                <Button onClick={() => navigate(previousRoute())} disabled={busy}>
                  Cancel
                </Button>
                <Button type="submit" variant="primary" busy={busy}>
                  Save
                </Button>
              </>
            }
          />
          <div className="space-y-4 px-4 py-4">
            <StatusLine status={status} onDismiss={clear} />
            <Field label="Current password" required error={field === "current" ? msg : null}>
              <TextInput
                type="password"
                autoComplete="current-password"
                value={current}
                onChange={(e) => setCurrent(e.target.value)}
                autoFocus
              />
            </Field>
            <Field
              label="New password"
              required
              hint="At least 6 characters."
              error={field === "next" ? msg : null}
            >
              <TextInput
                type="password"
                autoComplete="new-password"
                value={next}
                onChange={(e) => setNext(e.target.value)}
              />
            </Field>
            <Field label="Confirm new password" required error={field === "confirm" ? msg : null}>
              <TextInput
                type="password"
                autoComplete="new-password"
                value={confirm}
                onChange={(e) => setConfirm(e.target.value)}
              />
            </Field>
          </div>
        </Panel>
      </form>
    </>
  );
}
