import { useState } from "react";
import { useSession } from "../app/session";
import { deactivateLoginLock, saveRow } from "../data/api";
import type { User } from "../data/types";
import { ACCOUNT_TYPES, RIGHTS } from "../data/types";
import { uid } from "../data/money";
import {
  Button,
  Checkbox,
  DataTable,
  EmptyState,
  Field,
  FormSection,
  Notice,
  PageHeader,
  Panel,
  PanelHeader,
  Select,
  StatusLine,
  Tag,
  TextInput,
  useConfirm,
  useStatus,
} from "../ui/kit";
import { fmtDateTime } from "../ui/format";

export function Users() {
  const { db, can, user: me } = useSession();
  const confirm = useConfirm();
  const [draft, setDraft] = useState<User | null>(null);
  const [isNew, setIsNew] = useState(false);
  const [password, setPassword] = useState("");
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [busy, setBusy] = useState<string | null>(null);
  const { status, setStatus, clear } = useStatus();
  const saved = draft ? db.users.find((u) => u.id === draft.id) : undefined;
  const editable = isNew ? can.add : can.edit;

  const open = (u: User) => {
    setDraft(structuredClone(u));
    setIsNew(false);
    setPassword("");
    setErrors({});
    clear();
  };
  const startNew = () => {
    setDraft({
      id: `u${uid()}`,
      username: "",
      password: "",
      name: "",
      accountType: "Staff Payroll",
      superuser: false,
      rights: Object.fromEntries(RIGHTS.map((r) => [r.key, false])) as User["rights"],
      canAdd: false,
      canEdit: false,
      canDelete: false,
      activeLogin: false,
      lastLoginAt: null,
      lastLogoutAt: null,
    });
    setIsNew(true);
    setPassword("");
    setErrors({});
    clear();
  };
  const save = async () => {
    if (!draft) return;
    const errs: Record<string, string> = {};
    if (!draft.username.trim()) errs.username = "Username is required.";
    else if (
      db.users.some(
        (u) =>
          u.username.toLowerCase() === draft.username.trim().toLowerCase() && u.id !== draft.id,
      )
    )
      errs.username = `Username ${draft.username} is already used.`;
    if (isNew && password.length < 6) errs.password = "Password needs at least 6 characters.";
    if (!isNew && password && password.length < 6)
      errs.password = "Password needs at least 6 characters.";
    setErrors(errs);
    if (Object.keys(errs).length) {
      setStatus({ tone: "error", title: "User not saved", body: Object.values(errs)[0] });
      return;
    }
    setBusy("save");
    const next = {
      ...draft,
      username: draft.username.trim(),
      name: draft.name || draft.username.trim(),
      password: password || saved?.password || "",
    };
    await saveRow("users", next, "user");
    setBusy(null);
    setIsNew(false);
    setPassword("");
    setDraft(structuredClone(next));
    setStatus({ tone: "success", title: `Saved user ${next.username}.` });
  };
  const unlock = async () => {
    if (!saved) return;
    const ok = await confirm({
      title: "Deactivate login lock",
      body: (
        <>
          Clear the active login for <b>{saved.username}</b>? Their open session can no longer save.
        </>
      ),
      confirmLabel: "Deactivate login lock",
    });
    if (!ok) return;
    setBusy("unlock");
    await deactivateLoginLock(saved.id);
    setBusy(null);
    setDraft((d) => (d ? { ...d, activeLogin: false } : d));
    setStatus({ tone: "success", title: `Cleared the login lock for ${saved.username}.` });
  };
  const set = <K extends keyof User>(k: K, v: User[K]) =>
    setDraft((d) => (d ? { ...d, [k]: v } : d));

  return (
    <>
      <PageHeader
        section="Administration"
        title="Users"
        description="Accounts, rights, and login locks. Passwords are entered here and never shown back."
      />
      <div className="grid items-start gap-4 xl:grid-cols-[minmax(0,1fr)_minmax(420px,1fr)]">
        <Panel>
          <PanelHeader
            title="Users"
            meta={`${db.users.length} accounts`}
            actions={
              can.add && (
                <Button onClick={startNew} disabled={!!busy}>
                  New
                </Button>
              )
            }
          />
          <DataTable
            rows={db.users}
            rowKey={(u) => u.id}
            selected={draft?.id}
            onSelect={open}
            empty={
              <EmptyState action={can.add ? <Button onClick={startNew}>New</Button> : undefined}>
                No user accounts yet.
              </EmptyState>
            }
            columns={[
              {
                key: "u",
                header: "Username",
                render: (u) => <span className="font-medium">{u.username}</span>,
              },
              { key: "t", header: "Account type", render: (u) => u.accountType },
              {
                key: "a",
                header: "Active login",
                render: (u) =>
                  u.activeLogin ? (
                    <Tag tone="blue">Logged in{u.id === me.id ? " (you)" : ""}</Tag>
                  ) : (
                    <span className="text-grey-400">No</span>
                  ),
              },
              {
                key: "s",
                header: "Superuser",
                render: (u) => (u.superuser ? "Yes" : <span className="text-grey-400">No</span>),
              },
              {
                key: "l",
                header: "Last login",
                render: (u) => <span className="text-grey-600">{fmtDateTime(u.lastLoginAt)}</span>,
              },
            ]}
          />
        </Panel>

        <Panel>
          {!draft ? (
            <>
              <PanelHeader title="No user open" />
              <div className="space-y-3 px-4 py-4">
                <StatusLine status={status} onDismiss={clear} />
                <p className="text-sm text-grey-600">
                  Choose a user to edit rights{can.add ? ", or use New" : ""}.
                </p>
              </div>
            </>
          ) : (
            <>
              <PanelHeader
                title={isNew ? "New user" : saved?.username}
                actions={
                  <>
                    {!isNew && saved?.activeLogin && saved.id !== me.id && can.edit && (
                      <Button onClick={unlock} busy={busy === "unlock"} disabled={!!busy}>
                        Deactivate login lock
                      </Button>
                    )}
                    {editable && (
                      <>
                        <Button
                          onClick={() => (isNew ? setDraft(null) : saved && open(saved))}
                          disabled={!!busy}
                        >
                          Cancel
                        </Button>
                        <Button
                          variant="primary"
                          onClick={save}
                          busy={busy === "save"}
                          disabled={!!busy}
                        >
                          Save
                        </Button>
                      </>
                    )}
                  </>
                }
              />
              {(status || !editable) && (
                <div className="space-y-2 px-4 pt-4">
                  <StatusLine status={status} onDismiss={clear} />
                  {!editable && (
                    <Notice tone="neutral" title="Read only">
                      Your account cannot edit users.
                    </Notice>
                  )}
                </div>
              )}
              <fieldset disabled={!editable} className="contents">
                <FormSection title="Account">
                  <div className="grid gap-3 sm:grid-cols-2">
                    <Field label="Username" required error={errors.username}>
                      <TextInput
                        value={draft.username}
                        autoComplete="off"
                        onChange={(e) => set("username", e.target.value)}
                      />
                    </Field>
                    <Field
                      label="Password"
                      required={isNew}
                      error={errors.password}
                      hint={isNew ? undefined : "Leave blank to keep the current password."}
                    >
                      <TextInput
                        type="password"
                        autoComplete="new-password"
                        value={password}
                        onChange={(e) => setPassword(e.target.value)}
                      />
                    </Field>
                    <Field
                      label="Account type"
                      hint="FMS types see only FMS clients. Staff types see only staff clients."
                    >
                      <Select
                        value={draft.accountType}
                        onChange={(e) => set("accountType", e.target.value as User["accountType"])}
                      >
                        {ACCOUNT_TYPES.map((t) => (
                          <option key={t}>{t}</option>
                        ))}
                      </Select>
                    </Field>
                    <div className="flex items-end pb-1.5">
                      <Checkbox
                        checked={draft.superuser}
                        disabled={!editable}
                        onChange={(v) => set("superuser", v)}
                        label="Superuser"
                        hint="Sees every area."
                      />
                    </div>
                  </div>
                </FormSection>
                <FormSection
                  title="Rights"
                  description={
                    draft.superuser
                      ? "A superuser sees every area; these rights are ignored."
                      : "Each area shows only when its right is on."
                  }
                >
                  <div className="divide-y divide-grey-100 rounded-ctl border border-grey-100">
                    {RIGHTS.map((r) => (
                      <div key={r.key} className="px-3 py-2">
                        <Checkbox
                          checked={draft.rights[r.key]}
                          disabled={!editable || draft.superuser}
                          onChange={(v) => set("rights", { ...draft.rights, [r.key]: v })}
                          label={r.label}
                          hint={r.areas}
                        />
                      </div>
                    ))}
                  </div>
                </FormSection>
                <FormSection
                  title="Record actions"
                  description="Separate rights. A button is hidden when its right is off."
                >
                  <div className="flex flex-wrap gap-6">
                    <Checkbox
                      checked={draft.canAdd}
                      disabled={!editable || draft.superuser}
                      onChange={(v) => set("canAdd", v)}
                      label="Add"
                    />
                    <Checkbox
                      checked={draft.canEdit}
                      disabled={!editable || draft.superuser}
                      onChange={(v) => set("canEdit", v)}
                      label="Edit"
                    />
                    <Checkbox
                      checked={draft.canDelete}
                      disabled={!editable || draft.superuser}
                      onChange={(v) => set("canDelete", v)}
                      label="Delete"
                    />
                  </div>
                </FormSection>
              </fieldset>
            </>
          )}
        </Panel>
      </div>
    </>
  );
}
