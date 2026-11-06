import { useState } from "react";
import { useSession } from "../app/session";
import { saveCompany } from "../data/api";
import type { Company } from "../data/types";
import {
  Button,
  Field,
  FormSection,
  Notice,
  PageHeader,
  Panel,
  PanelHeader,
  StatusLine,
  TextInput,
  useStatus,
} from "../ui/kit";

const blank: Company = {
  name: "",
  address: "",
  tin: "",
  telephone: "",
  sssNo: "",
  philhealthNo: "",
  pagibigNo: "",
};

export function CompanyScreen() {
  const { db, can } = useSession();
  const [draft, setDraft] = useState<Company>(db.company ?? blank);
  const [busy, setBusy] = useState(false);
  const [nameError, setNameError] = useState<string | null>(null);
  const { status, setStatus, clear } = useStatus();
  const editable = db.company ? can.edit : can.add;
  const set = (k: keyof Company, v: string) => setDraft((d) => ({ ...d, [k]: v }));

  const save = async () => {
    setNameError(null);
    if (!draft.name.trim()) {
      setNameError("Company name is required.");
      setStatus({ tone: "error", title: "Company not saved", body: "Company name is required." });
      return;
    }
    setBusy(true);
    try {
      await saveCompany(draft);
      setStatus({
        tone: "success",
        title: `Saved company ${draft.name}. Report headers now use it.`,
      });
    } catch (err) {
      setStatus({
        tone: "error",
        title: "Company not saved",
        body: err instanceof Error ? err.message : "The payroll database did not accept the save.",
      });
    } finally {
      setBusy(false);
    }
  };

  return (
    <>
      <PageHeader
        section="Administration"
        title="Company"
        description="The agency printed at the top of every report and payslip."
      />
      <div className="max-w-[760px] space-y-4">
        {!db.company && (
          <Notice tone="info" title="No company saved yet">
            Report headers stay blank until this is saved. Enter at least the company name.
          </Notice>
        )}
        <Panel>
          <PanelHeader
            title={db.company?.name || "Company details"}
            actions={
              editable && (
                <>
                  <Button
                    onClick={() => {
                      setDraft(db.company ?? blank);
                      setNameError(null);
                      clear();
                    }}
                    disabled={busy}
                  >
                    Cancel
                  </Button>
                  <Button variant="primary" onClick={save} busy={busy}>
                    Save
                  </Button>
                </>
              )
            }
          />
          {status && (
            <div className="px-4 pt-4">
              <StatusLine status={status} onDismiss={clear} />
            </div>
          )}
          <FormSection title="Identity" description="Shown on report headers.">
            <div className="grid gap-3 sm:grid-cols-2">
              <Field label="Company name" required error={nameError} className="sm:col-span-2">
                <TextInput
                  value={draft.name}
                  disabled={!editable}
                  onChange={(e) => set("name", e.target.value)}
                  placeholder="TAS Trans Corporation"
                />
              </Field>
              <Field label="Address" className="sm:col-span-2">
                <TextInput
                  value={draft.address}
                  disabled={!editable}
                  onChange={(e) => set("address", e.target.value)}
                />
              </Field>
              <Field label="TIN">
                <TextInput
                  value={draft.tin}
                  disabled={!editable}
                  onChange={(e) => set("tin", e.target.value)}
                />
              </Field>
              <Field label="Telephone">
                <TextInput
                  value={draft.telephone}
                  disabled={!editable}
                  onChange={(e) => set("telephone", e.target.value)}
                />
              </Field>
            </div>
          </FormSection>
          <FormSection
            title="Employer numbers"
            description="Printed on SSS, PhilHealth, and HDMF reports."
          >
            <div className="grid gap-3 sm:grid-cols-3">
              <Field label="SSS employer number">
                <TextInput
                  value={draft.sssNo}
                  disabled={!editable}
                  onChange={(e) => set("sssNo", e.target.value)}
                />
              </Field>
              <Field label="PhilHealth employer number">
                <TextInput
                  value={draft.philhealthNo}
                  disabled={!editable}
                  onChange={(e) => set("philhealthNo", e.target.value)}
                />
              </Field>
              <Field label="Pag-IBIG employer number">
                <TextInput
                  value={draft.pagibigNo}
                  disabled={!editable}
                  onChange={(e) => set("pagibigNo", e.target.value)}
                />
              </Field>
            </div>
          </FormSection>
        </Panel>
      </div>
    </>
  );
}
