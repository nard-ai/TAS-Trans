import { useDb, useDbError, useDbStatus, useSaveError } from "./data/store";
import { SessionProvider, useSession } from "./app/session";
import { Shell } from "./app/Shell";
import { useRoute } from "./app/router";
import { resolveRoute } from "./app/routes";
import { Login } from "./screens/Login";
import { ConfirmProvider, EmptyState, PageHeader, Panel } from "./ui/kit";

export default function App() {
  const status = useDbStatus();
  const db = useDb();
  const problem = useSaveError();
  const errorText = useDbError();
  if (status === "loading") return <Gate>Loading the payroll desk</Gate>;
  if (status === "error") return <Gate>{errorText}</Gate>;
  return (
    <ConfirmProvider>
      {problem ? (
        <div className="border-b border-red-200 bg-red-50 px-4 py-2 text-center text-sm text-red-800">
          {problem}
        </div>
      ) : null}
      {db.session ? (
        <SessionProvider>
          <Routed />
        </SessionProvider>
      ) : (
        <Login />
      )}
    </ConfirmProvider>
  );
}

function Gate({ children }: { children: string }) {
  return (
    <div className="grid min-h-screen place-items-center bg-grey-50 px-6 text-center text-sm text-grey-700">
      {children}
    </div>
  );
}

function Routed() {
  const { user } = useSession();
  const { path, param } = useRoute();
  const { match, exists } = resolveRoute(path, param, user);
  const El = match?.element;
  return (
    <Shell>
      {El ? (
        <El key={`${path}/${param ?? ""}`} param={param} />
      ) : (
        <>
          <PageHeader title={exists ? "Not available to your account" : "Page not found"} />
          <Panel>
            <EmptyState
              action={
                <a className="text-sm font-medium text-blue-700" href="#/home">
                  Go to work queue
                </a>
              }
            >
              {exists
                ? "This area needs a right your account does not have. Ask an administrator to turn it on in Users."
                : "There is no screen at this address."}
            </EmptyState>
          </Panel>
        </>
      )}
    </Shell>
  );
}
