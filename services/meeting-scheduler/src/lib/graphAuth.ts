// Reuse the existing cached app-only provider; do not duplicate OAuth logic.
// The scheduler build includes this dependency in its own deployment output.
// Configuration is read from this Function App's environment, not the KB app.
// A null token means authentication failed: consumers must fail closed.
export { getGraphAccessToken } from "../../../Knowledge Base/src/lib/graphAuth";
