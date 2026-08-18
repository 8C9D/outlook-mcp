// Interactive login: runs the device-code flow (if needed) and exits after
// caching tokens. This is the ONLY entry point that may prompt for sign-in;
// the MCP server itself never initiates authentication.
import { getAccessToken } from "./auth.js";
import { callGraph } from "./graph.js";

try {
  await getAccessToken();
  const me = await callGraph("/me?$select=displayName,userPrincipalName");
  console.error(`Signed in as ${me.displayName} <${me.userPrincipalName}>. Tokens cached.`);
} catch (err) {
  console.error(`Login failed: ${err instanceof Error ? err.message : String(err)}`);
  process.exitCode = 1;
}
