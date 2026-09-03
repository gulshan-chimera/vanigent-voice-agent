// src/tests/callReport_test.ts
//
// End-to-end test for the follow-up email flow: POSTs a synthetic
// "end-of-call-report" event — shaped exactly like vapiCallReport.ts
// expects — at a running Function App (local `func start` by default,
// or any URL you point it at). Exercises the whole path: webhook auth,
// structured-output parsing, caller lookup, and both Graph sendMail
// calls — not just the Graph permission in isolation.
//
// Usage:
//   npm run build
//   func start                    (in a separate terminal)
//   npm run test:callreport
// (or: node dist/tests/callReport_test.js)

import * as fs from "fs";
import * as path from "path";
import * as readline from "readline";

// Falls back to local.settings.json (the file `func start` normally reads)
// or a .env file if VAPI_REPORT_SECRET isn't already in the environment —
// needed because this script runs as a plain node process, not through
// the Azure Functions host. Same approach as auth_test.ts.
function loadEnvFileIfNeeded(): void {
  if (process.env.VAPI_REPORT_SECRET) {
    return;
  }

  const root = path.join(__dirname, "..", "..");
  loadFromLocalSettingsJson(path.join(root, "local.settings.json"));
  loadFromDotEnv(path.join(root, ".env"));
}

function loadFromLocalSettingsJson(settingsPath: string): void {
  if (!fs.existsSync(settingsPath)) return;

  const parsed = JSON.parse(fs.readFileSync(settingsPath, "utf-8")) as {
    Values?: Record<string, string>;
  };

  for (const [key, value] of Object.entries(parsed.Values ?? {})) {
    if (key && !process.env[key]) process.env[key] = value;
  }
}

function loadFromDotEnv(envPath: string): void {
  if (!fs.existsSync(envPath)) return;

  const lines = fs.readFileSync(envPath, "utf-8").split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eq = trimmed.indexOf("=");
    if (eq === -1) continue;
    const key = trimmed.slice(0, eq).trim();
    const value = trimmed.slice(eq + 1).trim();
    if (key && !process.env[key]) process.env[key] = value;
  }
}

/** Builds a payload matching EndOfCallReportBody in vapiCallReport.ts exactly. */
function buildPayload(callerNumber: string, followUpRequired: boolean): unknown {
  return {
    message: {
      type: "end-of-call-report",
      startedAt: new Date().toISOString(),
      durationSeconds: 42,
      transcript: "Caller asked about their benefits enrollment date.",
      analysis: {
        structuredOutputs: {
          "test-output-id": {
            name: "followUpReport",
            result: {
              summary: "Caller asked when benefits begin; not covered in the KB.",
              urgency: "medium",
              customerName: "Test Caller",
              openQuestions: followUpRequired ? ["When do benefits begin?"] : [],
              followUpRequired,
            },
          },
        },
      },
      call: {
        id: `test-call-${Date.now()}`,
        customer: { number: callerNumber || undefined },
      },
    },
  };
}

async function sendTestReport(
  baseUrl: string,
  callerNumber: string,
  followUpRequired: boolean
): Promise<void> {
  const secret = process.env.VAPI_REPORT_SECRET;
  if (!secret) {
    console.error("VAPI_REPORT_SECRET is not set — check local.settings.json or .env.");
    return;
  }

  const url = `${baseUrl.replace(/\/$/, "")}/api/vapiCallReport`;
  const payload = buildPayload(callerNumber, followUpRequired);

  console.log(`\nPOST ${url}`);
  console.log(`followUpRequired=${followUpRequired}, callerNumber=${callerNumber || "(none)"}`);

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${secret}`,
      },
      body: JSON.stringify(payload),
    });

    const text = await response.text();
    console.log(`Status: ${response.status}`);
    console.log(`Body: ${text || "(empty)"}`);

    if (response.status === 200) {
      console.log(
        followUpRequired
          ? "\nCheck the function's logs for [CALL-REPORT]/[GRAPH-MAIL] lines, and check the support/caller inboxes."
          : "\nfollowUpRequired was false, so no email should have been sent — this just confirms the endpoint accepted the event."
      );
    }
  } catch (error) {
    console.error(`Request failed: ${(error as Error).message}`);
    console.error("Is the Function App actually running (func start) at that URL?");
  }
}

function main(): void {
  loadEnvFileIfNeeded();

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  console.log("vapiCallReport end-to-end test — sends a synthetic end-of-call-report event.\n");

  rl.question("Function App base URL [http://localhost:7071]: ", (baseUrlAnswer) => {
    const baseUrl = baseUrlAnswer.trim() || "http://localhost:7071";

    console.log("\nType a phone number to also test caller-email resolution, or 'exit' to quit.\n");

    const ask = (): void => {
      rl.question("Caller phone number (blank = support-only test): ", async (answer) => {
        const trimmed = answer.trim();
        if (trimmed.toLowerCase() === "exit" || trimmed.toLowerCase() === "quit") {
          rl.close();
          return;
        }
        await sendTestReport(baseUrl, trimmed, true);
        ask();
      });
    };

    ask();
  });
}

main();
