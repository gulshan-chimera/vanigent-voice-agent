import { app, InvocationContext, Timer } from "@azure/functions";

export async function keepWarm(timer: Timer, context: InvocationContext): Promise<void> {
  context.log("[KEEP-WARM] Ping.");
}

app.timer("keepWarm", {
  schedule: "%KEEP_WARM_SCHEDULE%",
  runOnStartup: false,
  handler: keepWarm,
});