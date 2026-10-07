import { app, HttpRequest, HttpResponseInit } from "@azure/functions";

export async function healthCheck(
  request: HttpRequest,
): Promise<HttpResponseInit> {
  return {
    status: 200,
    jsonBody: {
      status: "ok",
      service: "meeting-schedular",
      timestamp: new Date().toISOString(),
    },
  };
}

app.http("healthCheck", {
  methods: ["GET"],
  authLevel: "function",
  handler: healthCheck,
});
