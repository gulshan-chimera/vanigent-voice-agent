import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { handleAvailability } from "../lib/availabilityRequest";
export function checkMeetingAvailability(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  return handleAvailability(request, context);
}
app.http("checkMeetingAvailability", {
  methods: ["POST"], authLevel: "anonymous", route: "checkMeetingAvailability", handler: checkMeetingAvailability,
});
