import {
  app,
  HttpRequest,
  HttpResponseInit,
  InvocationContext,
} from "@azure/functions";
import { handleMeetingTool } from "../lib/vapiMeetingTool";

export function scheduleMeeting(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  return handleMeetingTool(request, context, "schedule_meeting");
}

app.http("scheduleMeeting", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "scheduleMeeting",
  handler: scheduleMeeting,
});
