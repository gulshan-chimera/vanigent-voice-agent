import { EntraAuthenticatedInvoker, EntraTokenProvider } from "../auth/entraAuthWrapper";
import { AgentToolkit } from "../core/agentToolkit";
import { ConsoleToolLogger } from "../core/logger";
import { ToolCallContext } from "../core/types";
import { mockSecureTool } from "../tools/mockSecureTool";

function assert(condition: unknown, message: string): void {
  if (!condition) {
    throw new Error(`Assertion failed: ${message}`);
  }
}

class MockEntraTokenProvider implements EntraTokenProvider {
  async getAccessToken(): Promise<{ accessToken: string; tokenType: "Bearer" }> {
    return {
      accessToken: "mock-entra-token",
      tokenType: "Bearer"
    };
  }
}

async function run(): Promise<void> {
  const toolkit = new AgentToolkit([mockSecureTool], new ConsoleToolLogger());
  const authInvoker = new EntraAuthenticatedInvoker(toolkit, {
    scopes: ["api://mock-service/.default"],
    tokenProvider: new MockEntraTokenProvider()
  });

  const context: ToolCallContext = {
    agentId: "test-agent",
    traceId: "trace-mock-001"
  };

  const success = await authInvoker.call<{ payload: string }, { payload: string; authHeader: string | null }>(
    "mock.secureEcho",
    { payload: "ping" },
    context
  );

  assert(success.ok === true, "Expected secure mock tool call to succeed");
  if (success.ok) {
    assert(success.requestId.length > 0, "Expected response to include requestId");
    assert(success.toolName === "mock.secureEcho", "Expected response to include tool name");
    assert(success.output.payload === "ping", "Expected payload echo in output");
    assert(
      success.output.authHeader === "Bearer mock-entra-token",
      "Expected Entra token to be injected in Authorization header"
    );
    assert(success.meta.agentId === "test-agent", "Expected agentId in metadata");
  }

  const validationFailure = await authInvoker.call(
    "mock.secureEcho",
    { payload: "" },
    context
  );

  assert(validationFailure.ok === false, "Expected validation failure for empty payload");
  if (!validationFailure.ok) {
    assert(
      validationFailure.error.code === "TOOL_VALIDATION_ERROR",
      "Expected TOOL_VALIDATION_ERROR code"
    );
  }

  const missingTool = await authInvoker.call(
    "mock.missing",
    { payload: "ping" },
    context
  );

  assert(missingTool.ok === false, "Expected missing tool invocation to fail");
  if (!missingTool.ok) {
    assert(missingTool.error.code === "TOOL_NOT_FOUND", "Expected TOOL_NOT_FOUND code");
  }

  console.log("Toolkit mock test passed");
}

run().catch((error) => {
  console.error("Toolkit mock test failed", error);
  throw error;
});
