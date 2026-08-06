# Agent Toolkit: Generic Tool-Calling Framework (Node.js)

This project provides a reusable, typed tool-calling framework for agent systems.

## What it gives you

- Common `ToolDefinition` interface for every tool
- Central `ToolRegistry` for registration and discovery
- Runtime validation hooks per tool
- Unified `ToolExecutor` with normalized execution results
- Shared `AgentToolkit` facade for all agents
- Standardized request/response schema for tool invocations
- Structured logging for start/success/failure events
- Reusable Entra token injection wrapper for authenticated calls

## Structure

```
src/
  core/
    types.ts
    errors.ts
    logger.ts
    toolRegistry.ts
    toolExecutor.ts
    agentToolkit.ts
  auth/
    entraAuthWrapper.ts
  tools/
    echoTool.ts
    math/
      addTool.ts
    mockSecureTool.ts
  agents/
    exampleAgent.ts
  tests/
    mockTool.test.ts
  index.ts
```

## Quick start

1. Install dependencies:

```bash
npm install
```

2. Build:

```bash
npm run build
```

3. Run demo:

```bash
npm start
```

4. Run toolkit test with mock tool:

```bash
npm test
```

## Core concepts

### `ToolDefinition<TInput, TOutput>`

Every tool implements:

- `name`: globally unique tool name
- `description`: plain language description
- `validate(input)`: returns validated input or throws
- `execute(input, context)`: performs business logic

### `AgentToolkit`

Agents use one entrypoint:

```ts
const result = await toolkit.call("math.add", { a: 2, b: 3 }, context);
```

This keeps all agents consistent and simplifies tracing, auditing, and testing.

### Standard request schema

`ToolCallRequest<TInput>` defines the invocation contract:

- `requestId`
- `toolName`
- `input`
- `context` (`agentId`, `traceId`, metadata, auth)
- `createdAt`

### Standard response schema

`ToolCallResponse<TOutput>` is always normalized:

- Success: `ok`, `requestId`, `toolName`, `output`, `meta`
- Failure: `ok`, `requestId`, `toolName`, `error`, `meta`

Metadata includes timing and trace fields for observability.

### Logging and error handling

`ToolExecutor` emits structured logs for:

- `tool.call.started`
- `tool.call.succeeded`
- `tool.call.failed`

All errors are normalized to stable error codes (`TOOL_NOT_FOUND`, `TOOL_VALIDATION_ERROR`, `TOOL_EXECUTION_ERROR`, `TOOL_AUTHENTICATION_ERROR`).

### Entra auth wrapper

Use `EntraAuthenticatedInvoker` to inject bearer tokens into all calls:

```ts
const authInvoker = new EntraAuthenticatedInvoker(toolkit, {
  scopes: ["api://my-service/.default"],
  tokenProvider
});
```

The wrapper injects:

- `context.auth.entra.accessToken`
- `context.metadata.headers.Authorization`
