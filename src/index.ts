import { ExampleAgent } from "./agents/exampleAgent";
import { AgentToolkit } from "./core/agentToolkit";
import { addTool } from "./tools/math/addTool";
import { echoTool } from "./tools/echoTool";

async function main(): Promise<void> {
  const toolkit = new AgentToolkit([echoTool, addTool]);
  const agent = new ExampleAgent(toolkit);

  console.log("Registered tools:", toolkit.listTools());

  console.log(await agent.handleUserMessage("hello tool framework"));
  console.log(await agent.handleUserMessage("add 12 30"));
}

main().catch((error) => {
  console.error("Fatal error", error);
});
