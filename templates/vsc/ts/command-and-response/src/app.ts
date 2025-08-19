import { App } from "@microsoft/teams.apps";
import { HelloWorldCommandHandler } from "./helloworldCommandHandler";
import { GenericCommandHandler } from "./genericCommandHandler";

export const app = new App({});

// Initialize command handlers
const helloWorldHandler = new HelloWorldCommandHandler();
const genericHandler = new GenericCommandHandler();

// Register message handler
app.on("message", async ({ activity, send }) => {
  const text = activity.text || "";

  // Check if helloWorld command
  if (helloWorldHandler.canHandle(text)) {
    const reply = await helloWorldHandler.handleCommandReceived(activity);
    if (reply) {
      await send(reply);
    }
    return;
  }

  // Handle all other messages with generic handler
  const reply = await genericHandler.handleCommandReceived(activity);
  if (reply) {
    await send(reply);
  }
});

export default app;
