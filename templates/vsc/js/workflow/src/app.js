const { App } = require("@microsoft/teams.apps");
const { DoStuffActionHandler } = require("./cardActions/doStuffActionHandler");
const { GenericCommandHandler } = require("./commands/genericCommandHandler");
const { HelloWorldCommandHandler } = require("./commands/helloworldCommandHandler");

// Create the app with logger
const app = new App();

// Initialize command handlers
const helloworldCommandHandler = new HelloWorldCommandHandler();
const genericCommandHandler = new GenericCommandHandler();
const doStuffActionHandler = new DoStuffActionHandler();

app.on("message", async ({ activity, send }) => {
  if (helloworldCommandHandler.shouldTrigger(activity.text)) {
    const response = await helloworldCommandHandler.handleCommandReceived(activity);
    if (response) {
      await send(response);
    }
    return;
  }

  const response = await genericCommandHandler.handleCommandReceived(activity);
  if (response) {
    await send(response);
  }
});

// Handle adaptive card actions
app.on("card.action", async ({ activity, send }) => {
  const verb = activity.value?.action?.verb;
  if (verb === doStuffActionHandler.triggerVerb) {
    const response = await doStuffActionHandler.handleActionInvoked();
    await send(response);
  } else {
    return {
      statusCode: 400,
      type: "application/vnd.microsoft.error",
      value: {
        code: "BadRequest",
        message: "Unknown action",
        innerHttpError: {
          statusCode: 400,
          body: { error: "Unknown action" },
        },
      },
    };
  }
});

// Export the app
module.exports = { app };
