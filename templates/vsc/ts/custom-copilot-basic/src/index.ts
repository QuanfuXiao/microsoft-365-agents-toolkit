// Import the Teams AI v2 app
import app from "./app";

// Start the application - the App class handles the server hosting
(async () => {
  await app.start(process.env.port || process.env.PORT || 3978);
  console.log(`\nAgent started, app listening to`, process.env.port || process.env.PORT || 3978);
})();
