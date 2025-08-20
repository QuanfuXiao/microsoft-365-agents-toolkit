// Import the Teams AI v2 app
import app from "./app/app";

// Start the application - the App class handles the server hosting
(async () => {
  await app.start(process.env.PORT || process.env.port || 3978);
  console.log(`\nAgent started, app listening to`, process.env.PORT || process.env.port || 3978);
})();
