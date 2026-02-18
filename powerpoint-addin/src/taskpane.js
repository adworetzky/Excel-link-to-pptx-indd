/**
 * powerpoint-addin/src/taskpane.js
 *
 * Office JS entry point. Waits for Office.onReady() before mounting React.
 * All Office JS APIs must be called after this resolves.
 */
import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App.jsx";

Office.onReady(() => {
  const container = document.getElementById("root");
  if (container) {
    const root = createRoot(container);
    root.render(<App />);
    // Notify the App component that Office is ready
    setTimeout(() => {
      if (typeof window.__datalinkSetReady === "function") {
        window.__datalinkSetReady();
      }
    }, 0);
  }
});
