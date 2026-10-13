import React from "react";
import { createRoot } from "react-dom/client";
import { App } from "./components/App";
import "./taskpane.css";

/* global document, Office */

const container = document.getElementById("container");
if (container) {
  const root = createRoot(container);
  Office.onReady(() => {
    root.render(<App />);
  });
}
