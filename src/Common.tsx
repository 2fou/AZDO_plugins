import React from "react";
import { createRoot } from "react-dom/client";
import "./Common.scss";

export function showRootComponent(component: React.ReactElement<any>) {
  const rootElement = document.getElementById("root");
  if (rootElement) {
    const root = createRoot(rootElement);
    root.render(component);
  }
}