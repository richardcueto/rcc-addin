import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require, HTMLElement */

const title = "Richard Addin";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App title={title} />
    </FluentProvider>
  );
});

/* HMR (Hot Module Replacement) nativo de Vite */
if (import.meta.hot) {
  import.meta.hot.accept("./components/App", () => {
    // Vite recarga el componente automáticamente sin necesidad de re-renderizar manualmente
  });
}
