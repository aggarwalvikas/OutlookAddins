import * as React from "react";
import * as ReactDOM from "react-dom";
import App from "./components/App";

/* global document, Office, module, require, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");

/* Render application after Office initializes */
Office.onReady(() => {
  if (rootElement) {
    ReactDOM.render(<App />, rootElement);
  }
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    if (rootElement) {
      ReactDOM.render(<NextApp />, rootElement);
    }
  });
}
