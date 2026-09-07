import * as React from "react";
import * as ReactDOM from "react-dom";
import App from "./components/App";

/* global document, Office, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");

/* Render application after Office initializes */
Office.onReady(() => {
  if (rootElement) {
    ReactDOM.render(<App />, rootElement);
  }
});
