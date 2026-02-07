import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import "@mantine/core/styles.css";
import "@mantine/notifications/styles.css";
import "./index.css";

import { MantineProvider, createTheme } from '@mantine/core';
import { Notifications } from "@mantine/notifications";

import "@fontsource/noto-sans-sc/400.css";
import "@fontsource/noto-sans-sc/500.css";
import "@fontsource/noto-sans-sc/600.css";

import "./index.css";

const theme = createTheme({
  primaryColor: 'blue',
  defaultRadius: 'md',
  fontFamily:
    'Noto Sans SC, ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, "Apple Color Emoji","Segoe UI Emoji"',
  headings: {
    fontFamily:
      'Noto Sans SC, ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial',
    fontWeight: '600',
  },
});

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <MantineProvider theme={theme}>
      <Notifications position="top-right" />
      <App />
    </MantineProvider>
  </React.StrictMode>
);
