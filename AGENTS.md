# AGENTS.md

This document provides context for AI agents working on this repository.

## Project Overview

This is a Google Apps Script project that generates contact text from Google Calendar events based on rules defined in a Google Sheet. It is written in TypeScript and uses `clasp` for deployment. It also functions as a web application.

## Tech Stack

- **Language:** TypeScript
- **Environment:** Google Apps Script (V8 runtime)
- **Package Manager:** pnpm
- **Deployment:** clasp
- **Linting:** ESLint
- **Formatting:** Prettier

## Project Structure

- `src/`: Contains the TypeScript source code.
  - `main.ts`: The main script file, containing both spreadsheet-bound functions and web app logic.
  - `appsscript.json`: The Google Apps Script manifest file.
  - `index.html`: The HTML file for the web application's user interface.
- `dist/`: Contains the compiled JavaScript code and HTML files that get deployed to Google Apps Script.
- `package.json`: Defines project scripts and dependencies.
- `tsconfig.json`: TypeScript compiler configuration.
- `.clasp.json`: `clasp` configuration file (this file is in `.gitignore`).
- `README.md`: Project documentation for humans.
- `LICENSE`: The ISC license file.

## Development Workflow

- **Install dependencies:** `pnpm install`
- **Build:** `pnpm run build` (compiles TypeScript, copies `appsscript.json` and `index.html` to `dist/`)
- **Deploy:** `pnpm run deploy` (runs the build script and then `clasp push`)
- **Lint:** `pnpm run lint`
- **Format:** `pnpm run format`

## Key Configurations

### `tsconfig.json`

- The `"lib"` option is set to `["ES2019", "ScriptHost"]` to avoid type conflicts between the standard DOM library and Google Apps Script's type definitions.

### `.clasp.json` (local file)

- **`scriptId`**: This must be set to the ID of your Google Apps Script project.
- **`rootDir`**: This is set to `"dist"` to ensure that `clasp` deploys the compiled JavaScript files and HTML files, not the source TypeScript files.

### Script Properties

- **`CALENDAR_ID`**: This script property must be set in the Google Apps Script project settings. It holds the ID of the Google Calendar to read events from. If not set, it defaults to the user's default calendar.

### Spreadsheet Setup

The script relies on a Google Sheet with a sheet named "ルール設定" (Rule Settings). This sheet must have the following columns:

- **Column A:** Keyword (for searching calendar events)
- **Column B:** Output Word (for use in the generated text; falls back to Keyword if empty)
- **Column C:** Action (for grouping the output)
- **Column D:** Required (a checkbox; if checked, the script will warn if no event with the keyword is found)

## Web Application Details

- **Entry Point:** `doGet(e)` function in `main.ts` serves `index.html`.
- **UI:** `index.html` provides a button to trigger text generation, an editable `textarea` for results, a warning area, and a debug table for calendar events.
- **Client-Server Communication:** Uses `google.script.run` to call `generateTextForWebApp()`.
- **Data Format:** `generateTextForWebApp()` returns a JSON object `{ mainOutput: string, debugEvents: {time: string, title: string}[], missingKeywordsWarning: string }`.
- **Debug Table:** Displays events with date grouping and `rowspan` for improved readability.
- **Copy Functionality:** A button allows copying the generated text to the clipboard.
- **Responsiveness:** The web app is designed to be mobile-friendly, with font size adjustments and horizontal scrolling for the table on smaller screens.