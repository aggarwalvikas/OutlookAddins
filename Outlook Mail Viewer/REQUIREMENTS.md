# Outlook Mail Viewer

## 1. Overview

**Outlook Mail Viewer** is an Outlook task pane add-in built with React and TypeScript. It reads the subject and sender of the email currently opened in Outlook and displays those details in the task pane.

## 2. Functional Requirements

| ID    | Requirement                                                                | Status   |
| ----- | -------------------------------------------------------------------------- | -------- |
| FR-01 | The add-in must be named `Outlook Mail Viewer`.                            | Complete |
| FR-02 | The Outlook command label must be `Show Details`.                          | Complete |
| FR-03 | The add-in must activate for a message opened in read mode.                | Complete |
| FR-04 | The task pane must display the email subject.                              | Complete |
| FR-05 | The task pane must display the sender display name.                        | Complete |
| FR-06 | The task pane must display the sender email address.                       | Complete |
| FR-07 | An empty subject must display as `(No subject)`.                           | Complete |
| FR-08 | Missing sender information must display a readable fallback message.       | Complete |
| FR-09 | Unsupported contexts must display an instruction to open an email message. | Complete |
| FR-10 | The add-in must not modify the email or its contents.                      | Complete |

## 3. Technical Requirements

- Frontend: React 17 and TypeScript.
- Office integration: Office.js.
- Host: Outlook desktop, including Outlook 2016.
- Manifest: Classic Outlook XML add-in manifest.
- Permission: `ReadItem`.
- Activation rule: `ItemIs` with `ItemType="Message"` and `FormType="Read"`.
- Development server: HTTPS on `https://localhost:3000`.
- Build tool: Webpack.
- Browser compatibility: Babel targets Internet Explorer 11-era webviews for legacy Outlook compatibility.
- UI styling: Native React inline styles to avoid unsupported modern UI runtime dependencies in Outlook 2016.

## 4. User Flow

1. Start the HTTPS development server.
2. Sideload `manifest.xml` into Outlook.
3. Open an email message in Outlook.
4. Select **Show Details** from the Outlook ribbon.
5. The task pane displays the subject and sender details.

## 5. Project Structure

```text
Outlook Mail Viewer/
├── manifest.xml                         Outlook activation and command configuration
├── package.json                         Scripts and dependencies
├── tsconfig.json                        TypeScript configuration
├── webpack.config.js                    Development and production bundling
├── babel.config.json                    Legacy browser transpilation
├── assets/                              Add-in icons
└── src/
    ├── commands/commands.ts             Office command entry point
    └── taskpane/
        ├── index.tsx                    Office-ready React entry point
        ├── taskpane.html                Task pane HTML host
        └── components/App.tsx           Subject and sender display
```

## 6. Development Commands

Run commands from the `Outlook Mail Viewer` directory:

```powershell
npm install
npm run build
npm run validate
npm run lint
npm run dev-server
```

The production output is written to `dist/`.

## 7. Validation Checklist

- `npx tsc --noEmit` completes without TypeScript errors.
- `npm run build` completes successfully.
- `npm run validate` reports that the manifest is valid.
- `npm run lint` completes successfully.
- Outlook 2016 can load the manifest.
- **Show Details** appears for an opened email.
- Subject, sender name, and sender email address are displayed correctly.
- Blank subject and missing sender data are handled without a runtime error.

## 8. Known Constraints

- The add-in reads the message currently opened in Outlook. Selecting a message row without opening it is not the supported flow for this Outlook 2016 implementation.
- The development server must remain running while Outlook loads the local task pane.
- The local HTTPS certificate must be trusted on the development computer.
- Runtime behavior should be tested in the target Outlook 2016 installation because manifest validation and TypeScript compilation do not replace host testing.

## 9. Skills and Capabilities Used

| Skill                            | Application                                                                              |
| -------------------------------- | ---------------------------------------------------------------------------------------- |
| React and TypeScript             | Build the task pane UI and typed message details model.                                  |
| Office.js Outlook APIs           | Access `Office.context.mailbox.item.subject` and `item.from`.                            |
| Outlook add-in manifest design   | Configure read-mode activation, ribbon command, permissions, icons, and task pane URLs.  |
| Legacy browser compatibility     | Target Outlook 2016 webviews with ES2015 TypeScript output and Babel IE11 transpilation. |
| Webpack                          | Bundle the task pane, command entry point, HTML files, assets, and manifest.             |
| TypeScript diagnostics           | Validate source types and resolve Office.js and browser global declarations.             |
| ESLint and Office add-in linting | Check project conventions and source quality.                                            |
| Git workflow                     | Commit reviewed changes and push them to the configured GitHub remote.                   |

## 10. Out of Scope

- Reading or displaying the message body.
- Reading attachments.
- Editing, sending, moving, or deleting messages.
- Reading messages from a selected list row without opening them.
- Production hosting or deployment.
