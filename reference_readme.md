# SPFx Modern Calendar Web Part

A modern, responsive calendar web part for SharePoint Online, built with the SharePoint Framework (SPFx). This web part displays events from a SharePoint list with advanced filtering, explicit timezone support (CST/EST), and a custom details panel.

## Key Features

- **Modern UI**: Built with React and Fluent UI for a seamless SharePoint experience.
- **Timezone Awareness**: Explicitly displays start and end times in both **Central (CST)** and **Eastern (EST)** time zones to avoid confusion.
- **Rich Details Panel**: Custom side panel showing extended event metadata (SME, Contact, Prework, etc.).
- **Category Styling**: Color-coded events based on categories.
- **Absolute Asset Loading**: Robust asset loading logic to bypass local host issues in production.

## Technical Stack

- **Framework**: SharePoint Framework (SPFx) v1.22
- **Library**: React v17
- **Language**: TypeScript & SCSS
- **Utilities**: 
  - `moment-timezone`: For precise time handling.
  - `dompurify`: For sanitizing HTML descriptions.
  - `jquery`: For DOM manipulation support.

## Backend Configuration (SharePoint List)

The web part specifically looks for the following columns in your SharePoint List. Ensure these Internal Names exist:

| Display Name Example       | Preferred Internal Name       | Type           | Description |
|----------------------------|-------------------------------|----------------|-------------|
| **Title**                  | `Title`                       | Single Line    | Event Title |
| **Start Date**             | `EventDate`                   | Date & Time    | Standard Calendar Start |
| **End Date**               | `EndDate`                     | Date & Time    | Standard Calendar End |
| **Description**            | `Description`                 | Multi-line     | Rich text description |
| **Category**               | `Category`                    | Choice         | Event Category |
| **Color**                  | `CategoryColor`               | Single Line    | Hex code (e.g., #FF0000) |
| **Target Audience**        | `TargetAudience`              | Choice/Text    | e.g. "All Employees" |
| **Location**               | `Location`                    | Single Line    | Room or URL |
| **Department**             | `Department`                  | Choice/Text    | Hosting department |
| **Contact**                | `Contact`                     | Single Line    | Name/Email of contact |
| **Subject Matter Expert**  | `SubjectMatterExpert`         | Single Line    | Name of SME |
| **Prework**                | `Prework`                     | Choice/Text    | Yes/No or Description |
| **Registration Deadline**  | `RegistrationDate`            | Date & Time    | Deadline for sign-up |
| **Start Time CST**         | `StartTimeZoneCST`            | Single Line    | Display text (e.g. "9:00 AM") |
| **End Time CST**           | `EndTimeZoneCST`              | Single Line    | Display text (e.g. "10:00 AM") |
| **Start Time EST**         | `StartTimeZoneEST`            | Single Line    | Display text (e.g. "10:00 AM") |
| **End Time EST**           | `EndTimeZoneEST`              | Single Line    | Display text (e.g. "11:00 AM") |

> **Note**: The web part attempts to auto-map these columns. You can map them manually in the Web Part Property Pane if names differ.

## Build and Deployment

This project uses a standard SPFx build pipeline with a custom Node.js helper to ensure correct asset extraction.

### Prerequisites
- Node.js v16 or v18 (Recommended)
- Gulp CLI (`npm install -g gulp-cli`)

### Commands

**1. Install Dependencies**
```bash
npm install
```

**2. Create Production Build**
Generates the optimized JavaScript bundles.
```bash
npm run build -- --production
```

**3. Package Solution**
Creates the `.sppkg` file and automatically fixes the manifest to ensure assets are extracted.
```bash
npm run package-complete
```
*Note: This command runs `npm run package-solution -- --production` and then executes `scripts/fix-package.js`.*

### Deployment
1. Navigate to `sharepoint/solution/modern-calendar.sppkg`.
2. Upload this file to your tenant's **App Catalog**.
3. SharePoint will ask to trust the solution. Click **Deploy**.
4. The assets (JavaScript files) will be automatically extracted to the `ClientSideAssets` library in the App Catalog site.

## Troubleshooting

- **Assets Not Loading (404)**: Ensure you ran the build with `--production`. Development builds point to `localhost`.
- **Timezone Issues**: Verify the `Start Time CST/EST` columns are populated in the list. The web part calculates dates but relies on these text fields for the specific display in the details panel.
