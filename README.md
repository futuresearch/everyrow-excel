# EveryRow Excel Add-in

Screen, rank, dedupe, and merge your Excel spreadsheets using natural language.

## Features

- **Rank** - Score and sort rows by qualitative criteria
- **Screen** - Filter rows that match specific conditions
- **Dedupe** - Remove semantic duplicates
- **Agent** - Run AI web research on each row
- **Merge** - Join tables using AI-powered matching

## Development Setup

### Prerequisites

- Node.js 18+
- pnpm
- Excel (desktop or web)

### Installation

```bash
# Install dependencies
pnpm install

# Generate HTTPS certificates for local development
pnpm certs
```

### Running Locally

```bash
# Start the development server with HTTPS
pnpm start
```

The add-in will be available at `https://localhost:3000`.

### Sideloading the Add-in

#### Excel on the Web

1. Open Excel Online (office.com/excel)
2. Open or create a workbook
3. Go to **Insert > Office Add-ins > Upload My Add-in**
4. Upload the `manifest.xml` file
5. Click on the **EveryRow** button in the Home tab

#### Excel Desktop (Windows)

1. Open Excel
2. Go to **Insert > My Add-ins > Shared Folder**
3. Browse to this project's `manifest.xml`
4. Click **Add**

#### Excel Desktop (Mac)

1. Open Excel
2. Go to **Insert > Add-ins > My Add-ins**
3. Click **Manage My Add-ins** (gear icon)
4. Click **Add a custom add-in > Add from file**
5. Select the `manifest.xml` file

### Building for Production

```bash
pnpm build
```

The built files will be in the `dist/` directory.

## Usage

1. Get an API key at [everyrow.io/api-key](https://everyrow.io/api-key)
2. Open the EveryRow task pane from the Home ribbon
3. Enter your API key
4. Select data in your spreadsheet (including headers)
5. Choose an operation and configure it
6. Click Run

Results will be written to a new worksheet.

## Project Structure

```
everyrow-excel/
├── src/
│   ├── taskpane/           # React UI components
│   │   ├── App.tsx         # Main app component
│   │   └── components/     # UI components
│   ├── api/                # EveryRow API client
│   ├── excel/              # Excel data handling
│   └── config/             # Settings storage
├── assets/                 # Icons
├── manifest.xml            # Office Add-in manifest
├── taskpane.html           # Entry HTML
└── vite.config.ts          # Build configuration
```

## API

This add-in communicates with the EveryRow API at `https://engine.futuresearch.ai`.

See [everyrow-sdk](https://github.com/futuresearch/everyrow-sdk) for the Python SDK.

## License

MIT
