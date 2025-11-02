# SheetGrid - AI-Powered Excel Editor

A modern, AI-powered spreadsheet editor built with Vite, React, and OpenRouter. Upload Excel files, edit them in a beautiful interface, and interact with your data using natural language AI commands.

## Features

### Spreadsheet Capabilities
- 📊 **Excel File Upload**: Upload and edit `.xlsx` or `.xls` files
- ✏️ **Interactive Editing**: Full-featured spreadsheet editing
- 📐 **Formatting**: Bold, italic, borders, colors, and more
- 🔀 **Merge Cells**: Combine multiple cells
- 📊 **Sorting & Filtering**: Organize data easily
- 📋 **Copy/Paste**: Full clipboard support
- 📏 **Resize**: Drag to resize columns and rows
- 🗂️ **Multiple Sheets**: Manage multiple worksheets
- 💬 **Comments**: Add cell comments
- 🔗 **Auto-sizing**: Automatic column/row sizing
- ↶ **Undo/Redo**: Full undo/redo support
- 🖱️ **Drag Fill**: Fill cells by dragging
- 🔍 **Context Menu**: Right-click for quick actions

### AI-Powered Features
- 🤖 **AI Assistant**: Chat with AI to analyze and edit spreadsheets
- 🔧 **Tool Calling**: AI can read/modify cells, ranges, and data
- 🎯 **Natural Language**: Ask questions in plain English
- 📈 **Data Analysis**: AI-powered insights and operations

## Prerequisites

- Node.js 18+ and pnpm
- An [OpenRouter](https://openrouter.ai) API key

### Getting an OpenRouter API Key

1. Sign up at [https://openrouter.ai](https://openrouter.ai)
2. Create an API key in your dashboard
3. Copy your API key

## Installation

1. Clone the repository:

```bash
git clone <repository-url>
cd sheetgrid
```

2. Install dependencies:

```bash
pnpm install
```

3. Configure your OpenRouter API key:

Create a `.env` file in the root directory:

```bash
echo "VITE_OPENROUTER_API_KEY=your_api_key_here" > .env
```

4. Start the development server:

```bash
pnpm dev
```

5. Open [http://localhost:5173](http://localhost:5173) in your browser

### Generate Sample Data

Optionally, you can generate a sample Excel file for testing:

```bash
pnpm exec node scripts/generate-sample-data.js
```

This will create `examples/sample-data.xlsx`.

## Usage

### Basic Editing

1. The spreadsheet starts with an empty sheet ready to use
2. Click "Upload" button to load an existing Excel file
3. Edit cells by clicking and typing directly
4. Right-click on cells for formatting options
5. Use column headers for sorting and filtering
6. Drag column/row borders to resize

### AI Assistant

The AI assistant can help you:

- Read cell values
- Update multiple cells
- Analyze data in ranges
- Perform bulk operations

**Example commands:**

- "What's in cell A1?"
- "Set cell B5 to 'Total'"
- "Show me the range from A1 to C10"
- "Set all values in column D to zero"
- "Calculate the sum of cells A1 to A10 and put it in A11"

### Keyboard Shortcuts

- `Enter`: Save cell and move down
- `Tab`: Save cell and move right
- `Ctrl+C / Cmd+C`: Copy selected cells
- `Ctrl+V / Cmd+V`: Paste
- `Ctrl+Z / Cmd+Z`: Undo
- `Ctrl+Y / Cmd+Y`: Redo
- `Ctrl+F / Cmd+F`: Find
- `Delete`: Clear selected cells
- `Escape`: Cancel editing

## Tech Stack

- **Vite**: Fast build tool and development server
- **React 19**: UI library
- **TypeScript**: Type safety
- **Tailwind CSS**: Styling
- **Univer**: Professional Excel-like spreadsheet component
- **Univer MCP**: MCP integration for AI tool calling
- **SheetJS (xlsx)**: Excel file parsing and generation
- **ECharts**: Data visualization and charting
- **file-saver**: File download capabilities
- **jszip**: ZIP file manipulation
- **fast-xml-parser**: Fast XML parsing
- **OpenRouter**: AI model routing with tool calling support

## Project Structure

```
sheetgrid/
├── src/
│   ├── main.tsx          # Entry point
│   ├── App.tsx           # Main app component
│   └── index.css         # Global styles
├── components/
│   ├── Spreadsheet.tsx   # Univer spreadsheet component
│   └── ChatPanel.tsx     # AI chat interface with tool calling
├── scripts/
│   └── generate-sample-data.js  # Sample Excel generator
├── index.html            # HTML template
├── vite.config.ts        # Vite configuration
└── package.json
```

## Development

Run the development server with hot reload:

```bash
pnpm dev
```

Build for production:

```bash
pnpm build
pnpm preview
```

## MCP Integration

SheetGrid supports MCP (Model Context Protocol) integration, allowing AI assistants like Cursor to directly manipulate spreadsheets through natural language.

See [MCP_SETUP.md](./MCP_SETUP.md) for detailed setup instructions.

## Troubleshooting

### OpenRouter Connection Issues

If you see errors about OpenRouter API:

1. Make sure your API key is set in `.env`: `VITE_OPENROUTER_API_KEY=your_key`
2. Verify your API key is valid at [https://openrouter.ai/keys](https://openrouter.ai/keys)
3. Check that you have sufficient credits in your OpenRouter account

### Model Selection

You can choose from various free-tier models with tool calling support in the chat panel dropdown:
- `meta-llama/llama-3.2-3b-instruct:free` - Meta's Llama 3.2
- `google/gemini-flash-1.5` - Google's Gemini Flash 1.5
- `mistralai/mistral-7b-instruct:free` - Mistral 7B
- `mistralai/mixtral-8x7b-instruct:free` - Mixtral 8x7B

## License

MIT License

## Contributing

Contributions welcome! Please feel free to submit a Pull Request.
