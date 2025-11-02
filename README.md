# SheetGrid - AI-Powered Excel Editor

A modern, AI-powered spreadsheet editor built with Next.js and Ollama. Upload Excel files, edit them in a beautiful interface, and interact with your data using natural language AI commands.

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
- [Ollama](https://ollama.ai/) installed and running locally
- The `llama3.2` model installed in Ollama

### Installing Ollama

1. Download and install Ollama from [https://ollama.ai](https://ollama.ai)
2. Start the Ollama service
3. Install the llama3.2 model:

```bash
ollama pull llama3.2
```

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

3. Start the development server:

```bash
pnpm dev
```

4. Open [http://localhost:3000](http://localhost:3000) in your browser

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

- **Next.js 16**: React framework with App Router
- **TypeScript**: Type safety
- **Tailwind CSS**: Styling
- **Univer**: Professional Excel-like spreadsheet component
- **Univer MCP**: MCP integration for AI tool calling
- **SheetJS (xlsx)**: Excel file parsing and generation
- **ECharts**: Data visualization and charting
- **file-saver**: File download capabilities
- **jszip**: ZIP file manipulation
- **fast-xml-parser**: Fast XML parsing
- **Ollama**: Local AI inference with function calling

## Project Structure

```
sheetgrid/
├── app/
│   ├── layout.tsx        # Root layout
│   ├── page.tsx          # Main page component
│   └── globals.css       # Global styles
├── components/
│   ├── Spreadsheet.tsx   # Univer spreadsheet component
│   └── ChatPanel.tsx     # AI chat interface with tool calling
├── scripts/
│   └── generate-sample-data.js  # Sample Excel generator
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
pnpm start
```

## MCP Integration

SheetGrid supports MCP (Model Context Protocol) integration, allowing AI assistants like Cursor to directly manipulate spreadsheets through natural language.

See [MCP_SETUP.md](./MCP_SETUP.md) for detailed setup instructions.

## Troubleshooting

### Ollama Connection Issues

If you see errors about Ollama connection:

1. Make sure Ollama is running: `ollama serve`
2. Verify the model is installed: `ollama list`
3. Test the API: `curl http://localhost:11434/api/tags`

### Model Performance

For better performance with large spreadsheets, consider using a more powerful model:

```bash
ollama pull llama3.3
```

Then update the model name in `components/ChatPanel.tsx` (line ~143).

## License

MIT License

## Contributing

Contributions welcome! Please feel free to submit a Pull Request.
