import { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { loadChatMessages, saveAllChatMessages, updateChat } from '../src/utils/indexeddb';

interface ChatPanelProps {
  chatId: string;
  chatTitle: string;
  onCreateNewChat: () => void;
  onChatTitleChange?: (title: string) => void;
  onSelectChat?: (chatId: string) => void;
  onDeleteChat?: (chatId: string) => void;
  allChats?: Array<{ id: string; title: string; updatedAt: string }>;
}

interface Message {
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
}

interface ToolCall {
  name: string;
  arguments: Record<string, unknown>;
}

export default function ChatPanel({ chatId, chatTitle, onCreateNewChat, onChatTitleChange, onSelectChat, onDeleteChat, allChats = [] }: ChatPanelProps) {
  // System message should not be persisted
  const systemMessage: Message = {
    role: 'system',
    content: `You are a helpful assistant that can edit Excel spreadsheets using tools. 

CRITICAL: When generating random numbers or values, you MUST provide actual numeric values in your tool calls:
- DO NOT use JavaScript code like "Math.floor(Math.random() * 1001)"
- DO NOT use Python code like "[Math.round(Math.random() * 1000) for i in range(30)]"
- DO NOT use any code syntax - only provide the actual numbers
- Example: If asked for 30 random numbers 0-1000, generate 30 actual numbers like [342, 891, 123, ...] not code
- Always calculate the values yourself and provide the complete array of actual numbers

IMPORTANT RULES:
1. Be concise and direct in your responses - users only want to see results, not your analysis or thinking process
2. Before performing operations, silently check the sheet structure using get_sheets and get_range_data
3. Do not explain what you're doing step-by-step - just execute the operations
4. After completing operations, provide only a brief confirmation of what was done
5. When the user asks you to create a column with dates or fill a range, you MUST:
   - Use set_column_header to create the header (if needed)
   - Use set_range_data to fill in ALL the requested values completely
   - NEVER stop after just a few rows - you must fill the ENTIRE requested range
   - If the user asks for dates from Nov 1 to Nov 30, you must create ALL 30 dates, not just 5
   - Make multiple tool calls if needed to complete the full range - never leave partial data
6. Always complete the ENTIRE task before responding - partial completion is not acceptable

CALCULATIONS AND NUMERIC OPERATIONS:
- When asked to calculate totals, sums, averages, or perform arithmetic operations:
  - Use get_range or get_range_data to retrieve the relevant numeric data
  - Parse formatted numbers from spreadsheet data (remove commas, spaces, handle parentheses as negatives)
  - Example: " 146,493 " → 146493, " (13,306)" → -13306, " -   " → 0
  - Perform the calculation yourself (sum, average, etc.)
  - Return the result directly with a clear answer
  - DO NOT just retrieve a single cell - you must process all the data and calculate
- When you receive range data, parse ALL numeric values and calculate based on the user's request
- Format your final answer clearly: "The total revenue is $X,XXX" or "The sum is XXX"
- IMPORTANT: When you receive tool results with data in markdown format like:
  "- Row 2
    - H2: v= 146,493 ;
  - Row 3
    - H3: v= (13,306);"
  You MUST:
  1. Extract ALL numeric values from ALL rows (skip header row with "Revenue", "Gross Profit", etc.)
  2. Parse each value: remove spaces, remove commas, treat parentheses as negative signs
  3. Convert each to a number: " 146,493 " becomes 146493, " (13,306)" becomes -13306, " -   " becomes 0
  4. Sum ALL the numeric values (excluding the header row)
  5. Return the final calculation: "The total revenue of all rows is $X,XXX,XXX"
- DO NOT respond with "Operations completed successfully" when asked for a calculation - you MUST provide the actual calculated result

Example of good response: "I've created a Date column with all dates from November 1 to November 30, 2025."
Example of bad response: "I set the first 5 dates, you can use auto_fill to extend the rest..."
Example of calculation response: "The total revenue of all rows is $542,893."`,
    timestamp: new Date(),
  };

  const [messages, setMessages] = useState<Message[]>([systemMessage]);
  const [isLoadingHistory, setIsLoadingHistory] = useState(true);
  const [input, setInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [selectedModel, setSelectedModel] = useState('anthropic/claude-3-haiku');
  const [isModelDropdownOpen, setIsModelDropdownOpen] = useState(false);
  const availableModels = [
    'anthropic/claude-3-haiku',
  ];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [mcpTools, setMcpTools] = useState<any[]>([]);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const modelDropdownRef = useRef<HTMLDivElement>(null);
  const abortControllerRef = useRef<AbortController | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Function to clean context from message content
  const cleanMessageContent = (content: string): string => {
    // Remove [Current Sheet Context] block if it exists
    const contextPattern = /\n\n\[Current Sheet Context\][\s\S]*$/;
    return content.replace(contextPattern, '').trim();
  };

  // Load chat history from IndexedDB when chatId changes
  useEffect(() => {
    const loadHistory = async () => {
      setIsLoadingHistory(true);
      try {
        const savedMessages = await loadChatMessages(chatId);
        if (savedMessages && savedMessages.length > 0) {
          // Filter out system messages from saved history (we'll add our own)
          // Also clean any context that might have been saved in old messages
          const userMessages = savedMessages
            .filter((msg: Message) => msg.role !== 'system')
            .map((msg: Message) => ({
              ...msg,
              content: cleanMessageContent(msg.content),
            }));
          setMessages([systemMessage, ...userMessages]);
          console.log(`Loaded ${userMessages.length} messages from IndexedDB for chat ${chatId}`);
        } else {
          // No saved messages, just set system message
          setMessages([systemMessage]);
        }
      } catch (error) {
        console.error('Error loading chat history:', error);
        setMessages([systemMessage]);
      } finally {
        setIsLoadingHistory(false);
      }
    };

    loadHistory();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [chatId]); // Reload when chatId changes

  // Save messages to IndexedDB whenever they change (debounced)
  useEffect(() => {
    if (isLoadingHistory) return; // Don't save while loading

    const saveTimeout = setTimeout(async () => {
      try {
        // Filter out system message before saving
        const messagesToSave = messages.filter(msg => msg.role !== 'system');
        if (messagesToSave.length > 0) {
          await saveAllChatMessages(chatId, messagesToSave);
        }
      } catch (error) {
        console.error('Error saving chat history:', error);
      }
    }, 1000); // Debounce saves by 1 second

    return () => clearTimeout(saveTimeout);
  }, [messages, isLoadingHistory, chatId]);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  // Function to get current sheet context
  const getSheetContext = (): string => {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const univerAPI = (window as any).univerAPI;
      if (!univerAPI) return '';

      const workbook = univerAPI.getActiveWorkbook();
      if (!workbook) return '';

      const sheet = workbook.getActiveSheet();
      if (!sheet) return '';

      // Get sheet name - try multiple methods to find the correct name
      let sheetName = 'Sheet1'; // Default to Sheet1 as Univer's default
      try {
        // Method 1: Get active sheet ID and find it in sheets list
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sheetId = (sheet as any).getSheetId?.();
        if (sheetId) {
          const sheets = workbook.getSheets();
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const foundSheet = sheets.find((s: any) => {
            try {
              return s.getSheetId?.() === sheetId;
            } catch {
              return false;
            }
          });
          if (foundSheet) {
            // Try multiple ways to get the name
            if (typeof foundSheet.getName === 'function') {
              sheetName = foundSheet.getName();
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
            } else if ((foundSheet as any).name) {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              sheetName = (foundSheet as any).name;
            }
          }
        }

        // Method 2: Try getName() directly on the sheet object
        if (sheetName === 'Sheet1' && typeof sheet.getName === 'function') {
          try {
            const directName = sheet.getName();
            if (directName) sheetName = directName;
          } catch {
            // Ignore
          }
        }
      } catch (e) {
        console.warn('Could not get sheet name:', e);
      }

      // Ensure we have a valid name (fallback to Sheet1 if empty)
      if (!sheetName || sheetName === 'Active Sheet') {
        sheetName = 'Sheet1';
      }

      // Get a sample of data (first 10 rows, first 10 columns) to understand structure
      let context = `Current Sheet: "${sheetName}"\n`;

      try {
        // Get headers (row 0) - check first 20 columns
        const headers = [];
        for (let col = 0; col < 20; col++) {
          try {
            const cellRange = sheet.getRange(0, col);
            const value = cellRange.getValue();
            if (value !== null && value !== undefined && String(value).trim() !== '') {
              headers.push(String(value));
            } else if (headers.length > 0) {
              // Stop if we hit an empty cell after finding headers
              break;
            }
          } catch {
            break;
          }
        }

        if (headers.length > 0) {
          context += `Column Headers: ${headers.join(', ')}\n`;
        }

        // Try to get dimensions by reading a larger range
        // Read first 100 rows and 20 columns to understand structure
        let rowCount = 0;
        let colCount = headers.length || 10;

        try {
          const sampleRange = sheet.getRange(0, 0, 100, Math.max(colCount, 20));
          const sampleData = sampleRange.getValues();
          rowCount = sampleData.length;
          colCount = sampleData[0]?.length || colCount;
        } catch {
          // Fallback: try smaller range
          try {
            const smallRange = sheet.getRange(0, 0, 10, 10);
            const smallData = smallRange.getValues();
            rowCount = smallData.length;
            colCount = smallData[0]?.length || colCount;
          } catch {
            // Use defaults
            rowCount = 100;
          }
        }

        context += `Sheet Dimensions: ~${rowCount} rows × ~${colCount} columns\n`;

        // Get a sample of the first few data rows (skip header row 0)
        const sampleRows = Math.min(5, Math.max(1, rowCount - 1));
        context += `\nSample Data (first ${sampleRows} data rows):\n`;

        for (let row = 1; row <= sampleRows && row < rowCount; row++) {
          const rowData = [];
          for (let col = 0; col < Math.min(headers.length || colCount, colCount); col++) {
            try {
              const cellRange = sheet.getRange(row, col);
              const value = cellRange.getValue();
              rowData.push(value !== null && value !== undefined ? String(value).substring(0, 30) : '');
            } catch {
              rowData.push('');
            }
          }
          if (rowData.some(cell => cell.trim() !== '')) {
            context += `Row ${row}: ${rowData.join(' | ')}\n`;
          }
        }
      } catch (e) {
        console.warn('Error getting sheet context:', e);
      }

      return context;
    } catch (error) {
      console.warn('Could not get sheet context:', error);
      return '';
    }
  };


  // Auto-resize textarea
  useEffect(() => {
    const textarea = textareaRef.current;
    if (textarea) {
      textarea.style.height = 'auto';
      textarea.style.height = `${Math.min(textarea.scrollHeight, 120)}px`;
    }
  }, [input]);

  // Close dropdown when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (modelDropdownRef.current && !modelDropdownRef.current.contains(event.target as Node)) {
        setIsModelDropdownOpen(false);
      }
    };

    if (isModelDropdownOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }

    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isModelDropdownOpen]);

  // Fetch MCP tools from Univer MCP server
  useEffect(() => {
    const fetchMcpTools = async () => {
      try {
        const apiKey = import.meta.env.VITE_UNIVER_MCP_API_KEY || '';
        if (!apiKey) {
          console.warn('No MCP API key configured, using local tools only');
          return;
        }

        // Wait for MCP connection to be established
        await new Promise(resolve => setTimeout(resolve, 3000));

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sessionId = (window as any).univerSessionId || 'default';

        console.log('Fetching MCP tools with sessionId:', sessionId);

        // Try HTTP POST to fetch tools (official MCP endpoint)
        try {
          const response = await fetch(`https://mcp.univer.ai/mcp/?univer_session_id=${sessionId}`, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': `Bearer ${apiKey}`,
              'Accept': 'application/json, text/event-stream',
            },
            body: JSON.stringify({
              jsonrpc: '2.0',
              id: 1,
              method: 'tools/list',
              params: {},
            }),
          });

          console.log('MCP tools HTTP response status:', response.status);
          console.log('MCP tools response content-type:', response.headers.get('content-type'));

          if (response.ok) {
            // Check if response is SSE (text/event-stream) or JSON
            const contentType = response.headers.get('content-type') || '';
            let data;

            if (contentType.includes('text/event-stream')) {
              // Parse SSE stream
              const text = await response.text();
              console.log('MCP tools SSE response:', text.substring(0, 200));

              // Parse SSE format: event: message\ndata: {...}\n\n
              const lines = text.split('\n');
              let jsonData = '';
              for (let i = 0; i < lines.length; i++) {
                if (lines[i].startsWith('data: ')) {
                  jsonData = lines[i].substring(6); // Remove 'data: ' prefix
                  break;
                }
              }

              if (jsonData) {
                data = JSON.parse(jsonData);
              } else {
                // Try to find JSON in the response
                const jsonMatch = text.match(/\{[\s\S]*\}/);
                if (jsonMatch) {
                  data = JSON.parse(jsonMatch[0]);
                } else {
                  throw new Error('No JSON data found in SSE response');
                }
              }
            } else {
              // Regular JSON response
              data = await response.json();
            }

            console.log('MCP tools parsed response:', data);
            if (data.result && data.result.tools) {
              // Convert MCP tool format to OpenAI format
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const convertedTools = data.result.tools.map((tool: any) => ({
                type: 'function',
                function: {
                  name: tool.name,
                  description: tool.description || '',
                  parameters: tool.inputSchema || { type: 'object', properties: {}, required: [] },
                },
              }));
              console.log('Loaded MCP tools from server:', convertedTools.length);
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              console.log('MCP tool names:', convertedTools.map((t: any) => t.function.name));
              setMcpTools(convertedTools);
              return;
            }
          } else {
            const errorText = await response.text();
            console.warn('MCP tools HTTP failed:', response.status, errorText);
            console.warn('Make sure: 1) API key is valid, 2) Univer instance is running, 3) sessionId matches');
          }
        } catch (httpError) {
          console.error('MCP tools HTTP request failed:', httpError);
          console.warn('Falling back to local tool implementations');
        }
      } catch (error) {
        console.error('Failed to fetch MCP tools, exception:', error);
      }
    };

    fetchMcpTools();
  }, []);

  // Return only MCP tools from server
  const getTools = () => {
    if (mcpTools.length === 0) {
      console.warn('No MCP tools available yet. Make sure the MCP API key is configured.');
      return [];
    }
    console.log(`getTools: returning ${mcpTools.length} tools from MCP server`);
    return mcpTools;
  };

  const executeTool = async (toolCall: ToolCall): Promise<string> => {
    const { name, arguments: args } = toolCall;

    // Check if this is an MCP tool (from server)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const isMcpTool = mcpTools.some((tool: any) => tool.function.name === name);

    if (!isMcpTool) {
      return JSON.stringify({ error: `Tool "${name}" is not available. Only MCP server tools are supported.` });
    }

    console.log(`Tool "${name}": routing to MCP server`, 'args:', JSON.stringify(args));

    const apiKey = import.meta.env.VITE_UNIVER_MCP_API_KEY || '';
    if (!apiKey) {
      return JSON.stringify({ error: 'MCP API key not configured' });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const sessionId = (window as any).univerSessionId || 'default';

    try {
      const response = await fetch(`https://mcp.univer.ai/mcp/?univer_session_id=${sessionId}`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${apiKey}`,
          'Accept': 'application/json, text/event-stream',
        },
        body: JSON.stringify({
          jsonrpc: '2.0',
          id: Date.now(),
          method: 'tools/call',
          params: {
            name: name,
            arguments: args,
          },
        }),
      });

      if (!response.ok) {
        try {
          const errorText = await response.text();
          // Try to parse as SSE first
          let errorData;
          if (errorText.includes('data: ')) {
            const lines = errorText.split('\n');
            for (const line of lines) {
              if (line.startsWith('data: ')) {
                errorData = JSON.parse(line.substring(6));
                break;
              }
            }
          } else {
            errorData = JSON.parse(errorText);
          }
          throw new Error(`MCP tool execution failed: ${errorData.error?.message || 'Unknown error'}`);
        } catch {
          throw new Error(`MCP tool execution failed: ${response.status} ${response.statusText}`);
        }
      }

      // Check if response is SSE (text/event-stream) or JSON
      const contentType = response.headers.get('content-type') || '';
      let data;

      if (contentType.includes('text/event-stream')) {
        // Parse SSE stream
        const text = await response.text();
        const lines = text.split('\n');
        let jsonData = '';
        for (let i = 0; i < lines.length; i++) {
          if (lines[i].startsWith('data: ')) {
            jsonData = lines[i].substring(6); // Remove 'data: ' prefix
            break;
          }
        }

        if (jsonData) {
          data = JSON.parse(jsonData);
        } else {
          // Try to find JSON in the response
          const jsonMatch = text.match(/\{[\s\S]*\}/);
          if (jsonMatch) {
            data = JSON.parse(jsonMatch[0]);
          } else {
            throw new Error('No JSON data found in SSE response');
          }
        }
      } else {
        // Regular JSON response
        data = await response.json();
      }

      // MCP tools/call returns { result: { content: [...] } }
      if (data.result && data.result.content) {
        const content = data.result.content;
        // Content is an array, extract text
        if (Array.isArray(content) && content.length > 0) {
          return content[0].text || JSON.stringify(content[0]);
        }
        return JSON.stringify(content);
      }
      return JSON.stringify(data.result || { success: true });
    } catch (error) {
      return JSON.stringify({ error: `MCP tool execution failed: ${error}` });
    }
  };

  const stopMessage = () => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
      abortControllerRef.current = null;
    }
    setIsLoading(false);
  };

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    // Check if it's an Excel file
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
      alert('Please upload an Excel file (.xlsx or .xls)');
      return;
    }

    setIsUploading(true);
    const uploadStartTime = Date.now();

    try {
      // Read file as array buffer
      const arrayBuffer = await file.arrayBuffer();

      // Read workbook - make sure we're not limiting rows
      // Note: XLSX.read() doesn't have a default row limit, but we'll be explicit
      const workbook = XLSX.read(arrayBuffer, {
        type: 'array',
        cellText: false,
        cellDates: true,
        sheetRows: 0, // 0 means no limit - read all rows
      });

      console.log(`File "${file.name}" loaded: ${workbook.SheetNames.length} sheet(s)`, {
        size: (arrayBuffer.byteLength / 1024 / 1024).toFixed(2) + ' MB'
      });

      // Get Univer API
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const univerAPI = (window as any).univerAPI;
      if (!univerAPI) {
        alert('Spreadsheet not ready. Please wait a moment and try again.');
        return;
      }

      // Get active workbook
      const activeWorkbook = univerAPI.getActiveWorkbook();
      if (!activeWorkbook) {
        alert('Could not access spreadsheet. Please refresh the page.');
        return;
      }

      // Clear existing sheets (keep at least one)
      const existingSheets = activeWorkbook.getSheets();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      existingSheets.forEach((sheet: any, index: number) => {
        if (index > 0) {
          try {
            activeWorkbook.deleteSheet(sheet.getSheetId());
          } catch {
            console.warn('Could not delete sheet');
          }
        }
      });

      // Process each sheet in the Excel file
      workbook.SheetNames.forEach((sheetName, sheetIndex) => {
        const worksheet = workbook.Sheets[sheetName];

        // Find the actual maximum row and column by scanning ALL cell keys
        // This ensures we catch all data, even if !ref is incomplete
        let maxRow = -1;
        let maxCol = -1;

        // Scan all keys in the worksheet object to find the actual maximum row/col
        for (const key in worksheet) {
          // Skip metadata keys that start with '!'
          if (key.startsWith('!')) continue;

          // Decode the cell address (e.g., 'A1' -> row 0, col 0)
          const cellAddr = XLSX.utils.decode_cell(key);
          if (cellAddr.r > maxRow) maxRow = cellAddr.r;
          if (cellAddr.c > maxCol) maxCol = cellAddr.c;
        }

        // Also check !ref as a fallback/validation
        const sheetRange = worksheet['!ref'];
        if (sheetRange) {
          const range = XLSX.utils.decode_range(sheetRange);
          if (range.e.r > maxRow) maxRow = range.e.r;
          if (range.e.c > maxCol) maxCol = range.e.c;
          console.log(`Sheet "${sheetName}": !ref says ${sheetRange}, actual scan found rows up to ${maxRow + 1}, cols up to ${maxCol + 1}`);
        } else {
          console.log(`Sheet "${sheetName}": No !ref, scanning found rows up to ${maxRow + 1}, cols up to ${maxCol + 1}`);
        }

        if (maxRow < 0 || maxCol < 0) {
          console.log(`Sheet "${sheetName}" appears empty, skipping`);
          return;
        }

        console.log(`Reading ${maxRow + 1} rows and ${maxCol + 1} columns from sheet "${sheetName}"`);

        // Read the sheet cell-by-cell to ensure we get ALL cells
        // This ensures we capture all data even if there are gaps
        const jsonData: string[][] = [];

        // Initialize all rows up to maxRow
        for (let row = 0; row <= maxRow; row++) {
          jsonData[row] = [];
          // Read all columns up to maxCol for this row
          for (let col = 0; col <= maxCol; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = worksheet[cellAddress];

            if (cell) {
              // Cell has data - use its value
              if (cell.v !== undefined && cell.v !== null) {
                jsonData[row][col] = String(cell.v);
              } else {
                jsonData[row][col] = '';
              }
            } else {
              // Cell is empty
              jsonData[row][col] = '';
            }
          }
        }

        console.log(`Loaded ${jsonData.length} rows from sheet "${sheetName}"`);

        // Verify we have data beyond row 1000 if the file should have it
        if (jsonData.length > 1000) {
          console.log(`✅ Sheet has ${jsonData.length} rows (beyond 1000) - checking for data in rows 1000+`);
          // Check if there's actual data beyond row 1000
          let hasDataBeyond1000 = false;
          for (let checkRow = 1000; checkRow < Math.min(1010, jsonData.length); checkRow++) {
            if (jsonData[checkRow] && jsonData[checkRow].some(cell => cell && cell.trim() !== '')) {
              hasDataBeyond1000 = true;
              console.log(`✅ Found data in row ${checkRow + 1}`);
              break;
            }
          }
          if (!hasDataBeyond1000) {
            console.warn(`⚠️ No data found in rows 1000-1010 - file might only have data up to row 1000`);
          }
        }

        // jsonData is already properly sized - all rows are maxCol+1 length
        // No need to pad since we read cell-by-cell to the exact range bounds
        const paddedData = jsonData;
        const maxCols = maxCol + 1; // Number of columns (maxCol is 0-indexed)

        if (sheetIndex === 0) {
          // Use the first existing sheet
          const firstSheet = activeWorkbook.getActiveSheet();
          if (firstSheet && paddedData.length > 0 && maxCols > 0) {
            try {
              // Convert to 2D array of strings for setValues
              const values2D: string[][] = paddedData.map(row =>
                row.map(cell => cell !== undefined && cell !== null ? String(cell) : '')
              );

              console.log(`Setting ${values2D.length} rows into Univer sheet (max row index: ${values2D.length - 1})`);

              // Process all files in small batches to avoid Univer issues
              const BATCH_SIZE = 100; // Use 100 rows per batch for reliability
              const totalRows = values2D.length;

              console.log(`Processing ${totalRows} rows in batches of ${BATCH_SIZE}`);

              for (let startRow = 0; startRow < totalRows; startRow += BATCH_SIZE) {
                const endRow = Math.min(startRow + BATCH_SIZE - 1, totalRows - 1);
                const batch = values2D.slice(startRow, endRow + 1);

                // Ensure batch is properly formatted
                const formattedBatch = batch.map(row => {
                  const formattedRow: string[] = [];
                  for (let i = 0; i < maxCols; i++) {
                    formattedRow[i] = row && row[i] !== undefined && row[i] !== null ? String(row[i]) : '';
                  }
                  return formattedRow;
                });

                try {
                  const range = firstSheet.getRange(startRow, 0, endRow, maxCols - 1);
                  range.setValues(formattedBatch);
                  console.log(`✅ Set rows ${startRow} to ${endRow}`);
                } catch (batchError) {
                  console.warn(`⚠️ Failed to set batch ${startRow}-${endRow}, skipping:`, batchError);
                  // Skip this batch and continue
                }
              }

              console.log(`✅ Finished processing all ${totalRows} rows`);
            } catch (e) {
              console.warn('Error processing file, showing partial load:', e);
              // Continue - we may have loaded some data
            }

            // Rename the sheet
            try {
              firstSheet.setName(sheetName);
            } catch (e) {
              console.warn('Could not rename sheet:', e);
            }
          }
        } else {
          // Create new sheets for additional sheets in the Excel file
          try {
            const newSheet = activeWorkbook.insertSheet(sheetName);

            if (paddedData.length > 0 && maxCols > 0) {
              try {
                // Convert to 2D array of strings
                const values2D: string[][] = paddedData.map(row =>
                  row.map(cell => cell !== undefined && cell !== null ? String(cell) : '')
                );

                // Process in batches
                const BATCH_SIZE = 100;
                const totalRows = values2D.length;

                console.log(`Processing ${totalRows} rows in sheet "${sheetName}" in batches of ${BATCH_SIZE}`);
                for (let startRow = 0; startRow < totalRows; startRow += BATCH_SIZE) {
                  const endRow = Math.min(startRow + BATCH_SIZE - 1, totalRows - 1);
                  const batch = values2D.slice(startRow, endRow + 1);

                  // Ensure batch is properly formatted
                  const formattedBatch = batch.map(row => {
                    const formattedRow: string[] = [];
                    for (let i = 0; i < maxCols; i++) {
                      formattedRow[i] = row && row[i] !== undefined && row[i] !== null ? String(row[i]) : '';
                    }
                    return formattedRow;
                  });

                  try {
                    const range = newSheet.getRange(startRow, 0, endRow, maxCols - 1);
                    range.setValues(formattedBatch);
                    console.log(`Set rows ${startRow} to ${endRow} in sheet "${sheetName}"`);
                  } catch {
                    console.warn(`Failed to set batch ${startRow}-${endRow} in "${sheetName}", skipping`);
                  }
                }
              } catch (e) {
                console.error(`Error processing sheet "${sheetName}":`, e);
              }
            }
          } catch (e) {
            console.warn(`Could not create sheet "${sheetName}":`, e);
          }
        }
      });

      // Reset file input
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }

      // Trigger save after upload
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      if (typeof window !== 'undefined' && (window as any).saveWorkbookData) {
        setTimeout(async () => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          await (window as any).saveWorkbookData();
        }, 1000);
      }

      // Show success message with performance info
      const uploadTime = ((Date.now() - uploadStartTime) / 1000).toFixed(2);
      const fileSizeMB = (file.size / (1024 * 1024)).toFixed(2);
      const message: Message = {
        role: 'assistant',
        content: `Successfully loaded "${file.name}" (${fileSizeMB} MB) with ${workbook.SheetNames.length} sheet(s) in ${uploadTime}s.`,
        timestamp: new Date(),
      };
      setMessages((prev) => [...prev, message]);
    } catch (error) {
      console.error('Error loading file:', error);
      const errorMessage: Message = {
        role: 'assistant',
        content: `Failed to load file: ${error instanceof Error ? error.message : 'Unknown error'}`,
        timestamp: new Date(),
      };
      setMessages((prev) => [...prev, errorMessage]);
    } finally {
      setIsUploading(false);
    }
  };

  const handleUploadClick = () => {
    fileInputRef.current?.click();
  };

  const sendMessage = async () => {
    if (!input.trim() || isLoading) return;

    // Create new AbortController for this request
    const abortController = new AbortController();
    abortControllerRef.current = abortController;

    // Get current sheet context (for AI, not for display)
    const context = getSheetContext();

    // Store only user's input in the message (clean, no context)
    const userMessage: Message = {
      role: 'user',
      content: input.trim(),
      timestamp: new Date(),
    };

    // Update chat title if this is the first user message
    const userMessages = messages.filter(msg => msg.role === 'user');
    if (userMessages.length === 0) {
      // This is the first user message, update chat title
      const titlePrefix = input.trim().slice(0, 40).replace(/\n/g, ' ').trim();
      if (titlePrefix) {
        updateChat(chatId, { title: titlePrefix }).then(() => {
          if (onChatTitleChange) {
            onChatTitleChange(titlePrefix);
          }
        }).catch(console.error);
      }
    }

    const updatedMessages = [...messages, userMessage];
    setMessages(updatedMessages);
    setInput('');
    // Reset textarea height
    if (textareaRef.current) {
      textareaRef.current.style.height = 'auto';
    }
    setIsLoading(true);

    try {
      const apiKey = import.meta.env.VITE_OPENROUTER_API_KEY || '';
      if (!apiKey) {
        throw new Error('OpenRouter API key is not configured. Please set VITE_OPENROUTER_API_KEY in your .env file.');
      }

      // Call OpenRouter API with tool support
      const response = await fetch('https://openrouter.ai/api/v1/chat/completions', {
        signal: abortController.signal,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${apiKey}`,
          'HTTP-Referer': window.location.origin,
          'X-Title': 'SheetGrid',
        },
        body: JSON.stringify({
          model: selectedModel,
          messages: [
            ...updatedMessages.filter(m => m.role !== 'system').map((m) => {
              // Add context to the latest user message for AI
              if (m.role === 'user' && m.content === input.trim() && context) {
                return { role: 'user' as const, content: `${m.content}\n\n[Current Sheet Context]\n${context}` };
              }
              return { role: m.role, content: m.content };
            }),
          ],
          tools: getTools(),
          tool_choice: 'auto',
          max_tokens: 4096,
        }),
      });

      if (!response.ok) {
        let errorMessage = `HTTP ${response.status}: ${response.statusText}`;
        try {
          const errorData = await response.json();
          errorMessage = errorData.error?.message || errorData.message || errorMessage;
          console.error('OpenRouter API Error:', {
            status: response.status,
            statusText: response.statusText,
            error: errorData,
          });
        } catch {
          const text = await response.text();
          console.error('OpenRouter API Error (non-JSON):', {
            status: response.status,
            statusText: response.statusText,
            body: text.substring(0, 500),
          });
        }

        if (response.status === 401) {
          throw new Error(`OpenRouter authentication failed. Please check your API key in .env file. Error: ${errorMessage}`);
        }
        throw new Error(`Failed to get response from OpenRouter: ${errorMessage}`);
      }

      const data = await response.json();

      // Handle tool calls
      console.log('AI response:', data.choices?.[0]?.message);
      if (data.choices?.[0]?.message?.tool_calls && data.choices[0].message.tool_calls.length > 0) {
        console.log('Tool calls:', data.choices[0].message.tool_calls);
        const assistantMessage: Message = {
          role: 'assistant',
          content: 'Executing operations...',
          timestamp: new Date(),
        };
        setMessages((prev) => [...prev, assistantMessage]);

        // Execute all tool calls
        const toolResults: Array<{ name: string; result: string }> = [];
        for (const toolCall of data.choices[0].message.tool_calls) {
          // Check if request was aborted
          if (abortController.signal.aborted) {
            console.log('Request aborted, stopping tool execution');
            return;
          }
          console.log('Executing tool:', toolCall.function.name, toolCall.function.arguments);
          const result = await executeTool({
            name: toolCall.function.name,
            arguments: toolCall.function.arguments ? JSON.parse(toolCall.function.arguments) : {},
          });
          console.log('Tool result:', result);
          toolResults.push({ name: toolCall.function.name, result });
        }

        // Prepare tool responses for OpenRouter
        const toolResponses = toolResults.map((tr) => ({
          role: 'tool' as const,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          tool_call_id: data.choices[0].message.tool_calls.find((tc: any) => tc.function.name === tr.name).id,
          content: tr.result,
        }));

        // Check if request was aborted before sending follow-up
        if (abortController.signal.aborted) {
          console.log('Request aborted, skipping follow-up request');
          return;
        }

        // Send the results back to OpenRouter for a final response
        const apiKey2 = import.meta.env.VITE_OPENROUTER_API_KEY || '';
        if (!apiKey2) {
          throw new Error('OpenRouter API key is not configured. Please set VITE_OPENROUTER_API_KEY in your .env file.');
        }

        const followUpResponse = await fetch('https://openrouter.ai/api/v1/chat/completions', {
          signal: abortController.signal,
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey2}`,
            'HTTP-Referer': window.location.origin,
            'X-Title': 'SheetGrid',
          },
          body: JSON.stringify({
            model: selectedModel,
            messages: [
              ...updatedMessages.filter(m => m.role !== 'system').map((m) => {
                // Add context to the latest user message for AI
                if (m.role === 'user' && m.content === input.trim() && context) {
                  return { role: 'user' as const, content: `${m.content}\n\n[Current Sheet Context]\n${context}` };
                }
                return { role: m.role, content: m.content };
              }),
              { role: 'assistant', tool_calls: data.choices[0].message.tool_calls },
              ...toolResponses,
            ],
            tools: getTools(),
            tool_choice: 'auto',
            max_tokens: 4096,
          }),
        });

        if (!followUpResponse.ok) {
          let errorMessage = `HTTP ${followUpResponse.status}: ${followUpResponse.statusText}`;
          try {
            const errorData = await followUpResponse.json();
            errorMessage = errorData.error?.message || errorData.message || errorMessage;
            console.error('OpenRouter Follow-up API Error:', {
              status: followUpResponse.status,
              statusText: followUpResponse.statusText,
              error: errorData,
            });
          } catch {
            const text = await followUpResponse.text();
            console.error('OpenRouter Follow-up API Error (non-JSON):', {
              status: followUpResponse.status,
              statusText: followUpResponse.statusText,
              body: text.substring(0, 500),
            });
          }

          if (followUpResponse.status === 401) {
            throw new Error(`OpenRouter authentication failed. Please check your API key in .env file. Error: ${errorMessage}`);
          }
          throw new Error(`Failed to get follow-up response from OpenRouter: ${errorMessage}`);
        }

        const followUpData = await followUpResponse.json();
        console.log('Follow-up response complete:', followUpData);
        console.log('Follow-up AI response:', followUpData.choices?.[0]?.message);
        console.log('Follow-up tool_calls:', followUpData.choices?.[0]?.message?.tool_calls);
        console.log('Follow-up tool_calls length:', followUpData.choices?.[0]?.message?.tool_calls?.length);

        // Check if follow-up has tool calls (multi-step tool execution)
        if (followUpData.choices?.[0]?.message?.tool_calls && followUpData.choices[0].message.tool_calls.length > 0) {
          console.log('Follow-up tool calls:', JSON.stringify(followUpData.choices[0].message.tool_calls, null, 2));
          const followUpToolResults: Array<{ name: string; result: string }> = [];
          for (const toolCall of followUpData.choices[0].message.tool_calls) {
            // Check if request was aborted
            if (abortController.signal.aborted) {
              console.log('Request aborted, stopping follow-up tool execution');
              return;
            }
            console.log('Executing follow-up tool:', toolCall.function.name, toolCall.function.arguments);
            const result = await executeTool({
              name: toolCall.function.name,
              arguments: toolCall.function.arguments ? JSON.parse(toolCall.function.arguments) : {},
            });
            console.log('Follow-up tool result:', result);
            followUpToolResults.push({ name: toolCall.function.name, result });
          }

          // Send another follow-up if needed (multi-round tool calling)
          const followUpToolResponses = followUpToolResults.map((tr) => ({
            role: 'tool' as const,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            tool_call_id: followUpData.choices[0].message.tool_calls.find((tc: any) => tc.function.name === tr.name).id,
            content: tr.result,
          }));

          // Check if request was aborted before sending final follow-up
          if (abortController.signal.aborted) {
            console.log('Request aborted, skipping final follow-up request');
            return;
          }

          const finalFollowUpResponse = await fetch('https://openrouter.ai/api/v1/chat/completions', {
            signal: abortController.signal,
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': `Bearer ${apiKey2}`,
              'HTTP-Referer': window.location.origin,
              'X-Title': 'SheetGrid',
            },
            body: JSON.stringify({
              model: selectedModel,
              messages: [
                ...updatedMessages.filter(m => m.role !== 'system').map((m) => {
                  // Add context to the latest user message for AI
                  if (m.role === 'user' && m.content === input.trim() && context) {
                    return { role: 'user' as const, content: `${m.content}\n\n[Current Sheet Context]\n${context}` };
                  }
                  return { role: m.role, content: m.content };
                }),
                { role: 'assistant', tool_calls: data.choices[0].message.tool_calls },
                ...toolResponses,
                { role: 'assistant', tool_calls: followUpData.choices[0].message.tool_calls },
                ...followUpToolResponses,
              ],
              tools: getTools(),
              tool_choice: 'auto',
              max_tokens: 4096,
            }),
          });

          if (!finalFollowUpResponse.ok) {
            let errorMessage = `HTTP ${finalFollowUpResponse.status}: ${finalFollowUpResponse.statusText}`;
            try {
              const errorData = await finalFollowUpResponse.json();
              errorMessage = errorData.error?.message || errorData.message || errorMessage;
              console.error('OpenRouter Final Follow-up API Error:', {
                status: finalFollowUpResponse.status,
                statusText: finalFollowUpResponse.statusText,
                error: errorData,
              });
            } catch {
              const text = await finalFollowUpResponse.text();
              console.error('OpenRouter Final Follow-up API Error (non-JSON):', {
                status: finalFollowUpResponse.status,
                statusText: finalFollowUpResponse.statusText,
                body: text.substring(0, 500),
              });
            }

            if (finalFollowUpResponse.status === 401) {
              throw new Error(`OpenRouter authentication failed. Please check your API key in .env file. Error: ${errorMessage}`);
            }
            throw new Error(`Failed to get final follow-up response from OpenRouter: ${errorMessage}`);
          }

          const finalFollowUpData = await finalFollowUpResponse.json();
          console.log('Final follow-up data:', finalFollowUpData);
          const finalMessage: Message = {
            role: 'assistant',
            content: finalFollowUpData.choices?.[0]?.message?.content || 'Operations completed successfully.',
            timestamp: new Date(),
          };
          setMessages((prev) => [...prev, finalMessage]);

          // Trigger save after operations
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          if (typeof window !== 'undefined' && (window as any).saveWorkbookData) {
            setTimeout(async () => {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              await (window as any).saveWorkbookData();
            }, 500);
          }
        } else {
          // No more tool calls, final response
          const finalMessage: Message = {
            role: 'assistant',
            content: followUpData.choices?.[0]?.message?.content || 'Operations completed successfully.',
            timestamp: new Date(),
          };
          setMessages((prev) => [...prev, finalMessage]);

          // Trigger save after operations
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          if (typeof window !== 'undefined' && (window as any).saveWorkbookData) {
            setTimeout(async () => {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              await (window as any).saveWorkbookData();
            }, 500);
          }
        }
      } else {
        // No tool calls, direct response
        const assistantMessage: Message = {
          role: 'assistant',
          content: data.choices?.[0]?.message?.content || 'No response',
          timestamp: new Date(),
        };
        setMessages((prev) => [...prev, assistantMessage]);
      }
    } catch (error) {
      console.error('Chat error:', error);
      const errorMessage: Message = {
        role: 'assistant',
        content: `Sorry, I encountered an error. ${error instanceof Error ? error.message : 'Unknown error'}`,
        timestamp: new Date(),
      };
      setMessages((prev) => [...prev, errorMessage]);
    } finally {
      setIsLoading(false);
      abortControllerRef.current = null;
    }
  };

  return (
    <div className="flex flex-col h-full bg-[#FAFAFA]">
      {/* Top Bar - Cursor style with browser tabs */}
      <div className="border-b border-[#E0E0E0] bg-white">
        {/* Tabs container */}
        <div className="flex items-end overflow-x-auto scrollbar-hide" style={{ scrollbarWidth: 'none', msOverflowStyle: 'none' }}>
          <div className="flex items-end min-w-full">
            {allChats.map((chat) => (
              <div
                key={chat.id}
                className={`group relative flex items-center border-b-2 ${chat.id === chatId
                  ? 'text-[#0066CC] border-[#0066CC] bg-white'
                  : 'text-[#666666] border-transparent hover:text-[#333333] hover:border-[#D0D0D0]'
                  }`}
              >
                <button
                  onClick={() => {
                    if (onSelectChat) {
                      onSelectChat(chat.id);
                    }
                  }}
                  className="px-4 py-2.5 text-sm font-medium transition-colors whitespace-nowrap flex items-center gap-2"
                >
                  <span className="truncate max-w-[200px] block">{chat.title || 'New Chat'}</span>
                </button>
                {onDeleteChat && (
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      if (window.confirm('Are you sure you want to delete this chat?')) {
                        onDeleteChat(chat.id);
                      }
                    }}
                    className="opacity-0 group-hover:opacity-100 p-1 mr-2 hover:bg-[#F0F0F0] rounded transition-all"
                    title="Delete chat"
                  >
                    <svg className="w-3.5 h-3.5 text-[#999999] hover:text-[#666666]" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
                      <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
                    </svg>
                  </button>
                )}
              </div>
            ))}

            {/* Plus button for new chat */}
            <button
              onClick={(e) => {
                e.stopPropagation();
                onCreateNewChat();
              }}
              className="p-2 text-[#666666] hover:text-[#333333] hover:bg-[#F0F0F0] transition-colors border-b-2 border-transparent"
              title="New Chat"
            >
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M12 4v16m8-8H4" />
              </svg>
            </button>
          </div>
        </div>
      </div>

      {/* Messages */}
      <div className="flex-1 overflow-y-auto">
        <div className="max-w-3xl mx-auto px-4 py-6">
          {messages
            .filter((m) => m.role !== 'system')
            .map((message, idx) => (
              <div
                key={idx}
                className={`mb-8 ${message.role === 'user' ? 'flex justify-end' : ''}`}
              >
                <div className={`flex gap-4 ${message.role === 'user' ? 'flex-row-reverse' : ''} max-w-[85%] ${message.role === 'user' ? 'ml-auto' : ''}`}>
                  {/* Avatar */}
                  <div className={`shrink-0 w-8 h-8 rounded-full flex items-center justify-center text-xs font-medium ${message.role === 'user'
                    ? 'bg-[#0066CC] text-white'
                    : 'bg-[#E8E8E8] text-[#666666]'
                    }`}>
                    {message.role === 'user' ? 'U' : 'AI'}
                  </div>

                  {/* Message content */}
                  <div className="flex-1 min-w-0">
                    <div className={`rounded-lg px-4 py-2.5 ${message.role === 'user'
                      ? 'bg-[#0066CC] text-white'
                      : 'bg-white text-[#333333] border border-[#E0E0E0]'
                      }`}>
                      <p className="text-sm leading-relaxed whitespace-pre-wrap break-words">
                        {cleanMessageContent(message.content)}
                      </p>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          {isLoading && (
            <div className="mb-8">
              <div className="flex gap-4 max-w-[85%]">
                <div className="shrink-0 w-8 h-8 rounded-full flex items-center justify-center text-xs font-medium bg-[#E8E8E8] text-[#666666]">
                  AI
                </div>
                <div className="flex-1 min-w-0">
                  <div className="rounded-lg px-4 py-2.5 bg-white border border-[#E0E0E0]">
                    <div className="flex items-center gap-1.5">
                      <div className="w-1.5 h-1.5 bg-[#999999] rounded-full animate-bounce" style={{ animationDelay: '0s' }} />
                      <div className="w-1.5 h-1.5 bg-[#999999] rounded-full animate-bounce" style={{ animationDelay: '0.2s' }} />
                      <div className="w-1.5 h-1.5 bg-[#999999] rounded-full animate-bounce" style={{ animationDelay: '0.4s' }} />
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}
          <div ref={messagesEndRef} />
        </div>
      </div>

      {/* Input Area - Cursor-style */}
      <div className="border-t border-[#E0E0E0] bg-white">
        <div className="w-full px-4 py-3">
          {/* Input field */}
          <div className="mb-2">
            <textarea
              ref={textareaRef}
              value={input}
              onChange={(e) => setInput(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && !e.shiftKey) {
                  e.preventDefault();
                  sendMessage();
                }
              }}
              placeholder="Chat with your AI"
              className="w-full px-3 py-2 border border-[#D0D0D0] rounded resize-none focus:outline-none focus:ring-1 focus:ring-[#0066CC] focus:border-[#0066CC] text-sm text-[#333333] placeholder-[#999999] bg-white overflow-hidden"
              style={{ minHeight: '32px', maxHeight: '120px' }}
              disabled={isLoading}
            />
          </div>

          {/* Tools row */}
          <div className="flex items-center justify-between">
            {/* Left side - Model selector */}
            <div className="flex items-center gap-2">
              <div className="relative" ref={modelDropdownRef}>
                <button
                  onClick={() => setIsModelDropdownOpen(!isModelDropdownOpen)}
                  className="text-xs text-[#666666] hover:text-[#333333] flex items-center gap-1 px-2 py-1 rounded hover:bg-[#F0F0F0] transition-colors font-medium"
                >
                  <span>{selectedModel}</span>
                  <svg
                    className={`w-3 h-3 transition-transform ${isModelDropdownOpen ? 'transform rotate-180' : ''}`}
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                    strokeWidth={2}
                  >
                    <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
                  </svg>
                </button>

                {isModelDropdownOpen && (
                  <div className="absolute z-50 bottom-full mb-2 w-56 bg-white border border-[#D0D0D0] rounded-md shadow-lg max-h-64 overflow-auto">
                    {availableModels.length === 0 ? (
                      <div className="px-3 py-2 text-sm text-[#999999]">No models found</div>
                    ) : (
                      availableModels.map((model) => (
                        <button
                          key={model}
                          onClick={() => {
                            setSelectedModel(model);
                            setIsModelDropdownOpen(false);
                          }}
                          className={`w-full px-3 py-2 text-left text-sm hover:bg-[#F5F5F5] transition-colors ${model === selectedModel
                            ? 'bg-[#E6F2FF] text-[#0066CC] font-medium'
                            : 'text-[#333333]'
                            }`}
                        >
                          {model}
                        </button>
                      ))
                    )}
                  </div>
                )}
              </div>
            </div>

            {/* Right side - Action buttons */}
            <div className="flex items-center gap-1">
              {/* Hidden file input */}
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileUpload}
                className="hidden"
              />

              {/* Upload document button */}
              <button
                onClick={handleUploadClick}
                disabled={isUploading}
                className="p-1.5 text-[#666666] hover:text-[#333333] hover:bg-[#F0F0F0] rounded transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                title={isUploading ? 'Uploading file...' : 'Upload Excel file (.xlsx, .xls)'}
              >
                {isUploading ? (
                  <svg className="w-4 h-4 animate-spin" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                  </svg>
                ) : (
                  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                  </svg>
                )}
              </button>

              {/* Send/Stop button */}
              {isLoading ? (
                <button
                  onClick={stopMessage}
                  className="p-1.5 text-[#666666] hover:text-[#333333] hover:bg-[#F0F0F0] rounded transition-colors"
                  title="Stop generating"
                >
                  <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 24 24">
                    <rect x="6" y="6" width="12" height="12" rx="1" />
                  </svg>
                </button>
              ) : (
                <button
                  onClick={sendMessage}
                  disabled={!input.trim()}
                  className="p-1.5 text-[#0066CC] hover:text-[#0052A3] disabled:text-[#CCCCCC] disabled:cursor-not-allowed transition-colors rounded hover:bg-[#F0F7FF] disabled:hover:bg-transparent"
                  title="Send message"
                >
                  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8" />
                  </svg>
                </button>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
