import { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { loadChatMessages, saveAllChatMessages, updateChat } from '../src/utils/indexeddb';
import { useUserApiKeys } from '../src/hooks/useUserApiKeys';

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
  // chatTitle is passed from parent but not currently used in this component
  void chatTitle;

  // Get user's API keys
  const { openRouterKey, univerMcpKey } = useUserApiKeys();
  // System message should not be persisted
  const systemMessage: Message = {
    role: 'system',
    content: `You are a helpful assistant that can edit Excel spreadsheets using tools. 

AVAILABLE CAPABILITIES:
📊 Data Operations: set_range_data (set cell values), get_range_data (read cell data), search_cells (find content), auto_fill (fill patterns), format_brush (copy formatting)
📋 Sheet Management: create_sheet, delete_sheet, rename_sheet, activate_sheet, move_sheet, set_sheet_display_status, get_sheets, get_active_unit_id
🏗️ Structure Operations: insert_rows/columns, delete_rows/columns, set_cell_dimensions (row height/column width), set_merge (merge cells)
🎨 Formatting & Styling: set_range_style, add_conditional_formatting_rule, set_conditional_formatting_rule, delete_conditional_formatting_rule, get_conditional_formatting_rules
✅ Data Validation: add_data_validation_rule, set_data_validation_rule, delete_data_validation_rule, get_data_validation_rules
🔍 Utility Functions: get_activity_status (workbook info), scroll_and_screenshot

CRITICAL RULE - NO CODE IN TOOL CALLS: When you need to generate random numbers or any calculated values, you MUST compute them yourself and provide ONLY the actual results as numbers/strings. NEVER include any code in your tool calls.
- NEVER write: "Math.floor(Math.random() * 101)" - write the actual number like: 47
- NEVER write: "[Math.round(Math.random() * 1000) for i in range(30)]" - write the array of actual numbers: [342, 891, 123, 567, ...]
- NEVER include ANY programming syntax (Math., random(), for loops, etc.) in tool call arguments
- ONLY provide the final calculated values as pure JSON data
- If asked for 5 random numbers 0-100, generate something like: [47, 82, 15, 93, 6]
- Example CORRECT usage for dates Nov 1-5: ["2023-11-01", "2023-11-02", "2023-11-03", "2023-11-04", "2023-11-05"]
- Example CORRECT usage for random profits: [47, 82, 15, 93, 6]

TOOL USAGE GUIDELINES:
- Before using tools, understand what you're working with by checking sheet structure first
- Use get_sheets to see available sheets and their names
- Use get_range_data to understand the data format before making changes
- Validate your tool arguments match the expected types (strings, numbers, arrays, objects)
- If a tool fails, read the error message carefully and adjust your approach
- Be efficient with tool calls - group related operations when possible

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
  const [isDownloading, setIsDownloading] = useState(false);
  const [selectedModel, setSelectedModel] = useState('anthropic/claude-3-haiku');
  const [isModelDropdownOpen, setIsModelDropdownOpen] = useState(false);
  const [showInfo, setShowInfo] = useState(false);
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

  // Configuration constants
  const MAX_TOOLS_PER_ROUND = 15;
  const MAX_TOOL_CALL_DEPTH = 10; // Maximum recursive depth for tool calling (state-of-the-art unlimited depth support)
  const API_RETRY_ATTEMPTS = 3; // Retry API calls on failure

  // Function to clean context from message content
  const cleanMessageContent = (content: string | undefined | null): string => {
    // Handle undefined/null content (e.g., assistant messages with tool_calls might not have content)
    if (!content) {
      return '';
    }
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
    if (messages.length === 0) return; // Don't save if no messages (except system message)

    // Use a longer debounce to ensure all messages from tool calls are captured
    const saveTimeout = setTimeout(async () => {
      try {
        // Filter out system message before saving
        const messagesToSave = messages.filter(msg => msg.role !== 'system');
        if (messagesToSave.length > 0) {
          console.log(`💾 Saving ${messagesToSave.length} messages to IndexedDB for chat ${chatId}`);
          await saveAllChatMessages(chatId, messagesToSave);
          console.log(`✅ Successfully saved ${messagesToSave.length} messages`);
        }
      } catch (error) {
        console.error('❌ Error saving chat history:', error);
      }
    }, 2000); // Increased debounce to 2 seconds to capture all messages from recursive tool calls

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
        if (!univerMcpKey) {
          console.warn('No MCP API key configured. Please add your Univer MCP API key in your profile settings. Using local tools only.');
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
              'Authorization': `Bearer ${univerMcpKey}`,
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
  }, [univerMcpKey]);

  // Return only MCP tools from server
  const getTools = () => {
    if (mcpTools.length === 0) {
      console.warn('No MCP tools available yet. Make sure the MCP API key is configured.');
      return [];
    }
    console.log(`getTools: returning ${mcpTools.length} tools from MCP server`);
    return mcpTools;
  };

  const executeTool = async (toolCall: ToolCall, retries = 2): Promise<string> => {
    const { name, arguments: args } = toolCall;

    // Check if this is an MCP tool (from server)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const isMcpTool = mcpTools.some((tool: any) => tool.function.name === name);

    if (!isMcpTool) {
      return JSON.stringify({ error: `Tool "${name}" is not available. Only MCP server tools are supported.` });
    }

    console.log(`Tool "${name}": routing to MCP server (attempt ${3 - retries}/${3})`, 'args:', JSON.stringify(args));

    if (!univerMcpKey) {
      return JSON.stringify({ error: 'MCP API key not configured. Please add your Univer MCP API key in your profile settings.' });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const sessionId = (window as any).univerSessionId || 'default';

    const executeWithTimeout = async (): Promise<Response> => {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 30000); // 30s timeout

      try {
        const response = await fetch(`https://mcp.univer.ai/mcp/?univer_session_id=${sessionId}`, {
          method: 'POST',
          signal: controller.signal,
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${univerMcpKey}`,
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
        clearTimeout(timeoutId);
        return response;
      } catch (error) {
        clearTimeout(timeoutId);
        throw error;
      }
    };

    try {
      const response = await executeWithTimeout();

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

      // MCP tools/call returns { result: { content: [...] } } or { error: {...} }
      // Check for errors in the response
      if (data.error) {
        const errorMsg = data.error.message || JSON.stringify(data.error);
        // Retry logic for tool-level errors
        if (retries > 0) {
          console.warn(`Tool "${name}" returned an error, retrying... (${retries} attempts left)`, errorMsg);
          await new Promise(resolve => setTimeout(resolve, 1000 * (3 - retries))); // 1s, 2s delays
          return executeTool(toolCall, retries - 1);
        }
        console.error(`Tool "${name}" failed after all retries:`, errorMsg);
        return JSON.stringify({ error: `MCP tool execution failed: ${errorMsg}` });
      }

      if (data.result && data.result.content) {
        const content = data.result.content;
        // Content is an array, extract text
        if (Array.isArray(content) && content.length > 0) {
          const resultText = content[0].text || JSON.stringify(content[0]);
          // Check if the result text contains an error
          try {
            const parsedResult = JSON.parse(resultText);
            if (parsedResult.error) {
              // Tool returned an error in the result - retry if we have retries left
              if (retries > 0) {
                console.warn(`Tool "${name}" result contains an error, retrying... (${retries} attempts left)`, parsedResult.error);
                await new Promise(resolve => setTimeout(resolve, 1000 * (3 - retries))); // 1s, 2s delays
                return executeTool(toolCall, retries - 1);
              }
            }
          } catch {
            // Not JSON or parse failed, check if it's a plain error string
            if (resultText.toLowerCase().includes('error') && retries > 0) {
              console.warn(`Tool "${name}" result may contain an error, retrying... (${retries} attempts left)`, resultText.substring(0, 100));
              await new Promise(resolve => setTimeout(resolve, 1000 * (3 - retries)));
              return executeTool(toolCall, retries - 1);
            }
          }
          return resultText;
        }
        return JSON.stringify(content);
      }
      return JSON.stringify(data.result || { success: true });
    } catch (error) {
      // Retry logic with exponential backoff
      if (retries > 0) {
        console.warn(`Tool "${name}" failed, retrying... (${retries} attempts left)`);
        await new Promise(resolve => setTimeout(resolve, 1000 * (3 - retries))); // 1s, 2s delays
        return executeTool(toolCall, retries - 1);
      }
      const errorMsg = error instanceof Error ? error.message : String(error);
      console.error(`Tool "${name}" failed after all retries:`, errorMsg);
      return JSON.stringify({ error: `MCP tool execution failed: ${errorMsg}` });
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
                  // getRange(row, column, numRows, numColumns) - note: numRows is the COUNT, not end index
                  const numRows = batch.length;
                  const numCols = maxCols;
                  const range = firstSheet.getRange(startRow, 0, numRows, numCols);
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

  // Helper function to retry API calls with exponential backoff (state-of-the-art)
  const fetchWithRetry = async (
    url: string,
    options: RequestInit,
    retries: number
  ): Promise<Response> => {
    let lastError: Error | null = null;

    for (let attempt = 0; attempt < retries; attempt++) {
      try {
        const response = await fetch(url, options);
        if (response.ok) {
          return response;
        }

        // Don't retry on client errors (4xx), only retry on server errors (5xx) and network errors
        if (response.status >= 400 && response.status < 500 && response.status !== 408) {
          // Read error response body to get detailed error message
          let errorMessage = `Client error: ${response.status} ${response.statusText}`;
          try {
            const errorData = await response.json();
            const openRouterError = errorData.error?.message || errorData.message || errorData.error || JSON.stringify(errorData);
            errorMessage = `Client error: ${response.status} ${response.statusText} - ${openRouterError}`;
            console.error('OpenRouter API Error Details:', {
              status: response.status,
              statusText: response.statusText,
              error: errorData,
              url: url,
              attempt: attempt + 1,
            });
          } catch (parseError) {
            // If JSON parsing fails, try to get text response
            try {
              const errorText = await response.text();
              errorMessage = `Client error: ${response.status} ${response.statusText} - ${errorText.substring(0, 500)}`;
              console.error('OpenRouter API Error (non-JSON):', {
                status: response.status,
                statusText: response.statusText,
                body: errorText.substring(0, 500),
                url: url,
                attempt: attempt + 1,
              });
            } catch (textError) {
              // If even text parsing fails, use status only
              console.error('OpenRouter API Error (could not parse response):', {
                status: response.status,
                statusText: response.statusText,
                url: url,
                attempt: attempt + 1,
              });
            }
          }
          throw new Error(errorMessage);
        }

        // For server errors (5xx), read error but still retry
        let serverErrorMessage = `Server error: ${response.status} ${response.statusText}`;
        try {
          const errorData = await response.json();
          const openRouterError = errorData.error?.message || errorData.message || errorData.error;
          if (openRouterError) {
            serverErrorMessage = `Server error: ${response.status} ${response.statusText} - ${openRouterError}`;
          }
          console.warn('OpenRouter API Server Error:', {
            status: response.status,
            statusText: response.statusText,
            error: errorData,
            url: url,
            attempt: attempt + 1,
          });
        } catch {
          // Ignore parsing errors for server errors, we'll retry anyway
        }
        lastError = new Error(serverErrorMessage);
      } catch (error) {
        lastError = error instanceof Error ? error : new Error(String(error));

        // Check if this is a client error - don't retry, throw immediately
        if (error instanceof Error && error.message.includes('Client error')) {
          console.error('Client error detected, not retrying:', error.message);
          throw error;
        }

        // For other errors, retry if we have attempts left
        if (attempt < retries - 1) {
          // Exponential backoff: 1s, 2s, 4s
          const delay = Math.pow(2, attempt) * 1000;
          console.warn(`API call failed, retrying in ${delay}ms... (attempt ${attempt + 1}/${retries})`, lastError.message);
          await new Promise(resolve => setTimeout(resolve, delay));
        }
      }
    }

    throw lastError || new Error('API call failed after all retries');
  };

  // Recursive function to handle tool calls with unlimited depth (state-of-the-art like Cursor)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const handleToolCallsRecursive = async (
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    toolCalls: any[],
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    conversationMessages: any[],
    depth: number = 0,
    abortSignal: AbortSignal,
    onProgress?: (depth: number, toolCount: number) => void
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ): Promise<{ messages: any[]; finalResponse: string }> => {
    // Safety check to prevent infinite loops
    if (depth >= MAX_TOOL_CALL_DEPTH) {
      console.warn(`Maximum tool call depth (${MAX_TOOL_CALL_DEPTH}) reached`);
      return {
        messages: conversationMessages,
        finalResponse: 'Maximum tool call depth reached. Please try breaking down your request into smaller steps.',
      };
    }

    // Limit tool calls per round for performance
    const toolCallsToExecute = toolCalls.slice(0, MAX_TOOLS_PER_ROUND);
    if (toolCalls.length > MAX_TOOLS_PER_ROUND) {
      console.warn(`Too many tool calls (${toolCalls.length}), limiting to ${MAX_TOOLS_PER_ROUND}`);
    }

    // Notify progress if callback provided
    if (onProgress) {
      onProgress(depth, toolCallsToExecute.length);
    }

    // Execute all tool calls - use Promise.allSettled for better error handling
    const toolResults: Array<{ name: string; result: string; tool_call_id: string }> = [];

    const toolPromises = toolCallsToExecute.map(async (toolCall) => {
      if (abortSignal.aborted) {
        return null;
      }

      console.log(`[Depth ${depth}] Executing tool: ${toolCall.function.name}`);
      let parsedArgs = {};
      if (toolCall.function.arguments) {
        try {
          parsedArgs = JSON.parse(toolCall.function.arguments);
          // Handle double-encoded JSON
          if (typeof parsedArgs === 'string') {
            parsedArgs = JSON.parse(parsedArgs);
          }
        } catch (e) {
          console.error('Failed to parse tool arguments:', e, toolCall.function.arguments);
          parsedArgs = {};
        }
      }

      try {
        const result = await executeTool({
          name: toolCall.function.name,
          arguments: parsedArgs,
        });
        return {
          name: toolCall.function.name,
          result,
          tool_call_id: toolCall.id,
        };
      } catch (error) {
        console.error(`Tool ${toolCall.function.name} failed:`, error);
        return {
          name: toolCall.function.name,
          result: JSON.stringify({ error: error instanceof Error ? error.message : String(error) }),
          tool_call_id: toolCall.id,
        };
      }
    });

    const settledResults = await Promise.allSettled(toolPromises);
    settledResults.forEach((result, index) => {
      if (result.status === 'fulfilled' && result.value) {
        toolResults.push(result.value);
      } else if (result.status === 'rejected') {
        // Handle failed tool execution
        const toolCall = toolCallsToExecute[index];
        toolResults.push({
          name: toolCall.function.name,
          result: JSON.stringify({ error: result.reason?.message || 'Tool execution failed' }),
          tool_call_id: toolCall.id,
        });
      }
    });

    // Prepare tool responses for the API - use Map for O(1) lookup instead of find()
    const toolCallMap = new Map(toolCallsToExecute.map(tc => [tc.id, tc]));
    const toolResponses = toolResults.map((tr) => ({
      role: 'tool' as const,
      tool_call_id: tr.tool_call_id,
      content: tr.result,
    }));

    // Add assistant message with tool calls and tool responses to conversation
    const updatedMessages = [
      ...conversationMessages,
      {
        role: 'assistant' as const,
        tool_calls: toolCallsToExecute,
      },
      ...toolResponses,
    ];

    // Make API call to get next response with retry logic
    try {
      const response = await fetchWithRetry(
        'https://openrouter.ai/api/v1/chat/completions',
        {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${openRouterKey}`,
            'HTTP-Referer': window.location.origin,
            'X-Title': 'SheetGrid',
          },
          body: (() => {
            const requestBody = {
              model: selectedModel,
              messages: [
                systemMessage,
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                ...updatedMessages.filter((m: any) => m.role !== 'system').map((m: any) => {
                  // Add context to the latest user message for AI
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const userInput = conversationMessages.find((msg: any) => msg.role === 'user')?.content;
                  if (m.role === 'user' && userInput && m.content === userInput && depth === 0) {
                    const context = getSheetContext();
                    if (context) {
                      return { role: 'user' as const, content: `${m.content}\n\n[Current Sheet Context]\n${context}` };
                    }
                  }

                  // For assistant messages with tool_calls, preserve the tool_calls field
                  if (m.role === 'assistant' && m.tool_calls) {
                    return {
                      role: 'assistant' as const,
                      tool_calls: m.tool_calls,
                      ...(m.content ? { content: m.content } : {}), // Only include content if it exists
                    };
                  }

                  // For tool messages, preserve the tool_call_id
                  if (m.role === 'tool') {
                    return {
                      role: 'tool' as const,
                      tool_call_id: m.tool_call_id,
                      content: m.content,
                    };
                  }

                  // For other messages (user, assistant without tool_calls), just return role and content
                  return { role: m.role, content: m.content };
                }),
              ],
              tools: getTools(),
              tool_choice: 'auto',
              max_tokens: 4096,
            };

            // Log request body for debugging (truncated to avoid console spam)
            console.log(`[Depth ${depth}] API Request:`, {
              model: requestBody.model,
              messageCount: requestBody.messages.length,
              toolCount: requestBody.tools.length,
              lastMessageRole: requestBody.messages[requestBody.messages.length - 1]?.role,
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              lastMessageHasToolCalls: !!(requestBody.messages[requestBody.messages.length - 1] as any)?.tool_calls,
            });

            return JSON.stringify(requestBody);
          })(),
          signal: abortSignal,
        },
        API_RETRY_ATTEMPTS
      );

      const data = await response.json();

      // Check if there are more tool calls (recursive case)
      if (data.choices?.[0]?.message?.tool_calls && data.choices[0].message.tool_calls.length > 0) {
        console.log(`[Depth ${depth}] More tool calls detected (${data.choices[0].message.tool_calls.length}), recursing...`);
        return handleToolCallsRecursive(
          data.choices[0].message.tool_calls,
          updatedMessages,
          depth + 1,
          abortSignal,
          onProgress
        );
      }

      // No more tool calls, return final response
      const finalResponse = data.choices?.[0]?.message?.content || 'Operations completed successfully.';
      const finalAssistantMessage = {
        role: 'assistant' as const,
        content: finalResponse,
        timestamp: new Date(), // Add timestamp for proper saving
      };
      return {
        messages: [
          ...updatedMessages,
          finalAssistantMessage,
        ],
        finalResponse,
      };
    } catch (error) {
      console.error(`API call failed at depth ${depth}:`, error);
      throw error;
    }
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
      if (!openRouterKey) {
        throw new Error('OpenRouter API key is not configured. Please add your OpenRouter API key in your profile settings.');
      }

      // Call OpenRouter API with tool support
      const response = await fetch('https://openrouter.ai/api/v1/chat/completions', {
        signal: abortController.signal,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${openRouterKey}`,
          'HTTP-Referer': window.location.origin,
          'X-Title': 'SheetGrid',
        },
        body: JSON.stringify({
          model: selectedModel,
          messages: [
            systemMessage,
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

      // Handle tool calls with recursive depth support (state-of-the-art)
      console.log('AI response:', data.choices?.[0]?.message);
      if (data.choices?.[0]?.message?.tool_calls && data.choices[0].message.tool_calls.length > 0) {
        console.log('Tool calls:', data.choices[0].message.tool_calls);
        const assistantMessage: Message = {
          role: 'assistant',
          content: 'Executing operations...',
          timestamp: new Date(),
        };
        setMessages((prev) => [...prev, assistantMessage]);

        // Use recursive function for unlimited tool call depth (state-of-the-art)
        try {
          const { messages: finalMessages, finalResponse } = await handleToolCallsRecursive(
            data.choices[0].message.tool_calls,
            updatedMessages,
            0, // Start at depth 0
            abortController.signal,
            (depth, toolCount) => {
              // Progress callback - could update UI here if needed
              console.log(`[Progress] Depth: ${depth}, Tools: ${toolCount}`);
            }
          );

          // Update messages with all the conversation including tool calls and final response
          // Remove the "Executing operations..." message and add the actual conversation
          setMessages((prev) => {
            // Remove the temporary "Executing operations..." message
            const withoutTemp = prev.filter((msg, idx) => !(idx === prev.length - 1 && msg.content === 'Executing operations...'));
            
            // Get the last message timestamp from existing messages to maintain order
            const lastTimestamp = withoutTemp.length > 0 
              ? withoutTemp[withoutTemp.length - 1].timestamp.getTime()
              : Date.now();
            
            // Add all messages from the recursive function
            // We only save assistant messages (tool messages are intermediate and don't need to be persisted)
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const newMessages = finalMessages
              .filter((msg: any) => {
                // Only include assistant messages (skip system, tool, and user messages - we already have user messages)
                return msg.role === 'assistant';
              })
              // Map to our Message interface, ensuring timestamps and content are set
              .map((msg: any, index: number) => ({
                role: 'assistant' as const,
                // For assistant messages with tool_calls but no content, use a placeholder
                // The final assistant message should have content
                content: msg.content || (msg.tool_calls ? 'Executing operations...' : ''),
                timestamp: msg.timestamp ? new Date(msg.timestamp) : new Date(lastTimestamp + (index + 1) * 1000), // Space them 1 second apart
              })) as Message[];
            
            return [...withoutTemp, ...newMessages];
          });

          // Trigger save after operations
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          if (typeof window !== 'undefined' && (window as any).saveWorkbookData) {
            setTimeout(async () => {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              await (window as any).saveWorkbookData();
            }, 500);
          }

          // OLD CODE REMOVED - All tool call handling now done by recursive handleToolCallsRecursive function above
        } catch (error) {
          console.error('Error in tool call execution:', error);
          const errorMessage: Message = {
            role: 'assistant',
            content: `Sorry, I encountered an error while executing operations: ${error instanceof Error ? error.message : 'Unknown error'}`,
            timestamp: new Date(),
          };
          setMessages((prev) => {
            // Remove the temporary "Executing operations..." message and add error
            const withoutTemp = prev.filter((msg, idx) => !(idx === prev.length - 1 && msg.content === 'Executing operations...'));
            return [...withoutTemp, errorMessage];
          });
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
      <div className="border-b border-[#E0E0E0] bg-white relative">
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

        {/* Action buttons */}
        <div className="absolute top-0 right-0 px-3 py-2.5">
          <button
            onClick={() => setShowInfo(!showInfo)}
            className="text-[#666666] hover:text-[#0066CC] transition-colors"
            title="Show capabilities"
          >
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
          </button>
        </div>
      </div>

      {/* Info Modal */}
      {showInfo && (
        <div className="absolute inset-0 bg-black bg-opacity-30 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-lg shadow-xl max-w-4xl w-full max-h-[85vh] overflow-y-auto">
            <div className="sticky top-0 bg-white border-b border-gray-200 px-6 py-4 flex items-center justify-between">
              <h2 className="text-2xl font-bold text-gray-900">SheetGrid Capabilities</h2>
              <button
                onClick={() => setShowInfo(false)}
                className="text-gray-400 hover:text-gray-600 transition-colors"
              >
                <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                </svg>
              </button>
            </div>

            <div className="p-6">
              <div className="mb-6">
                <h3 className="text-lg font-semibold text-gray-900 mb-3">📊 Data Operations</h3>
                <ul className="space-y-1 text-sm text-gray-700">
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">set_range_data</code> - Set values in cell ranges</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">get_range_data</code> - Read cell values and data</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">search_cells</code> - Search for specific content in cells</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">auto_fill</code> - Auto-fill data patterns</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">format_brush</code> - Copy and apply cell formatting</li>
                </ul>
              </div>

              <div className="mb-6">
                <h3 className="text-lg font-semibold text-gray-900 mb-3">📋 Sheet Management</h3>
                <ul className="space-y-1 text-sm text-gray-700">
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">create_sheet</code> - Create new worksheets</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">delete_sheet</code> - Remove worksheets</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">rename_sheet</code> - Rename existing sheets</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">activate_sheet</code> - Switch active worksheet</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">move_sheet</code> - Reorder sheet positions</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">set_sheet_display_status</code> - Show/hide sheets</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">get_sheets</code> - List all sheets</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">get_active_unit_id</code> - Get current workbook ID</li>
                </ul>
              </div>

              <div className="mb-6">
                <h3 className="text-lg font-semibold text-gray-900 mb-3">🏗️ Structure Operations</h3>
                <ul className="space-y-1 text-sm text-gray-700">
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">insert_rows</code> / <code className="bg-gray-100 px-1.5 py-0.5 rounded">insert_columns</code> - Add rows and columns</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">delete_rows</code> / <code className="bg-gray-100 px-1.5 py-0.5 rounded">delete_columns</code> - Remove rows and columns</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">set_cell_dimensions</code> - Adjust row heights and column widths</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">set_merge</code> - Merge cell ranges</li>
                </ul>
              </div>

              <div className="mb-6">
                <h3 className="text-lg font-semibold text-gray-900 mb-3">🎨 Formatting & Styling</h3>
                <ul className="space-y-1 text-sm text-gray-700">
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">set_range_style</code> - Apply cell styling</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">add_conditional_formatting_rule</code> - Add conditional formatting</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">set_conditional_formatting_rule</code> - Update conditional formatting</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">delete_conditional_formatting_rule</code> - Remove conditional formatting</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">get_conditional_formatting_rules</code> - List formatting rules</li>
                </ul>
              </div>

              <div className="mb-6">
                <h3 className="text-lg font-semibold text-gray-900 mb-3">✅ Data Validation</h3>
                <ul className="space-y-1 text-sm text-gray-700">
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">add_data_validation_rule</code> - Add validation rules</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">set_data_validation_rule</code> - Update validation rules</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">delete_data_validation_rule</code> - Remove validation rules</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">get_data_validation_rules</code> - List validation rules</li>
                </ul>
              </div>

              <div>
                <h3 className="text-lg font-semibold text-gray-900 mb-3">🔍 Utility Functions</h3>
                <ul className="space-y-1 text-sm text-gray-700">
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">get_activity_status</code> - Get workbook status and info</li>
                  <li><code className="bg-gray-100 px-1.5 py-0.5 rounded">scroll_and_screenshot</code> - Navigate and capture screenshots</li>
                </ul>
              </div>

              <div className="mt-8 p-4 bg-blue-50 border border-blue-200 rounded-lg">
                <p className="text-sm text-blue-900">
                  <strong>💡 Tip:</strong> Just describe what you want to do in natural language. The AI will automatically use the right tools to complete your request!
                </p>
              </div>
            </div>
          </div>
        </div>
      )}

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
                      {cleanMessageContent(message.content) ? (
                        <p className="text-sm leading-relaxed whitespace-pre-wrap break-words">
                          {cleanMessageContent(message.content)}
                        </p>
                      ) : message.role === 'assistant' ? (
                        <p className="text-sm text-[#999999] italic">Executing operations...</p>
                      ) : null}
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

              {/* Download document button */}
              <button
                type="button"
                onClick={async (e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  console.log('Download button clicked');

                  if (isDownloading) return; // Prevent multiple clicks

                  setIsDownloading(true);
                  try {
                    // Wait for the export function to be available (poll with timeout)
                    let exportFn = (window as any).exportWorkbookToXLSX;
                    const maxAttempts = 50; // Try for up to 5 seconds
                    let attempts = 0;

                    console.log('Checking for export function, attempt:', attempts, 'found:', !!exportFn);
                    while (!exportFn && attempts < maxAttempts) {
                      await new Promise(resolve => setTimeout(resolve, 100));
                      exportFn = (window as any).exportWorkbookToXLSX;
                      attempts++;
                      if (exportFn) {
                        console.log('Export function found after', attempts, 'attempts');
                      }
                    }

                    // Fallback: try to use univerAPI directly if available
                    if (!exportFn) {
                      console.log('Export function not found, trying fallback with univerAPI');
                      const univerAPI = (window as any).univerAPI;
                      if (univerAPI) {
                        try {
                          const workbook = univerAPI.getActiveWorkbook();
                          if (workbook) {
                            console.log('Got workbook from univerAPI, saving snapshot...');
                            const workbookSnapshot = workbook.save();
                            if (workbookSnapshot) {
                              // Import the export function dynamically
                              const { exportWorkbookToXLSX } = await import('../src/utils/xlsxConverter');
                              const filename = `workbook-${new Date().toISOString().split('T')[0]}.xlsx`;
                              console.log('Exporting workbook as', filename);
                              await exportWorkbookToXLSX(workbookSnapshot, filename);
                              console.log('Export completed successfully');
                              return;
                            }
                          }
                        } catch (fallbackError) {
                          console.error('Fallback export error:', fallbackError);
                          alert(`Failed to export workbook: ${fallbackError instanceof Error ? fallbackError.message : 'Unknown error'}`);
                          return;
                        }
                      }

                      alert('Export functionality is not available yet. Please wait for the spreadsheet to load.');
                      return;
                    }

                    console.log('Using export function, exporting...');
                    const filename = `workbook-${new Date().toISOString().split('T')[0]}.xlsx`;
                    await exportFn(filename);
                    console.log('Export completed successfully');
                  } catch (error) {
                    console.error('Error exporting workbook:', error);
                    alert(`Failed to export workbook: ${error instanceof Error ? error.message : 'Unknown error'}`);
                  } finally {
                    setIsDownloading(false);
                  }
                }}
                disabled={isDownloading}
                className="p-1.5 text-[#666666] hover:text-[#333333] hover:bg-[#F0F0F0] rounded transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                title={isDownloading ? 'Downloading...' : 'Download as XLSX'}
              >
                {isDownloading ? (
                  <svg className="w-4 h-4 animate-spin" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                  </svg>
                ) : (
                  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M9 19l3 3m0 0l3-3m-3 3V10" />
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
