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

Example of good response: "I've created a Date column with all dates from November 1 to November 30, 2025."
Example of bad response: "I set the first 5 dates, you can use auto_fill to extend the rest..."`,
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

  // Tool definitions for OpenRouter
  const getTools = () => {
    const tools = [
      {
        type: 'function',
        function: {
          name: 'get_cell_value',
          description: 'Get the value of a specific cell in the active sheet. Row and column are 0-indexed.',
          parameters: {
            type: 'object',
            properties: {
              row: { type: 'number', description: 'Row index (0-based)' },
              col: { type: 'number', description: 'Column index (0-based)' },
            },
            required: ['row', 'col'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'set_cell_value',
          description: 'Set the value of a specific cell in a sheet. Row and column are 0-indexed.',
          parameters: {
            type: 'object',
            properties: {
              row: { type: 'number', description: 'Row index (0-based)' },
              col: { type: 'number', description: 'Column index (0-based)' },
              value: { type: 'string', description: 'The value to set' },
              sheetName: { type: 'string', description: 'Optional sheet name. Defaults to active sheet.' },
            },
            required: ['row', 'col', 'value'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'get_range',
          description: 'Get a range of cells from a sheet',
          parameters: {
            type: 'object',
            properties: {
              startRow: { type: 'number', description: 'Start row index (0-based)' },
              endRow: { type: 'number', description: 'End row index (0-based)' },
              startCol: { type: 'number', description: 'Start column index (0-based)' },
              endCol: { type: 'number', description: 'End column index (0-based)' },
              sheetName: { type: 'string', description: 'Optional sheet name. Defaults to active sheet.' },
            },
            required: ['startRow', 'endRow', 'startCol', 'endCol'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'set_range',
          description: 'Set a range of cells with values',
          parameters: {
            type: 'object',
            properties: {
              startRow: { type: 'number', description: 'Start row index (0-based)' },
              values: {
                type: 'array',
                items: {
                  type: 'array',
                  items: { type: 'string' },
                },
                description: '2D array of values to set',
              },
              sheetName: { type: 'string', description: 'Optional sheet name. Defaults to active sheet.' },
            },
            required: ['startRow', 'values'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'set_column_header',
          description: 'Set the header value for a specific column. Use this to add or rename column headers. Column is 0-indexed (A=0, B=1, etc).',
          parameters: {
            type: 'object',
            properties: {
              col: { type: 'number', description: 'Column index (0-based, where A=0, B=1, etc)' },
              value: { type: 'string', description: 'The header value to set' },
              sheetName: { type: 'string', description: 'Optional sheet name. Defaults to active sheet.' },
            },
            required: ['col', 'value'],
          },
        },
      },
      // MCP Tools - Data Operations
      {
        type: 'function',
        function: {
          name: 'set_range_data',
          description: 'Set values in a cell range. IMPORTANT: Use this to fill ENTIRE ranges completely. If the user requests 30 dates, provide all 30 values in the values array. Never provide partial data - always complete the full requested range.',
          parameters: {
            type: 'object',
            properties: {
              startRow: { type: 'number' },
              startCol: { type: 'number' },
              endRow: { type: 'number' },
              endCol: { type: 'number' },
              values: { type: 'array', items: { type: 'array', items: { type: 'string' } } },
              unitId: { type: 'string', description: 'Workbook ID. Optional, defaults to active workbook.' },
              subUnitId: { type: 'string', description: 'Sheet ID. Optional, defaults to active sheet.' },
            },
            required: ['startRow', 'startCol', 'endRow', 'endCol', 'values'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'get_range_data',
          description: 'Read cell values and data',
          parameters: {
            type: 'object',
            properties: {
              startRow: { type: 'number' },
              startCol: { type: 'number' },
              endRow: { type: 'number' },
              endCol: { type: 'number' },
              unitId: { type: 'string', description: 'Workbook ID. Optional, defaults to active workbook.' },
              subUnitId: { type: 'string', description: 'Sheet ID. Optional, defaults to active sheet.' },
            },
            required: ['startRow', 'startCol', 'endRow', 'endCol'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'search_cells',
          description: 'Search for specific content in cells',
          parameters: {
            type: 'object',
            properties: {
              text: { type: 'string' },
              caseSensitive: { type: 'boolean', description: 'Optional' },
              unitId: { type: 'string', description: 'Workbook ID. Optional, defaults to active workbook.' },
            },
            required: ['text'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'auto_fill',
          description: 'Auto-fill data patterns',
          parameters: {
            type: 'object',
            properties: {
              sourceRange: { type: 'object', description: 'Source range object' },
              targetRange: { type: 'object', description: 'Target range object' },
              unitId: { type: 'string', description: 'Workbook ID. Optional, defaults to active workbook.' },
              subUnitId: { type: 'string', description: 'Sheet ID. Optional, defaults to active sheet.' },
            },
            required: ['sourceRange', 'targetRange'],
          },
        },
      },
      // MCP Tools - Sheet Management
      {
        type: 'function',
        function: {
          name: 'create_sheet',
          description: 'Create new worksheets',
          parameters: {
            type: 'object',
            properties: {
              name: { type: 'string' },
              unitId: { type: 'string', description: 'Workbook ID. Optional, defaults to active workbook.' },
            },
            required: ['name'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'delete_sheet',
          description: 'Remove worksheets',
          parameters: {
            type: 'object',
            properties: {
              subUnitId: { type: 'string', description: 'Sheet ID or name' },
              unitId: { type: 'string', description: 'Workbook ID. Optional, defaults to active workbook.' },
            },
            required: ['subUnitId'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'rename_sheet',
          description: 'Rename existing sheets',
          parameters: {
            type: 'object',
            properties: {
              subUnitId: { type: 'string', description: 'Sheet ID' },
              name: { type: 'string' },
              unitId: { type: 'string', description: 'Workbook ID. Optional, defaults to active workbook.' },
            },
            required: ['subUnitId', 'name'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'activate_sheet',
          description: 'Switch active worksheet',
          parameters: {
            type: 'object',
            properties: {
              subUnitId: { type: 'string', description: 'Sheet ID or name' },
              unitId: { type: 'string', description: 'Workbook ID. Optional, defaults to active workbook.' },
            },
            required: ['subUnitId'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'get_sheets',
          description: 'List all sheets',
          parameters: {
            type: 'object',
            properties: {
              unitId: { type: 'string', description: 'Workbook ID. Optional, defaults to active workbook.' },
            },
            required: [],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'get_active_unit_id',
          description: 'Get current workbook ID',
          parameters: {
            type: 'object',
            properties: {},
            required: [],
          },
        },
      },
      // MCP Tools - Structure Operations
      {
        type: 'function',
        function: {
          name: 'insert_rows',
          description: 'Add rows',
          parameters: {
            type: 'object',
            properties: {
              startRow: { type: 'number' },
              count: { type: 'number' },
              unitId: { type: 'string', description: 'Workbook ID. Optional, defaults to active workbook.' },
              subUnitId: { type: 'string', description: 'Sheet ID. Optional, defaults to active sheet.' },
            },
            required: ['startRow', 'count'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'insert_columns',
          description: 'Add columns',
          parameters: {
            type: 'object',
            properties: {
              startCol: { type: 'number' },
              count: { type: 'number' },
              unitId: { type: 'string', description: 'Workbook ID. Optional, defaults to active workbook.' },
              subUnitId: { type: 'string', description: 'Sheet ID. Optional, defaults to active sheet.' },
            },
            required: ['startCol', 'count'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'delete_rows',
          description: 'Remove rows',
          parameters: {
            type: 'object',
            properties: {
              startRow: { type: 'number' },
              count: { type: 'number' },
              unitId: { type: 'string', description: 'Workbook ID. Optional, defaults to active workbook.' },
              subUnitId: { type: 'string', description: 'Sheet ID. Optional, defaults to active sheet.' },
            },
            required: ['startRow', 'count'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'delete_columns',
          description: 'Remove columns',
          parameters: {
            type: 'object',
            properties: {
              startCol: { type: 'number' },
              count: { type: 'number' },
              unitId: { type: 'string', description: 'Workbook ID. Optional, defaults to active workbook.' },
              subUnitId: { type: 'string', description: 'Sheet ID. Optional, defaults to active sheet.' },
            },
            required: ['startCol', 'count'],
          },
        },
      },
      {
        type: 'function',
        function: {
          name: 'set_merge',
          description: 'Merge cell ranges',
          parameters: {
            type: 'object',
            properties: {
              startRow: { type: 'number' },
              startCol: { type: 'number' },
              endRow: { type: 'number' },
              endCol: { type: 'number' },
              unitId: { type: 'string', description: 'Workbook ID. Optional, defaults to active workbook.' },
              subUnitId: { type: 'string', description: 'Sheet ID. Optional, defaults to active sheet.' },
            },
            required: ['startRow', 'startCol', 'endRow', 'endCol'],
          },
        },
      },
    ];

    // Merge MCP tools from server (they take priority over local tools with same name)
    const toolMap = new Map(tools.map(t => [t.function.name, t]));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    mcpTools.forEach((mcpTool: any) => {
      toolMap.set(mcpTool.function.name, mcpTool);
    });

    const mergedTools = Array.from(toolMap.values());
    console.log(`getTools: returning ${mergedTools.length} tools (${mcpTools.length} from MCP, ${tools.length} local)`);
    return mergedTools;
  };

  const executeTool = async (toolCall: ToolCall): Promise<string> => {
    const { name, arguments: args } = toolCall;

    // Check if this is an MCP tool (from server)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const isMcpTool = mcpTools.some((tool: any) => tool.function.name === name);

    // For set_range_data, check if args contain code instead of actual values
    // If so, use local implementation to convert code to values
    let shouldUseLocal = false;
    if (name === 'set_range_data') {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const itemsStr = (args as any).items;
      if (typeof itemsStr === 'string') {
        // Check if it contains code (JavaScript or Python)
        if (itemsStr.includes('Math.random') || itemsStr.includes('Math.round') || itemsStr.includes('Math.floor') ||
          itemsStr.includes('for i in range') || itemsStr.includes('random()')) {
          console.log('set_range_data: Detected code in arguments, using local implementation to convert');
          shouldUseLocal = true;
        }
      }
    }

    console.log(`Tool "${name}": ${isMcpTool && !shouldUseLocal ? 'routing to MCP server' : 'using local implementation'}`, 'args:', JSON.stringify(args));

    if (isMcpTool && !shouldUseLocal) {
      try {
        const apiKey = import.meta.env.VITE_UNIVER_MCP_API_KEY || '';
        if (!apiKey) {
          return JSON.stringify({ error: 'MCP API key not configured' });
        }

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sessionId = (window as any).univerSessionId || 'default';

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
    }

    // Local tool implementations (fallback)
    try {
      // Access Univer API from window
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const univerAPI = (window as any).univerAPI;
      if (!univerAPI) {
        console.error('Tool execution failed: Univer API not available');
        return JSON.stringify({ error: 'Univer API not available yet. Please wait for the spreadsheet to load.' });
      }

      const workbook = univerAPI.getActiveWorkbook();
      if (!workbook) {
        console.error('Tool execution failed: No active workbook');
        return JSON.stringify({ error: 'No active workbook found. Please ensure a spreadsheet is loaded.' });
      }

      const activeSheet = workbook.getActiveSheet();
      if (!activeSheet) {
        console.error('Tool execution failed: No active sheet');
        return JSON.stringify({ error: 'No active sheet found. Please ensure a sheet is active.' });
      }

      // Log the actual sheet name for debugging
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sheetId = (activeSheet as any).getSheetId?.();
        const sheets = workbook.getSheets();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const foundSheet = sheets.find((s: any) => s.getSheetId?.() === sheetId);
        if (foundSheet) {
          const sheetName = typeof foundSheet.getName === 'function' ? foundSheet.getName() : 'Sheet1';
          console.log(`Tool executing on sheet: "${sheetName}" (ID: ${sheetId})`);
        }
      } catch (e) {
        console.warn('Could not log sheet name:', e);
      }

      // Helper to get column letter
      const getColLetter = (col: number): string => {
        let result = '';
        let num = col;
        while (num >= 0) {
          result = String.fromCharCode(65 + (num % 26)) + result;
          num = Math.floor(num / 26) - 1;
        }
        return result;
      };

      switch (name) {
        case 'get_cell_value': {
          const row = args.row as number;
          const col = args.col as number;
          const worksheet = workbook.getActiveSheet();
          const range = worksheet.getRange(row, col);
          const value = range.getValue();
          return JSON.stringify({
            value: value,
            cell: `${getColLetter(col)}${row + 1}`
          });
        }

        case 'set_cell_value': {
          const row = args.row as number;
          const col = args.col as number;
          const value = args.value as string;
          const worksheet = workbook.getActiveSheet();
          const range = worksheet.getRange(row, col);
          range.setValue(value);
          return JSON.stringify({
            success: true,
            cell: `${getColLetter(col)}${row + 1}`
          });
        }

        case 'get_range': {
          const startRow = args.startRow as number;
          const endRow = args.endRow as number;
          const startCol = args.startCol as number;
          const endCol = args.endCol as number;
          const worksheet = workbook.getActiveSheet();
          const range = worksheet.getRange(startRow, startCol, endRow, endCol);
          const values = range.getValues();
          return JSON.stringify({ range: values });
        }

        case 'set_range': {
          try {
            const startRow = args.startRow as number;
            const values = args.values as string[][];
            console.log('set_range called with:', { startRow, valuesLength: values.length, firstValue: values[0] });
            const worksheet = workbook.getActiveSheet();
            // If startRow is 0, assume row 0 is a header, so start from row 1 instead
            const actualStartRow = startRow === 0 ? 1 : startRow;
            console.log('Using actualStartRow:', actualStartRow, 'EndRow:', actualStartRow + values.length - 1);

            // Set values cell by cell instead of using setValues
            for (let i = 0; i < values.length; i++) {
              const row = actualStartRow + i;
              for (let j = 0; j < values[i].length; j++) {
                const col = j;
                const value = values[i][j];
                const cellRange = worksheet.getRange(row, col);
                cellRange.setValue(value);
              }
            }

            console.log('Values set successfully (cell by cell)');
            return JSON.stringify({ success: true });
          } catch (err) {
            console.error('set_range error:', err);
            return JSON.stringify({ error: `set_range failed: ${err}` });
          }
        }

        case 'set_column_header': {
          const col = args.col as number;
          const value = args.value as string;
          const worksheet = workbook.getActiveSheet();
          // Set the header in row 0 (first row)
          const range = worksheet.getRange(0, col);
          range.setValue(value);
          return JSON.stringify({
            success: true,
            column: getColLetter(col)
          });
        }

        // MCP Tools - Data Operations
        case 'set_range_data': {
          try {
            // Handle both MCP server format (items) and local format (startRow, startCol, etc.)
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            let startRow: number, startCol: number, endRow: number, endCol: number, values: any;

            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            if ((args as any).items) {
              // MCP server format: { items: "[{ range: 'B1:B30', value: [...] }]" }
              console.log('set_range_data: Detected MCP format with items parameter');
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              let itemsData = (args as any).items;

              // Parse items if it's a string
              if (typeof itemsData === 'string') {
                try {
                  // The JSON string contains unquoted JavaScript code like:
                  // "value": [Math.floor(Math.random() * 1001), ...]
                  // We need to quote these code strings first, then replace with numbers

                  // Find all unquoted Math.random expressions and quote them
                  // Pattern matches: Math.floor(Math.random() * 1001) or Math.round(Math.random() * 1000)
                  const codePattern = /(Math\.(floor|round|random)\(Math\.random\(\)\s*\*\s*\d+\))/g;
                  let placeholderCounter = 0;
                  const placeholderMap = new Map<string, number>();

                  const cleanedString = itemsData.replace(codePattern, (match) => {
                    // Check if already processed
                    if (!placeholderMap.has(match)) {
                      placeholderMap.set(match, placeholderCounter++);
                    }
                    // Return quoted placeholder that we can identify
                    return `"__RANDOM_${placeholderMap.get(match)}__"`;
                  });

                  // Parse the cleaned JSON (now all values are properly quoted)
                  itemsData = JSON.parse(cleanedString);

                  // Now replace placeholders with actual random numbers recursively
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const replaceRandomNumbers = (obj: any): any => {
                    if (Array.isArray(obj)) {
                      return obj.map((item) => {
                        // Check for placeholder strings like "__RANDOM_0__"
                        if (typeof item === 'string' && item.startsWith('__RANDOM_') && item.endsWith('__')) {
                          return Math.floor(Math.random() * 1001);
                        } else if (typeof item === 'object' && item !== null) {
                          return replaceRandomNumbers(item);
                        }
                        // Check if item is a string containing code
                        if (typeof item === 'string' && (
                          item.includes('Math.floor') ||
                          item.includes('Math.round') ||
                          item.includes('Math.random')
                        )) {
                          return Math.floor(Math.random() * 1001);
                        }
                        return item;
                      });
                    } else if (typeof obj === 'object' && obj !== null) {
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      const result: any = {};
                      for (const key in obj) {
                        if (typeof obj[key] === 'string' && obj[key].startsWith('__RANDOM_') && obj[key].endsWith('__')) {
                          result[key] = Math.floor(Math.random() * 1001);
                        } else if (typeof obj[key] === 'string' && (
                          obj[key].includes('Math.floor') ||
                          obj[key].includes('Math.round') ||
                          obj[key].includes('Math.random')
                        )) {
                          result[key] = Math.floor(Math.random() * 1001);
                        } else if (Array.isArray(obj[key]) || (typeof obj[key] === 'object' && obj[key] !== null)) {
                          result[key] = replaceRandomNumbers(obj[key]);
                        } else {
                          result[key] = obj[key];
                        }
                      }
                      return result;
                    }
                    return obj;
                  };

                  itemsData = replaceRandomNumbers(itemsData);
                } catch (e) {
                  console.error('Failed to parse items:', e, 'Original string:', itemsData.substring(0, 200));
                  return JSON.stringify({ error: `Failed to parse items: ${e}` });
                }
              }

              // Handle array of items
              if (Array.isArray(itemsData) && itemsData.length > 0) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const item = itemsData[0] as any;

                // Parse range like "B1:B30"
                if (item.range) {
                  const rangeMatch = item.range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
                  if (!rangeMatch) {
                    return JSON.stringify({ error: `Invalid range format: ${item.range}. Expected format: A1:B30` });
                  }

                  // Convert column letters to numbers
                  const colToNum = (col: string): number => {
                    let num = 0;
                    for (let i = 0; i < col.length; i++) {
                      num = num * 26 + (col.charCodeAt(i) - 64);
                    }
                    return num - 1; // 0-indexed
                  };

                  startCol = colToNum(rangeMatch[1]);
                  startRow = parseInt(rangeMatch[2], 10) - 1; // 0-indexed
                  endCol = colToNum(rangeMatch[3]);
                  endRow = parseInt(rangeMatch[4], 10) - 1; // 0-indexed

                  // Get values array
                  values = item.value || item.values;

                  // Filter out any JavaScript/Python code strings and convert to numbers
                  if (Array.isArray(values)) {
                    values = values.map((v) => {
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      const val = v as any;
                      // If it's a string that looks like code, generate random number
                      if (typeof val === 'string' && (
                        val.includes('Math.floor') ||
                        val.includes('Math.random') ||
                        val.includes('Math.round') ||
                        val.includes('random()') ||
                        val.includes('for i in range')
                      )) {
                        // Generate actual random number instead (0-1000)
                        return Math.floor(Math.random() * 1001);
                      }
                      // Convert to number if possible
                      const num = Number(val);
                      return isNaN(num) ? val : num;
                    });
                  }

                  // If values is a string containing Python list comprehension, generate the array
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  if (typeof (item as any).value === 'string' &&
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    (item as any).value.includes('for i in range')) {
                    // Extract the count from "for i in range(30)"
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const rangeMatch = (item as any).value.match(/range\((\d+)\)/);
                    if (rangeMatch) {
                      const count = parseInt(rangeMatch[1], 10);
                      values = Array.from({ length: count }, () => Math.floor(Math.random() * 1001));
                      console.log(`Generated ${count} random numbers (0-1000) from Python list comprehension`);
                    }
                  }
                } else {
                  return JSON.stringify({ error: 'Items format missing range' });
                }
              } else {
                return JSON.stringify({ error: 'Items must be a non-empty array' });
              }
            } else {
              // Local format: { startRow, startCol, endRow, endCol, values }
              startRow = args.startRow as number;
              startCol = args.startCol as number;
              endRow = args.endRow as number;
              endCol = args.endCol as number;
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              values = args.values as any;
            }

            console.log('set_range_data called with:', { startRow, startCol, endRow, endCol, valuesLength: values?.length, firstValue: values?.[0] });
            const worksheet = workbook.getActiveSheet();
            // If startRow is 0, assume row 0 is a header, so start from row 1 instead
            const actualStartRow = startRow === 0 ? 1 : startRow;
            const actualEndRow = actualStartRow + values.length - 1;

            // Ensure values is a 2D array
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const values2D: string[][] = Array.isArray(values[0]) ? values : values.map((v: any) => [String(v)]);
            console.log('Using actualStartRow:', actualStartRow, 'actualEndRow:', actualEndRow, 'rows:', values2D.length);

            // Set values row by row to ensure they persist
            for (let i = 0; i < values2D.length; i++) {
              const row = actualStartRow + i;
              const rowValues = values2D[i];
              for (let j = 0; j < rowValues.length; j++) {
                const col = startCol + j;
                if (col <= endCol) {
                  const cellRange = worksheet.getRange(row, col);
                  cellRange.setValue(rowValues[j]);
                }
              }
            }

            console.log('Values set cell-by-cell, verifying...');
            // Verify by reading back a few cells
            const verifyRange = worksheet.getRange(actualStartRow, startCol, Math.min(actualStartRow + 2, actualEndRow), endCol);
            const readBack = verifyRange.getValues();
            console.log('Read back from spreadsheet:', readBack);
            console.log('Values set successfully');
            return JSON.stringify({ success: true });
          } catch (err) {
            console.error('set_range_data error:', err);
            return JSON.stringify({ error: `set_range_data failed: ${err}` });
          }
        }

        case 'get_range_data': {
          const startRow = args.startRow as number;
          const startCol = args.startCol as number;
          const endRow = args.endRow as number;
          const endCol = args.endCol as number;
          const worksheet = workbook.getActiveSheet();
          const range = worksheet.getRange(startRow, startCol, endRow, endCol);
          const values = range.getValues();
          return JSON.stringify({ data: values });
        }

        case 'search_cells': {
          const text = args.text as string;
          const caseSensitive = args.caseSensitive as boolean || false;
          const worksheet = workbook.getActiveSheet();
          const matches: Array<{ row: number, col: number, value: string }> = [];

          // Simple search implementation - could be optimized
          for (let row = 0; row < 1000; row++) {
            for (let col = 0; col < 100; col++) {
              try {
                const range = worksheet.getRange(row, col);
                const value = String(range.getValue() || '');
                const searchValue = caseSensitive ? value : value.toLowerCase();
                const searchText = caseSensitive ? text : text.toLowerCase();
                if (searchValue.includes(searchText)) {
                  matches.push({ row, col, value });
                }
              } catch {
                // Skip invalid ranges
              }
            }
          }
          return JSON.stringify({ matches });
        }

        // MCP Tools - Sheet Management
        case 'create_sheet': {
          try {
            const name = args.name as string;
            const worksheet = workbook.insertSheet(name);
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            return JSON.stringify({ success: true, sheetId: (worksheet as any).getSheetId?.() || 'unknown' });
          } catch (err) {
            return JSON.stringify({ error: `create_sheet failed: ${err}` });
          }
        }

        case 'get_sheets': {
          try {
            const sheets = workbook.getSheets();
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const sheetList = sheets.map((_: any, idx: number) => ({
              id: `sheet-${idx}`,
              name: `Sheet${idx + 1}`,
            }));
            return JSON.stringify({ sheets: sheetList });
          } catch (err) {
            return JSON.stringify({ error: `get_sheets failed: ${err}` });
          }
        }

        case 'get_active_unit_id': {
          return JSON.stringify({ unitId: workbook.getId() });
        }

        case 'delete_sheet': {
          const subUnitId = args.subUnitId as string;
          // Try to find and delete the sheet by ID or name
          const sheets = workbook.getSheets();
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const sheet = sheets.find((s: any) => s.getSheetId() === subUnitId || s.getName() === subUnitId);
          if (sheet) {
            workbook.removeSheet(sheet);
            return JSON.stringify({ success: true });
          }
          return JSON.stringify({ error: 'Sheet not found' });
        }

        case 'rename_sheet': {
          const subUnitId = args.subUnitId as string;
          const name = args.name as string;
          const sheets = workbook.getSheets();
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const sheet = sheets.find((s: any) => s.getSheetId() === subUnitId || s.getName() === subUnitId);
          if (sheet) {
            sheet.setName(name);
            return JSON.stringify({ success: true });
          }
          return JSON.stringify({ error: 'Sheet not found' });
        }

        case 'activate_sheet': {
          try {
            const subUnitId = args.subUnitId as string;
            const sheets = workbook.getSheets();
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const sheet = sheets.find((s: any) => s.getSheetId?.() === subUnitId || s.getName?.() === subUnitId);
            if (sheet) {
              workbook.setActiveSheet(sheet);
              return JSON.stringify({ success: true });
            }
            return JSON.stringify({ error: 'Sheet not found' });
          } catch (err) {
            return JSON.stringify({ error: `activate_sheet failed: ${err}` });
          }
        }

        // MCP Tools - Structure Operations
        case 'insert_rows': {
          const startRow = args.startRow as number;
          const count = args.count as number;
          const worksheet = workbook.getActiveSheet();
          worksheet.insertRows(startRow, count);
          return JSON.stringify({ success: true });
        }

        case 'delete_rows': {
          const startRow = args.startRow as number;
          const count = args.count as number;
          const worksheet = workbook.getActiveSheet();
          worksheet.removeRows(startRow, count);
          return JSON.stringify({ success: true });
        }

        case 'insert_columns': {
          try {
            const startCol = args.startCol as number;
            const count = args.count as number;
            const worksheet = workbook.getActiveSheet();
            worksheet.insertColumns(startCol, count);
            return JSON.stringify({ success: true });
          } catch (err) {
            return JSON.stringify({ error: `insert_columns failed: ${err}` });
          }
        }

        case 'delete_columns': {
          const startCol = args.startCol as number;
          const count = args.count as number;
          const worksheet = workbook.getActiveSheet();
          worksheet.removeColumns(startCol, count);
          return JSON.stringify({ success: true });
        }

        case 'set_merge': {
          const startRow = args.startRow as number;
          const startCol = args.startCol as number;
          const endRow = args.endRow as number;
          const endCol = args.endCol as number;
          const worksheet = workbook.getActiveSheet();
          const range = worksheet.getRange(startRow, startCol, endRow, endCol);
          range.merge();
          return JSON.stringify({ success: true });
        }

        default:
          return JSON.stringify({ error: `Unknown tool: ${name}` });
      }
    } catch (error) {
      return JSON.stringify({ error: `Error executing tool: ${error}` });
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
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });

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

        // Convert sheet to JSON array format
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: '',
          raw: false
        }) as string[][];

        // Ensure jsonData is a proper 2D array and pad rows to same length
        const maxCols = Math.max(...jsonData.map(row => row.length), 0);
        const paddedData = jsonData.map(row => {
          const padded = [...row];
          while (padded.length < maxCols) {
            padded.push('');
          }
          return padded;
        });

        if (sheetIndex === 0) {
          // Use the first existing sheet
          const firstSheet = activeWorkbook.getActiveSheet();
          if (firstSheet && paddedData.length > 0 && maxCols > 0) {
            try {
              // Use batch operation to set all values at once
              const endRow = paddedData.length - 1;
              const endCol = maxCols - 1;
              const range = firstSheet.getRange(0, 0, endRow, endCol);

              // Convert to 2D array of strings for setValues
              const values2D: string[][] = paddedData.map(row =>
                row.map(cell => cell !== undefined && cell !== null ? String(cell) : '')
              );

              // Set all values in one batch operation
              range.setValues(values2D);
            } catch (e) {
              console.warn('Batch setValues failed, falling back to cell-by-cell:', e);
              // Fallback to cell-by-cell if batch fails
              paddedData.forEach((row, rowIndex) => {
                row.forEach((cellValue, colIndex) => {
                  if (cellValue !== undefined && cellValue !== null && cellValue !== '') {
                    try {
                      firstSheet.getRange(rowIndex, colIndex, rowIndex, colIndex).setValue(String(cellValue));
                    } catch {
                      // Ignore individual cell errors
                    }
                  }
                });
              });
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
                // Use batch operation for new sheets too
                const endRow = paddedData.length - 1;
                const endCol = maxCols - 1;
                const range = newSheet.getRange(0, 0, endRow, endCol);

                // Convert to 2D array of strings
                const values2D: string[][] = paddedData.map(row =>
                  row.map(cell => cell !== undefined && cell !== null ? String(cell) : '')
                );

                // Set all values in one batch operation
                range.setValues(values2D);
              } catch (e) {
                console.warn('Batch setValues failed for new sheet, falling back:', e);
                // Fallback to cell-by-cell
                paddedData.forEach((row, rowIndex) => {
                  row.forEach((cellValue, colIndex) => {
                    if (cellValue !== undefined && cellValue !== null && cellValue !== '') {
                      try {
                        newSheet.getRange(rowIndex, colIndex, rowIndex, colIndex).setValue(String(cellValue));
                      } catch {
                        // Ignore individual cell errors
                      }
                    }
                  });
                });
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
      // Call OpenRouter API with tool support
      const response = await fetch('https://openrouter.ai/api/v1/chat/completions', {
        signal: abortController.signal,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${import.meta.env.VITE_OPENROUTER_API_KEY || ''}`,
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
        const errorData = await response.json();
        throw new Error(`Failed to get response from OpenRouter: ${errorData.error?.message || 'Unknown error'}`);
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
        const followUpResponse = await fetch('https://openrouter.ai/api/v1/chat/completions', {
          signal: abortController.signal,
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${import.meta.env.VITE_OPENROUTER_API_KEY || ''}`,
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
              'Authorization': `Bearer ${import.meta.env.VITE_OPENROUTER_API_KEY || ''}`,
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
        // Direct response without tools
        const assistantMessage: Message = {
          role: 'assistant',
          content: data.choices?.[0]?.message?.content || 'I received your message.',
          timestamp: new Date(),
        };
        setMessages((prev) => [...prev, assistantMessage]);
      }
    } catch (error) {
      // Don't show error message if request was aborted by user
      if (error instanceof Error && error.name === 'AbortError') {
        console.log('Request was cancelled by user');
        return;
      }
      console.error('Chat error:', error);
      const errorMessage: Message = {
        role: 'assistant',
        content: `Sorry, I encountered an error. ${error instanceof Error ? error.message : 'Please check your OpenRouter API key and connection.'}`,
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
