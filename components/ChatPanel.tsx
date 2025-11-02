import { useState, useRef, useEffect } from 'react';

// eslint-disable-next-line @typescript-eslint/no-empty-object-type
interface ChatPanelProps { }

interface Message {
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
}

interface ToolCall {
  name: string;
  arguments: Record<string, unknown>;
}

export default function ChatPanel({ }: ChatPanelProps) {
  const [messages, setMessages] = useState<Message[]>([
    {
      role: 'system',
      content: 'You are a helpful assistant that can edit Excel spreadsheets using tools. When the user asks you to create a column with dates, you MUST: 1) use set_column_header to create the header, AND 2) use set_range_data to fill in all the date values. Never stop after just one tool call - you must complete ALL steps of the task using multiple tool calls.',
      timestamp: new Date(),
    },
  ]);
  const [input, setInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
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

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

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
          description: 'Set values in cell ranges',
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

    console.log(`Tool "${name}": ${isMcpTool ? 'routing to MCP server' : 'using local implementation'}`);

    if (isMcpTool) {
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
        return JSON.stringify({ error: 'Univer API not available yet' });
      }

      const workbook = univerAPI.getActiveWorkbook();
      if (!workbook) {
        return JSON.stringify({ error: 'No active workbook found' });
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
            const startRow = args.startRow as number;
            const startCol = args.startCol as number;
            const endRow = args.endRow as number;
            const endCol = args.endCol as number;
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const values = args.values as any;
            console.log('set_range_data called with:', { startRow, startCol, endRow, endCol, valuesLength: values.length, firstValue: values[0] });
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

  const sendMessage = async () => {
    if (!input.trim() || isLoading) return;

    // Create new AbortController for this request
    const abortController = new AbortController();
    abortControllerRef.current = abortController;

    const userMessage: Message = {
      role: 'user',
      content: input,
      timestamp: new Date(),
    };

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
          messages: updatedMessages.filter(m => m.role !== 'system').map((m) => ({ role: m.role, content: m.content })),
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
              ...updatedMessages.filter(m => m.role !== 'system').map((m) => ({ role: m.role, content: m.content })),
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
                ...updatedMessages.filter(m => m.role !== 'system').map((m) => ({ role: m.role, content: m.content })),
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
        } else {
          const finalMessage: Message = {
            role: 'assistant',
            content: followUpData.choices?.[0]?.message?.content || 'Operations completed successfully.',
            timestamp: new Date(),
          };
          setMessages((prev) => [...prev, finalMessage]);
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
                        {message.content}
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
              {/* Upload document button */}
              <button
                className="p-1.5 text-[#666666] hover:text-[#333333] hover:bg-[#F0F0F0] rounded transition-colors"
                title="Upload document"
              >
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                </svg>
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
