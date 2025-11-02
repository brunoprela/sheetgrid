'use client';

import { useState, useRef, useEffect } from 'react';

interface ChatPanelProps {
  workbookData: { [sheetName: string]: any[][] };
  updateCell: (row: number, col: number, value: string, sheetName?: string) => void;
  setWorkbookData: React.Dispatch<React.SetStateAction<{ [sheetName: string]: any[][] }>>;
  getColumnLetter: (col: number) => string;
  activeSheetName: string;
}

interface Message {
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
}

interface ToolCall {
  name: string;
  arguments: any;
}

export default function ChatPanel({ workbookData, updateCell, setWorkbookData, getColumnLetter, activeSheetName }: ChatPanelProps) {
  const [messages, setMessages] = useState<Message[]>([
    {
      role: 'system',
      content: 'You are a helpful assistant that can edit Excel spreadsheets. Use the available tools to read, update, and analyze spreadsheet data.',
      timestamp: new Date(),
    },
  ]);
  const [input, setInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  // Tool definitions for Ollama
  const getTools = () => [
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
  ];

  const executeTool = async (toolCall: ToolCall): Promise<string> => {
    const { name, arguments: args } = toolCall;

    try {
      switch (name) {
        case 'get_cell_value': {
          const { row, col } = args;
          const sheetData = workbookData[activeSheetName] || [];
          const value = sheetData[row]?.[col] || '';
          return JSON.stringify({ value, cell: `${getColumnLetter(col)}${row + 1}` });
        }

        case 'set_cell_value': {
          const { row, col, value, sheetName } = args;
          const targetSheet = sheetName || activeSheetName;
          const newData = { ...workbookData };
          if (!newData[targetSheet]) newData[targetSheet] = [];
          if (!newData[targetSheet][row]) newData[targetSheet][row] = [];
          newData[targetSheet][row][col] = value;
          setWorkbookData(newData);
          return JSON.stringify({ success: true, cell: `${getColumnLetter(col)}${row + 1}` });
        }

        case 'get_range': {
          const { startRow, endRow, startCol, endCol, sheetName } = args;
          const sheetData = workbookData[sheetName || activeSheetName] || [];
          const result: any[][] = [];
          for (let row = startRow; row <= endRow; row++) {
            const rowData: any[] = [];
            for (let col = startCol; col <= endCol; col++) {
              rowData.push(sheetData[row]?.[col] || '');
            }
            result.push(rowData);
          }
          return JSON.stringify({ range: result });
        }

        case 'set_range': {
          const { startRow, values, sheetName } = args;
          const targetSheet = sheetName || activeSheetName;
          const newData = { ...workbookData };
          if (!newData[targetSheet]) newData[targetSheet] = [];
          values.forEach((rowValues: any[], offset: number) => {
            const row = startRow + offset;
            if (!newData[targetSheet][row]) newData[targetSheet][row] = [];
            rowValues.forEach((value, col) => {
              newData[targetSheet][row][col] = value;
            });
          });
          setWorkbookData(newData);
          return JSON.stringify({ success: true });
        }

        default:
          return JSON.stringify({ error: `Unknown tool: ${name}` });
      }
    } catch (error) {
      return JSON.stringify({ error: `Error executing tool: ${error}` });
    }
  };

  const sendMessage = async () => {
    if (!input.trim() || isLoading) return;

    const userMessage: Message = {
      role: 'user',
      content: input,
      timestamp: new Date(),
    };

    const updatedMessages = [...messages, userMessage];
    setMessages(updatedMessages);
    setInput('');
    setIsLoading(true);

    try {
      // Call Ollama API with tool support
      const response = await fetch('http://localhost:11434/api/chat', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          model: 'llama3.2',
          messages: updatedMessages.filter(m => m.role !== 'system').map((m) => ({ role: m.role, content: m.content })),
          tools: getTools(),
          tool_choice: 'auto',
          stream: false,
        }),
      });

      if (!response.ok) {
        throw new Error('Failed to get response from Ollama');
      }

      const data = await response.json();

      // Handle tool calls
      if (data.message?.tool_calls && data.message.tool_calls.length > 0) {
        const assistantMessage: Message = {
          role: 'assistant',
          content: 'Executing operations...',
          timestamp: new Date(),
        };
        setMessages((prev) => [...prev, assistantMessage]);

        // Execute all tool calls
        const toolResults: Array<{ name: string; result: string }> = [];
        for (const toolCall of data.message.tool_calls) {
          const result = await executeTool({
            name: toolCall.function.name,
            arguments: JSON.parse(toolCall.function.arguments),
          });
          toolResults.push({ name: toolCall.function.name, result });
        }

        // Prepare tool responses for Ollama
        const toolResponses = toolResults.map((tr, idx) => ({
          role: 'tool' as const,
          name: tr.name,
          content: tr.result,
        }));

        // Send the results back to Ollama for a final response
        const followUpResponse = await fetch('http://localhost:11434/api/chat', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            model: 'llama3.2',
            messages: [
              ...updatedMessages.filter(m => m.role !== 'system').map((m) => ({ role: m.role, content: m.content })),
              { role: 'assistant', content: data.message.content },
              ...toolResponses,
            ],
            stream: false,
          }),
        });

        const followUpData = await followUpResponse.json();
        const finalMessage: Message = {
          role: 'assistant',
          content: followUpData.message?.content || 'Operations completed successfully.',
          timestamp: new Date(),
        };
        setMessages((prev) => [...prev, finalMessage]);
      } else {
        // Direct response without tools
        const assistantMessage: Message = {
          role: 'assistant',
          content: data.message?.content || 'I received your message.',
          timestamp: new Date(),
        };
        setMessages((prev) => [...prev, assistantMessage]);
      }
    } catch (error) {
      console.error('Chat error:', error);
      const errorMessage: Message = {
        role: 'assistant',
        content: 'Sorry, I encountered an error. Please make sure Ollama is running on localhost:11434 with the llama3.2 model installed.',
        timestamp: new Date(),
      };
      setMessages((prev) => [...prev, errorMessage]);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="flex flex-col h-full bg-white">
      {/* Messages */}
      <div className="flex-1 overflow-y-auto p-6 space-y-6">
        {messages
          .filter((m) => m.role !== 'system')
          .map((message, idx) => (
            <div
              key={idx}
              className={`flex gap-4 ${message.role === 'user' ? 'flex-row-reverse' : ''}`}
            >
              {/* Avatar */}
              <div className={`shrink-0 w-8 h-8 rounded-full flex items-center justify-center text-sm font-medium ${
                message.role === 'user' 
                  ? 'bg-blue-600 text-white' 
                  : 'bg-gray-200 text-gray-600'
              }`}>
                {message.role === 'user' ? 'U' : 'AI'}
              </div>

              {/* Message content */}
              <div className="flex-1">
                <p className="text-sm text-gray-900 leading-relaxed whitespace-pre-wrap">
                  {message.content}
                </p>
              </div>
            </div>
          ))}
        {isLoading && (
          <div className="flex gap-4">
            <div className="shrink-0 w-8 h-8 rounded-full flex items-center justify-center text-sm font-medium bg-gray-200 text-gray-600">
              AI
            </div>
            <div className="flex items-center gap-2">
              <div className="w-2 h-2 bg-gray-400 rounded-full animate-bounce" style={{ animationDelay: '0s' }} />
              <div className="w-2 h-2 bg-gray-400 rounded-full animate-bounce" style={{ animationDelay: '0.2s' }} />
              <div className="w-2 h-2 bg-gray-400 rounded-full animate-bounce" style={{ animationDelay: '0.4s' }} />
            </div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* Input Area */}
      <div className="border-t border-gray-200 p-4 bg-white">
        <div className="relative">
          <textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                sendMessage();
              }
            }}
            placeholder="Ask me to analyze or edit your spreadsheet..."
            className="w-full px-4 py-3 border border-gray-300 rounded-lg resize-none focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm text-gray-900 placeholder-gray-500"
            rows={4}
            disabled={isLoading}
          />
          <button
            onClick={sendMessage}
            disabled={isLoading || !input.trim()}
            className="absolute bottom-3 right-3 p-2 bg-blue-600 hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed text-white rounded transition-colors"
          >
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8" />
            </svg>
          </button>
        </div>
      </div>
    </div>
  );
}
