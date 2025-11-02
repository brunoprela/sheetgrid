import { useEffect, useRef } from 'react';
import { createUniver, LocaleType, UniverInstanceType, LogLevel, defaultTheme } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/presets/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/presets/preset-sheets-core/locales/en-US';
import { UniverSheetsAdvancedPreset } from '@univerjs/presets/preset-sheets-advanced';
import UniverPresetSheetsAdvancedEnUS from '@univerjs/presets/preset-sheets-advanced/locales/en-US';
import { UniverSheetsDrawingPreset } from '@univerjs/presets/preset-sheets-drawing';
import UniverPresetSheetsDrawingEnUS from '@univerjs/presets/preset-sheets-drawing/locales/en-US';
import { UniverSheetsFilterPreset } from '@univerjs/presets/preset-sheets-filter';
import UniverPresetSheetsFilterEnUS from '@univerjs/presets/preset-sheets-filter/locales/en-US';
import { UniverSheetsFindReplacePreset } from '@univerjs/presets/preset-sheets-find-replace';
import UniverPresetSheetsFindReplaceEnUS from '@univerjs/presets/preset-sheets-find-replace/locales/en-US';
import { UniverSheetsSortPreset } from '@univerjs/presets/preset-sheets-sort';
import UniverPresetSheetsSortEnUS from '@univerjs/presets/preset-sheets-sort/locales/en-US';
import { UniverSheetsHyperLinkPreset } from '@univerjs/presets/preset-sheets-hyper-link';
import UniverPresetSheetsHyperLinkEnUS from '@univerjs/presets/preset-sheets-hyper-link/locales/en-US';
import { UniverSheetsThreadCommentPreset } from '@univerjs/presets/preset-sheets-thread-comment';
import UniverPresetSheetsThreadCommentEnUS from '@univerjs/presets/preset-sheets-thread-comment/locales/en-US';
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/presets/preset-sheets-conditional-formatting';
import UniverPresetSheetsConditionalFormattingEnUS from '@univerjs/presets/preset-sheets-conditional-formatting/locales/en-US';
import { UniverSheetsDataValidationPreset } from '@univerjs/presets/preset-sheets-data-validation';
import UniverPresetSheetsDataValidationEnUS from '@univerjs/presets/preset-sheets-data-validation/locales/en-US';
import { UniverSheetsCrosshairHighlightPlugin } from '@univerjs/sheets-crosshair-highlight';
import UniverSheetsCrosshairHighlightEnUS from '@univerjs/sheets-crosshair-highlight/locale/en-US';
import { UniverSheetsZenEditorPlugin } from '@univerjs/sheets-zen-editor';
import UniverSheetsZenEditorEnUS from '@univerjs/sheets-zen-editor/locale/en-US';
import { UniverMCPPlugin } from '@univerjs-pro/mcp';
import { UniverMCPUIPlugin } from '@univerjs-pro/mcp-ui';
import univerMCPUIEnUS from '@univerjs-pro/mcp-ui/locale/en-US';
import { UniverSheetMCPPlugin } from '@univerjs-pro/sheets-mcp';
import '@univerjs/sheets/facade';
import '@univerjs-pro/mcp/facade';
import '@univerjs/presets/lib/styles/preset-sheets-core.css';
import '@univerjs/presets/lib/styles/preset-sheets-advanced.css';
import '@univerjs/presets/lib/styles/preset-sheets-filter.css';
import '@univerjs/presets/lib/styles/preset-sheets-find-replace.css';
import '@univerjs/presets/lib/styles/preset-sheets-sort.css';
import '@univerjs/presets/lib/styles/preset-sheets-hyper-link.css';
import '@univerjs/presets/lib/styles/preset-sheets-thread-comment.css';
import '@univerjs/presets/lib/styles/preset-sheets-conditional-formatting.css';
import '@univerjs/presets/lib/styles/preset-sheets-data-validation.css';
import '@univerjs/presets/lib/styles/preset-sheets-drawing.css';
import '@univerjs-pro/mcp-ui/lib/index.css';

// Define merge function to match start kit
const merge = <T extends Record<string, any>>(target: T, ...sources: any[]): T => {
  return Object.assign({}, target, ...sources);
};

interface SpreadsheetProps {}

export default function Spreadsheet({}: SpreadsheetProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerInstanceRef = useRef<{ univerAPI: any; univer: any } | null>(null);

  useEffect(() => {
    if (!containerRef.current || univerInstanceRef.current) return;

    // Generate a unique session ID for MCP server connection
    // Must be generated BEFORE creating Univer instance
    const sessionId = `session-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
    
    // Store session ID globally for MCP client connection
    if (typeof window !== 'undefined') {
      (window as any).univerSessionId = sessionId;
    }

    // Create Univer instance with preset and MCP plugins
    const universerEndpoint = window.location.host;
    const collaboration = undefined;
    const apiKey = import.meta.env.VITE_UNIVER_MCP_API_KEY || '';

    // Automatically store API key in localStorage for UniverMCPUIPlugin
    // UniverMCPUIPlugin stores API key in localStorage, so we pre-populate it
    // This allows the plugin to use the API key without manual user configuration
    if (apiKey && typeof window !== 'undefined') {
      try {
        // Try common localStorage keys that UniverMCPUIPlugin might use
        // Store in multiple possible locations to ensure compatibility
        const possibleKeys = [
          'univer_mcp_api_key',
          'univer-mcp-api-key',
          '@univerjs-pro/mcp-ui/api-key',
        ];
        
        possibleKeys.forEach(key => {
          localStorage.setItem(key, apiKey);
        });
        
        console.log('MCP API key automatically configured from environment variable');
        console.log('API key stored in localStorage, MCP should connect automatically');
      } catch (error) {
        console.warn('Failed to store API key in localStorage:', error);
      }
    } else if (!apiKey) {
      console.warn('No MCP API key found in environment. MCP will be disconnected.');
      console.warn('Set VITE_UNIVER_MCP_API_KEY in .env.local to enable MCP');
    }

    const { univerAPI, univer } = createUniver({
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: merge(
          {},
          UniverPresetSheetsCoreEnUS,
          UniverPresetSheetsAdvancedEnUS,
          UniverPresetSheetsDrawingEnUS,
          UniverPresetSheetsFilterEnUS,
          UniverPresetSheetsFindReplaceEnUS,
          UniverPresetSheetsSortEnUS,
          UniverPresetSheetsHyperLinkEnUS,
          UniverPresetSheetsThreadCommentEnUS,
          UniverPresetSheetsConditionalFormattingEnUS,
          UniverPresetSheetsDataValidationEnUS,
          UniverSheetsCrosshairHighlightEnUS,
          UniverSheetsZenEditorEnUS,
          univerMCPUIEnUS,
        ),
      },
      collaboration,
      logLevel: LogLevel.VERBOSE,
      theme: defaultTheme,
      presets: [
        UniverSheetsCorePreset({
          container: containerRef.current!,
          header: true,
        }),
        UniverSheetsDrawingPreset({
          collaboration,
        }),
        UniverSheetsAdvancedPreset({
          useWorker: false,
          universerEndpoint,
        }),
        UniverSheetsThreadCommentPreset({
          collaboration,
        }),
        UniverSheetsConditionalFormattingPreset(),
        UniverSheetsDataValidationPreset(),
        UniverSheetsFilterPreset(),
        UniverSheetsFindReplacePreset(),
        UniverSheetsSortPreset(),
        UniverSheetsHyperLinkPreset(),
      ],
      plugins: [
        UniverSheetsCrosshairHighlightPlugin,
        UniverSheetsZenEditorPlugin,
        // Register UniverMCPPlugin with configuration BEFORE UniverSheetMCPPlugin
        // UniverSheetMCPPlugin depends on it and should reuse this instance
        [UniverMCPPlugin, {
          sessionId: sessionId,
          ticketServerUrl: 'https://mcp.univer.ai/api/ticket',
          mcpServerUrl: 'wss://mcp.univer.ai/api/ws',
          enableAuth: !!apiKey,
        }],
        // Initialize UI plugin - this can manage API key storage
        [UniverMCPUIPlugin, {
          showDeveloperTools: true,
        }],
        // UniverSheetMCPPlugin registers 30+ spreadsheet tools
        // It depends on UniverMCPPlugin which is registered above
        UniverSheetMCPPlugin,
      ],
    });

    // Create a workbook
    univer.createUnit(UniverInstanceType.UNIVER_SHEET, {});

    // Store instance in ref to prevent double initialization
    univerInstanceRef.current = { univerAPI, univer };

    // Store API globally for MCP client connection
    if (typeof window !== 'undefined') {
      (window as any).univerAPI = univerAPI;
    }
    
    console.log('Univer MCP configured with sessionId:', sessionId);
    console.log('MCP API key available:', apiKey ? 'Yes (from env)' : 'No');
    
    // Try to automatically configure the MCP UI plugin with the API key
    if (apiKey && typeof window !== 'undefined') {
      // Wait for plugins to initialize, then try to set API key
      setTimeout(() => {
        // Try to find and interact with the MCP UI plugin
        // Check what localStorage keys exist after plugin init
        const localStorageKeys = Object.keys(localStorage).filter(key => 
          key.toLowerCase().includes('univer') || key.toLowerCase().includes('mcp')
        );
        console.log('MCP-related localStorage keys:', localStorageKeys);
        
        // Try to find the actual key the plugin uses
        localStorageKeys.forEach(key => {
          const value = localStorage.getItem(key);
          if (value === apiKey) {
            console.log(`Found API key stored at: ${key}`);
          }
        });
        
        // Try to trigger MCP connection by dispatching a storage event
        // This might trigger the plugin to re-read the API key
        window.dispatchEvent(new StorageEvent('storage', {
          key: '@univerjs-pro/mcp-ui/api-key',
          newValue: apiKey,
          storageArea: localStorage,
        }));
        
        // Also try setting it directly in case the plugin checks on load
        localStorage.setItem('@univerjs-pro/mcp-ui/api-key', apiKey);
        
        // Monitor network requests to see what MCP plugin is doing
        const originalFetch = window.fetch;
        window.fetch = function(...args) {
          const url = typeof args[0] === 'string' ? args[0] : args[0].url;
          if (url.includes('mcp.univer.ai') || url.includes('ticket.univer.ai')) {
            console.log('🔍 MCP Network Request:', {
              url,
              method: args[1]?.method || 'GET',
              headers: args[1]?.headers,
            });
          }
          return originalFetch.apply(this, args as any);
        };
        
        // Monitor WebSocket connections
        const originalWebSocket = window.WebSocket;
        window.WebSocket = function(...args: any[]) {
          const ws = new originalWebSocket(...args);
          const url = args[0];
          if (url.includes('mcp.univer.ai') || url.includes('ws')) {
            console.log('🔍 MCP WebSocket Connection:', url);
            ws.addEventListener('open', () => console.log('✅ MCP WebSocket opened'));
            ws.addEventListener('error', (e) => console.error('❌ MCP WebSocket error:', e));
            ws.addEventListener('close', () => console.log('🔌 MCP WebSocket closed'));
            ws.addEventListener('message', (e) => {
              try {
                const data = JSON.parse(e.data);
                console.log('📨 MCP WebSocket message:', data);
              } catch {
                console.log('📨 MCP WebSocket message (raw):', e.data.substring(0, 100));
              }
            });
          }
          return ws;
        } as any;
        
        console.log('Attempted to automatically configure MCP API key');
        console.log('Network monitoring enabled for MCP requests');
      }, 1000);
    }

    return () => {
      if (univerInstanceRef.current) {
        univerInstanceRef.current.univerAPI.dispose();
        univerInstanceRef.current = null;
      }
    };
  }, []);

  return <div ref={containerRef} className="h-full w-full" />;
}
