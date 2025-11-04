import { useEffect, useRef } from 'react';
import { createUniver, LocaleType, UniverInstanceType, LogLevel, defaultTheme } from '@univerjs/presets';
import { saveWorkbookData as saveWorkbookDataToIndexedDB, loadWorkbookData as loadWorkbookDataFromIndexedDB } from '../src/utils/indexeddb';
import { useUserApiKeys } from '../src/hooks/useUserApiKeys';
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
import UniverSheetsUIEnUS from '@univerjs/sheets-ui/locale/en-US';
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

// Define merge function with deep merging for nested objects
const merge = <T extends Record<string, any>>(target: T, ...sources: any[]): T => {
  const result: any = { ...target };

  for (const source of sources) {
    if (!source) continue;

    for (const key in source) {
      if (source.hasOwnProperty(key)) {
        const sourceValue = source[key];
        const targetValue = result[key];

        // Deep merge if both values are objects and not arrays
        if (
          typeof sourceValue === 'object' &&
          sourceValue !== null &&
          !Array.isArray(sourceValue) &&
          typeof targetValue === 'object' &&
          targetValue !== null &&
          !Array.isArray(targetValue)
        ) {
          result[key] = merge(targetValue, sourceValue);
        } else {
          result[key] = sourceValue;
        }
      }
    }
  }

  return result;
};

interface SpreadsheetProps { }

export default function Spreadsheet({ }: SpreadsheetProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerInstanceRef = useRef<{ univerAPI: any; univer: any } | null>(null);
  const keyDownHandlerRef = useRef<((e: KeyboardEvent) => void) | null>(null);
  const { univerMcpKey } = useUserApiKeys();

  // Register global keyboard handler early, before Univer initializes
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Check if this is a redo key combination first
      const isMac = navigator.platform.toUpperCase().indexOf('MAC') >= 0;
      const isRedo = isMac
        ? (e.metaKey && e.shiftKey && (e.key === 'z' || e.key === 'Z'))
        : ((e.ctrlKey || e.metaKey) && (e.key === 'y' || e.key === 'Y' || (e.shiftKey && (e.key === 'z' || e.key === 'Z'))));

      // Only process redo shortcuts - let everything else pass through immediately
      if (!isRedo) {
        return;
      }

      console.log('🔄 Redo shortcut detected, processing...');

      // Prevent default FIRST to stop any other handlers
      try {
        e.preventDefault();
        e.stopPropagation();
        if (typeof (e as any).stopImmediatePropagation === 'function') {
          (e as any).stopImmediatePropagation();
        }
        console.log('✓ Step 1: Event prevented');
      } catch (preventErr) {
        console.error('❌ Error preventing event:', preventErr);
      }

      // Only handle if we have a univer instance
      if (!univerInstanceRef.current) {
        console.log('⚠ Step 2: No univer instance yet');
        return;
      }
      console.log('✓ Step 2: Univer instance found');

      const target = e.target as HTMLElement;
      console.log('✓ Step 3: Got target', target?.tagName);

      // Skip ONLY if we're in a chat textarea (very specific check)
      if (target?.tagName === 'TEXTAREA') {
        const parent = target.parentElement;
        const hasChatClass = parent?.className?.includes('chat') ||
          target.closest('[class*="ChatPanel"]') ||
          target.closest('[class*="chat-panel"]');
        if (hasChatClass) {
          console.log('⚠ Step 4: In chat textarea, skipping');
          return;
        }
      }
      console.log('✓ Step 4: Not in chat, proceeding...');

      // For everything else (including spreadsheet cells), execute redo
      console.log('✓ Step 5: Executing redo...');

      try {
        const univerAPI = univerInstanceRef.current.univerAPI;

        if (!univerAPI) {
          console.warn('⚠ Step 6: No univerAPI available');
          return;
        }
        console.log('✓ Step 6: Got univerAPI');

        // Use the direct redo() method from univerAPI (async)
        if (typeof univerAPI.redo === 'function') {
          console.log('✓ Step 7: Calling univerAPI.redo()...');
          univerAPI.redo().then(() => {
            console.log('✅ Redo executed successfully via keyboard shortcut!');
          }).catch((err: Error) => {
            console.error('❌ Step 8: Error calling univerAPI.redo():', err);
            // Fallback: try executeCommand with univer.command.redo
            if (typeof univerAPI.executeCommand === 'function') {
              try {
                console.log('✓ Step 9: Trying fallback executeCommand...');
                univerAPI.executeCommand({ id: 'univer.command.redo' });
                console.log('✅ Redo executed via univerAPI.executeCommand');
              } catch (cmdErr) {
                console.error('❌ Failed to execute redo command:', cmdErr);
              }
            }
          });
        } else {
          console.warn('⚠ univerAPI.redo() not available, trying executeCommand...');
          if (typeof univerAPI.executeCommand === 'function') {
            try {
              univerAPI.executeCommand({ id: 'univer.command.redo' });
              console.log('✅ Redo executed via univerAPI.executeCommand');
            } catch (err) {
              console.error('❌ Failed to execute redo command:', err);
            }
          }
        }
      } catch (error) {
        console.error('❌ Fatal error executing redo:', error);
      }
    };

    keyDownHandlerRef.current = handleKeyDown;

    // Attach to window with capture phase for highest priority
    window.addEventListener('keydown', handleKeyDown, true);

    console.log('✓ Global redo keyboard handler registered on window');

    return () => {
      if (keyDownHandlerRef.current) {
        window.removeEventListener('keydown', keyDownHandlerRef.current, true);
      }
    };
  }, []); // Run once on mount

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
    // Only use the user's Univer MCP key from their profile - no env var fallback
    const universerEndpoint = window.location.host;
    const collaboration = undefined;
    const apiKey = univerMcpKey || null;

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
          UniverSheetsUIEnUS,
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

    // Function to save workbook data to IndexedDB
    const saveWorkbookData = async () => {
      try {
        const workbook = univerAPI.getActiveWorkbook();
        if (!workbook) {
          console.warn('No workbook available for saving');
          return;
        }

        const sheets = workbook.getSheets();
        if (!sheets || sheets.length === 0) {
          console.warn('No sheets available for saving');
          return;
        }

        // Use sheet names as keys (more stable than IDs)
        const workbookData: { [sheetName: string]: { name: string; data: any[][] } } = {};

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        sheets.forEach((sheet: any) => {
          try {
            const sheetName = sheet.getName?.() || 'Sheet';

            // Get all data from the sheet
            let allData: any[][] = [];
            const maxRowsToRead = 10000;
            const maxColsToRead = 100;

            try {
              const range = sheet.getRange(0, 0, maxRowsToRead - 1, maxColsToRead - 1);
              const values = range.getValues();
              allData = values || [];

              // Trim empty rows from the end
              while (allData.length > 0 && allData[allData.length - 1].every(cell => !cell || cell === '')) {
                allData.pop();
              }
            } catch (e) {
              console.warn('Could not read all sheet data, reading cell by cell:', e);
              // Fallback: read smaller chunks
              for (let row = 0; row < Math.min(1000, maxRowsToRead); row++) {
                const rowData: any[] = [];
                for (let col = 0; col < Math.min(50, maxColsToRead); col++) {
                  try {
                    const cellRange = sheet.getRange(row, col);
                    const value = cellRange.getValue();
                    rowData.push(value !== null && value !== undefined ? value : '');
                  } catch {
                    rowData.push('');
                  }
                }
                if (rowData.some(cell => cell !== '')) {
                  allData.push(rowData);
                } else if (row > 100 && allData.length > 0) {
                  // Stop if we hit many empty rows
                  break;
                }
              }
            }

            workbookData[sheetName] = {
              name: sheetName,
              data: allData,
            };
          } catch (e) {
            console.warn('Error saving sheet data:', e);
          }
        });

        // Save to IndexedDB
        if (typeof window !== 'undefined') {
          await saveWorkbookDataToIndexedDB(workbookData);
          console.log('✅ Workbook data saved to IndexedDB:', {
            sheetCount: Object.keys(workbookData).length,
            sheetNames: Object.keys(workbookData),
            totalRows: Object.values(workbookData).reduce((sum, sheet) => sum + (sheet.data?.length || 0), 0),
            timestamp: new Date().toISOString(),
          });
        }
      } catch (error) {
        console.error('Error saving workbook to IndexedDB:', error);
      }
    };

    // Function to load workbook data from IndexedDB
    const loadWorkbookData = async () => {
      try {
        if (typeof window === 'undefined') return;

        const savedData = await loadWorkbookDataFromIndexedDB();
        if (!savedData) {
          console.log('No saved data found, creating empty workbook');
          univer.createUnit(UniverInstanceType.UNIVER_SHEET, {});
          return;
        }

        const workbookData: { [sheetName: string]: { name: string; data: any[][] } } = savedData;
        const sheetNames = Object.keys(workbookData);

        if (sheetNames.length === 0) {
          console.log('Empty saved data, creating empty workbook');
          univer.createUnit(UniverInstanceType.UNIVER_SHEET, {});
          return;
        }

        console.log('Loading workbook data from IndexedDB:', {
          sheetCount: sheetNames.length,
          sheetNames: sheetNames,
        });

        // Create workbook with first sheet
        const firstSheetData = workbookData[sheetNames[0]];
        univer.createUnit(UniverInstanceType.UNIVER_SHEET, {});

        // Wait for workbook to be ready with retries
        let retries = 0;
        const maxRetries = 50; // Increased retries for slower systems
        const tryLoadData = () => {
          try {
            const workbook = univerAPI.getActiveWorkbook();
            if (!workbook) {
              retries++;
              if (retries < maxRetries) {
                setTimeout(tryLoadData, 100); // Faster retry interval
                return;
              } else {
                console.error('Workbook not ready after max retries. Retrying load...');
                // Try one more time after a longer delay
                setTimeout(() => {
                  const workbook = univerAPI.getActiveWorkbook();
                  if (workbook) {
                    tryLoadData();
                  } else {
                    console.error('Failed to load workbook after extended retry');
                  }
                }, 1000);
                return;
              }
            }

            // Load first sheet
            const firstSheet = workbook.getActiveSheet();
            if (firstSheet && firstSheetData && firstSheetData.data && firstSheetData.data.length > 0) {
              try {
                const data = firstSheetData.data;
                const maxRow = data.length - 1;
                const maxCol = Math.max(...data.map(row => row.length || 0), 0) - 1;

                if (maxRow >= 0 && maxCol >= 0) {
                  // Pad rows to same length
                  const paddedData = data.map(row => {
                    const padded = [...row];
                    while (padded.length < maxCol + 1) {
                      padded.push('');
                    }
                    return padded.map(cell => cell !== null && cell !== undefined ? String(cell) : '');
                  });

                  const range = firstSheet.getRange(0, 0, maxRow, maxCol);
                  range.setValues(paddedData);
                  console.log(`Loaded ${paddedData.length} rows × ${maxCol + 1} cols into first sheet`);

                  // Rename sheet if name differs
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const currentName = (firstSheet as any).getName?.() || 'Sheet';
                  if (firstSheetData.name && currentName !== firstSheetData.name) {
                    try {
                      firstSheet.setName(firstSheetData.name);
                    } catch (e) {
                      console.warn('Could not rename sheet:', e);
                    }
                  }
                }
              } catch (e) {
                console.error('Error loading first sheet data:', e);
              }
            }

            // Create additional sheets
            sheetNames.slice(1).forEach((sheetName, index) => {
              const sheetData = workbookData[sheetName];
              if (!sheetData) return;

              try {
                const newSheet = workbook.insertSheet(sheetData.name || sheetName || `Sheet${index + 2}`);

                if (sheetData.data && sheetData.data.length > 0) {
                  const data = sheetData.data;
                  const maxRow = data.length - 1;
                  const maxCol = Math.max(...data.map(row => row.length || 0), 0) - 1;

                  if (maxRow >= 0 && maxCol >= 0) {
                    // Pad rows to same length
                    const paddedData = data.map(row => {
                      const padded = [...row];
                      while (padded.length < maxCol + 1) {
                        padded.push('');
                      }
                      return padded.map(cell => cell !== null && cell !== undefined ? String(cell) : '');
                    });

                    const range = newSheet.getRange(0, 0, maxRow, maxCol);
                    range.setValues(paddedData);
                    console.log(`Loaded ${paddedData.length} rows × ${maxCol + 1} cols into sheet "${sheetData.name}"`);
                  }
                }
              } catch (e) {
                console.error(`Error creating sheet ${sheetData.name}:`, e);
              }
            });

            console.log('✅ Workbook data loaded successfully from IndexedDB');
          } catch (error) {
            console.error('❌ Error loading workbook data:', error);
          }
        };

        // Start loading after a delay to ensure workbook is ready
        // Increased delay to give Univer time to fully initialize
        setTimeout(tryLoadData, 500);

      } catch (error) {
        console.error('Error loading workbook from IndexedDB:', error);
        // Create empty workbook as fallback
        univer.createUnit(UniverInstanceType.UNIVER_SHEET, {});
      }
    };

    // Load saved data first
    loadWorkbookData();

    // Store instance in ref to prevent double initialization
    univerInstanceRef.current = { univerAPI, univer };

    // Store API globally for MCP client connection
    let autoSaveInterval: ReturnType<typeof setInterval> | null = null;
    let saveTimeout: ReturnType<typeof setTimeout> | null = null;
    let beforeUnloadHandler: (() => void) | null = null;

    if (typeof window !== 'undefined') {
      (window as any).univerAPI = univerAPI;
      // Expose save function globally so other components can trigger saves
      (window as any).saveWorkbookData = saveWorkbookData;
      // Expose debug function to check IndexedDB
      (window as any).debugWorkbookData = async () => {
        try {
          const saved = await loadWorkbookDataFromIndexedDB();
          if (saved) {
            console.log('Saved workbook data:', {
              sheetCount: Object.keys(saved).length,
              sheets: Object.keys(saved).map(name => ({
                name,
                rowCount: saved[name].data?.length || 0,
                colCount: saved[name].data?.[0]?.length || 0,
              })),
              fullData: saved,
            });
            return saved;
          } else {
            console.log('No saved workbook data found');
            return null;
          }
        } catch (error) {
          console.error('Error reading workbook data:', error);
          return null;
        }
      };

      // Expose debug function to test redo manually
      (window as any).testRedo = () => {
        if (!univerInstanceRef.current) {
          console.error('No univer instance available');
          return;
        }

        const univerAPI = univerInstanceRef.current.univerAPI;

        if (!univerAPI) {
          console.error('No univerAPI found');
          return;
        }

        console.log('🔍 Testing redo via univerAPI...');

        // Method 1: Direct redo() method (async)
        if (typeof univerAPI.redo === 'function') {
          console.log('✓ Found univerAPI.redo() method');
          univerAPI.redo().then(() => {
            console.log('✓ Redo executed successfully via univerAPI.redo()');
          }).catch((err: Error) => {
            console.error('❌ Error calling univerAPI.redo():', err);
          });
          return;
        } else {
          console.warn('⚠ univerAPI.redo() is not available');
        }

        // Method 2: executeCommand with univer.command.redo
        if (typeof univerAPI.executeCommand === 'function') {
          console.log('✓ Found univerAPI.executeCommand() method');
          try {
            const result = univerAPI.executeCommand({ id: 'univer.command.redo' });
            console.log('✓ Redo executed successfully via executeCommand:', result);
            return;
          } catch (err) {
            console.error('❌ Error calling executeCommand:', err);
          }
        } else {
          console.warn('⚠ univerAPI.executeCommand() is not available');
        }

        console.warn('⚠ Could not execute redo - no methods available');
      };

      // Set up auto-save with debouncing
      const debouncedSave = () => {
        if (saveTimeout) clearTimeout(saveTimeout);
        saveTimeout = setTimeout(async () => {
          await saveWorkbookData();
        }, 2000); // Save 2 seconds after last change
      };

      // Listen for changes (we'll poll since Univer might not expose change events directly)
      // Save periodically and after operations
      autoSaveInterval = setInterval(() => {
        debouncedSave();
      }, 10000); // Auto-save every 10 seconds

      // Also save when window is about to close
      beforeUnloadHandler = () => {
        // Use synchronous version for beforeunload
        saveWorkbookData().catch(console.error);
      };
      window.addEventListener('beforeunload', beforeUnloadHandler);

      // Note: Keyboard handler is registered in a separate useEffect above
      // to ensure it's attached before Univer initializes
    }

    // Clean up on unmount
    return () => {
      if (autoSaveInterval) clearInterval(autoSaveInterval);
      if (saveTimeout) clearTimeout(saveTimeout);
      if (beforeUnloadHandler && typeof window !== 'undefined') {
        window.removeEventListener('beforeunload', beforeUnloadHandler);
      }
      // Keyboard handler cleanup is in the separate useEffect above
      saveWorkbookData().catch(console.error); // Final save
      if (univerInstanceRef.current) {
        univerInstanceRef.current.univerAPI.dispose();
        univerInstanceRef.current = null;
      }
    };

    console.log('Univer MCP configured with sessionId:', sessionId);
  }, []);

  return <div ref={containerRef} className="h-full w-full" />;
}
