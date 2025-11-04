import { useEffect, useRef } from 'react';
import { createUniver, LocaleType, UniverInstanceType, LogLevel, defaultTheme } from '@univerjs/presets';
import { saveWorkbookData as saveWorkbookDataToIndexedDB, loadWorkbookData as loadWorkbookDataFromIndexedDB } from '../src/utils/indexeddb';
import { useUserApiKeys } from '../src/hooks/useUserApiKeys';
import { importXLSXToWorkbookData } from '../src/utils/xlsxConverter';
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
  const isLoadingDataRef = useRef<boolean>(true); // Track if data is still loading
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
          // Removed universerEndpoint - we use client-side import/export instead
          // Setting this would try to connect to a server that doesn't exist, causing 405 errors
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

    // Function to save workbook data to IndexedDB using Univer's official save() method
    const saveWorkbookData = async () => {
      // Don't save while data is still loading
      if (isLoadingDataRef.current) {
        console.log('⏸ Skipping save - data is still loading');
        return;
      }

      if (!univerInstanceRef.current) {
        console.log('⏸ Skipping save - no Univer instance');
        return;
      }

      const { univerAPI } = univerInstanceRef.current;
      const workbook = univerAPI.getActiveWorkbook();

      if (!workbook) {
        console.log('⏸ Skipping save - no active workbook');
        return;
      }

      try {
        const workbookSnapshot = workbook.save();

        // Ensure all sheets have at least 100,000 rows and 1,000 columns
        // This allows users to scroll to these limits. Univer's virtual scrolling handles rendering efficiently.
        const MIN_ROWS = 100000;
        const MIN_COLS = 1000;
        if (workbookSnapshot && workbookSnapshot.sheets) {
          for (const sheetId in workbookSnapshot.sheets) {
            const sheet = workbookSnapshot.sheets[sheetId];
            if (sheet) {
              // Set rowCount and columnCount to at least MIN_ROWS/MIN_COLS if not already set or if they're lower
              if (!(sheet as any).rowCount || (sheet as any).rowCount < MIN_ROWS) {
                (sheet as any).rowCount = MIN_ROWS;
              }
              if (!(sheet as any).columnCount || (sheet as any).columnCount < MIN_COLS) {
                (sheet as any).columnCount = MIN_COLS;
              }
            }
          }
        }

        console.log('💾 Saving workbook data to IndexedDB...');
        await saveWorkbookDataToIndexedDB(workbookSnapshot);
        console.log('✅ Workbook data saved successfully');
      } catch (error) {
        console.error('❌ Error saving workbook data:', error);
      }
    };

    // Function to load workbook data from IndexedDB using Univer's official createWorkbook() method
    const loadWorkbookData = async () => {
      try {
        if (typeof window === 'undefined') return;

        const savedData = await loadWorkbookDataFromIndexedDB();
        if (!savedData) {
          console.log('No saved data found, creating empty workbook');
          univer.createUnit(UniverInstanceType.UNIVER_SHEET, {});
          // Mark loading as complete immediately since there's nothing to load
          isLoadingDataRef.current = false;
          return;
        }

        // Check if savedData is in the old format (custom format) or new format (IWorkbookData)
        // If it has sheets property and sheetOrder, it's IWorkbookData
        // If it has sheet names as keys, it's the old format
        let workbookData: any = savedData;

        if (!savedData.sheets || !savedData.sheetOrder) {
          // Old format detected - log warning but continue with empty workbook
          console.warn('Old data format detected, creating empty workbook. Please re-save your data.');
          univer.createUnit(UniverInstanceType.UNIVER_SHEET, {});
          isLoadingDataRef.current = false;
          return;
        }

        console.log('Loading workbook data from IndexedDB:', {
          workbookId: workbookData.id,
          workbookName: workbookData.name,
          sheetCount: workbookData.sheetOrder?.length || 0,
          sheetNames: workbookData.sheetOrder || [],
        });

        // Ensure all sheets have at least 100,000 rows and 1,000 columns when loading
        const MIN_ROWS = 100000;
        const MIN_COLS = 1000;
        if (workbookData && workbookData.sheets) {
          for (const sheetId in workbookData.sheets) {
            const sheet = workbookData.sheets[sheetId];
            if (sheet) {
              // Set rowCount and columnCount to at least MIN_ROWS/MIN_COLS if not already set or if they're lower
              if (!(sheet as any).rowCount || (sheet as any).rowCount < MIN_ROWS) {
                (sheet as any).rowCount = MIN_ROWS;
              }
              if (!(sheet as any).columnCount || (sheet as any).columnCount < MIN_COLS) {
                (sheet as any).columnCount = MIN_COLS;
              }
            }
          }
        }

        // Use Univer's official createWorkbook() method to restore the complete workbook
        try {
          const createdWorkbook = univerAPI.createWorkbook(workbookData);
          console.log('✅ Workbook created from saved data:', createdWorkbook?.getId());
          // Mark loading as complete - safe to save now
          isLoadingDataRef.current = false;
        } catch (error) {
          console.error('❌ Error creating workbook from saved data:', error);
          // Fallback to empty workbook
          univer.createUnit(UniverInstanceType.UNIVER_SHEET, {});
          // Mark loading as complete even on error so we don't block saves forever
          isLoadingDataRef.current = false;
        }

      } catch (error) {
        console.error('Error loading workbook from IndexedDB:', error);
        // Create empty workbook as fallback
        univer.createUnit(UniverInstanceType.UNIVER_SHEET, {});
        // Mark loading as complete even on error so we don't block saves forever
        isLoadingDataRef.current = false;
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
        }, 500); // Save 500ms after last change (reduced for faster saves)
      };

      // Immediate save function (no debounce)
      const immediateSave = async () => {
        if (saveTimeout) clearTimeout(saveTimeout);
        await saveWorkbookData();
      };

      // Listen for keyboard events that finalize cell edits
      const handleCellEditFinish = (e: KeyboardEvent) => {
        // Enter, Tab, Escape typically finalize cell edits in spreadsheets
        if (e.key === 'Enter' || e.key === 'Tab' || e.key === 'Escape') {
          // Save immediately when cell edit is finalized
          immediateSave().catch(console.error);
        }
      };
      window.addEventListener('keydown', handleCellEditFinish);

      // Also save when user clicks outside the spreadsheet (blur)
      const handleWindowBlur = () => {
        immediateSave().catch(console.error);
      };
      window.addEventListener('blur', handleWindowBlur);

      // Store handlers for cleanup
      (univerInstanceRef.current as any).cellEditFinishHandler = handleCellEditFinish;
      (univerInstanceRef.current as any).windowBlurHandler = handleWindowBlur;

      // Try to listen to Univer's command execution events if available
      // Wait a bit for workbook to be ready
      setTimeout(() => {
        try {
          const workbook = univerAPI.getActiveWorkbook();
          // Use type assertion since getCommandService might not be in types but could exist at runtime
          const commandService = (workbook as any)?.getCommandService?.();
          if (commandService && typeof commandService.onCommandExecuted === 'function') {
            // Listen for set range values command (cell edits)
            const dispose = commandService.onCommandExecuted((commandInfo: any) => {
              if (commandInfo?.id && (
                commandInfo.id === 'sheet.command.set-range-values' ||
                commandInfo.id === 'sheet.operation.set-range-values' ||
                commandInfo.id.includes('set-range') ||
                commandInfo.id.includes('set-value') ||
                commandInfo.id.includes('set-range-values')
              )) {
                console.log('Cell edit detected via command event, triggering save');
                debouncedSave();
              }
            });
            // Store dispose function for cleanup (if available)
            if (typeof dispose === 'function') {
              if (!univerInstanceRef.current) univerInstanceRef.current = { univerAPI, univer };
              (univerInstanceRef.current as any).commandDispose = dispose;
            }
          }
        } catch (e) {
          console.log('Could not hook into Univer command events:', e);
        }
      }, 1000); // Wait 1 second for workbook to be ready

      // Listen for changes (polling as fallback)
      // Save more frequently to catch changes quickly
      autoSaveInterval = setInterval(() => {
        debouncedSave();
      }, 3000); // Auto-save every 3 seconds (reduced from 10)

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
      // Remove keyboard and blur event listeners
      if (univerInstanceRef.current) {
        const cellEditFinishHandler = (univerInstanceRef.current as any).cellEditFinishHandler;
        const windowBlurHandler = (univerInstanceRef.current as any).windowBlurHandler;
        if (cellEditFinishHandler) {
          window.removeEventListener('keydown', cellEditFinishHandler);
        }
        if (windowBlurHandler) {
          window.removeEventListener('blur', windowBlurHandler);
        }
      }
      // Clean up command listener if it exists
      if (univerInstanceRef.current && (univerInstanceRef.current as any).commandDispose) {
        try {
          (univerInstanceRef.current as any).commandDispose();
        } catch (e) {
          console.warn('Error disposing command listener:', e);
        }
      }
      // Keyboard handler cleanup is in the separate useEffect above
      saveWorkbookData().catch(console.error); // Final save

      // Cleanup global functions and file input
      delete (window as any).exportWorkbookToXLSX;
      delete (window as any).importXLSXFile;
      if (univerInstanceRef.current && (univerInstanceRef.current as any).fileInput) {
        ((univerInstanceRef.current as any).fileInput as HTMLInputElement).remove();
      }

      if (univerInstanceRef.current) {
        univerInstanceRef.current.univerAPI.dispose();
        univerInstanceRef.current = null;
      }
    };

    console.log('Univer MCP configured with sessionId:', sessionId);

    // Expose export function globally so ChatPanel can call it
    (window as any).exportWorkbookToXLSX = async (filename?: string) => {
      try {
        const workbook = univerAPI.getActiveWorkbook();
        if (!workbook) {
          throw new Error('No workbook available');
        }
        const workbookSnapshot = workbook.save();
        if (!workbookSnapshot) {
          throw new Error('Failed to get workbook snapshot');
        }
        // Import the export function dynamically to avoid circular dependencies
        const { exportWorkbookToXLSX } = await import('../src/utils/xlsxConverter');
        await exportWorkbookToXLSX(workbookSnapshot, filename || 'workbook.xlsx');
      } catch (error) {
        console.error('Error exporting workbook:', error);
        throw error;
      }
    };

    // Add file input for importing XLSX files
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.xlsx,.xls';
    fileInput.style.display = 'none';
    fileInput.addEventListener('change', async (e) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (!file) return;

      try {
        // Import XLSX file
        const workbookData = await importXLSXToWorkbookData(file);

        // Check if it's the old data format (for backwards compatibility)
        if (!workbookData.sheets || !workbookData.sheetOrder) {
          console.warn('Invalid workbook data format');
          return;
        }

        // Ensure all sheets have at least 100,000 rows and 1,000 columns
        // The XLSX import already sets these, but we ensure they're at least MIN_ROWS/MIN_COLS
        const MIN_ROWS = 100000;
        const MIN_COLS = 1000;
        if (workbookData && workbookData.sheets) {
          for (const sheetId in workbookData.sheets) {
            const sheet = workbookData.sheets[sheetId];
            if (sheet) {
              // Set rowCount and columnCount to at least MIN_ROWS/MIN_COLS if not already set or if they're lower
              if (!(sheet as any).rowCount || (sheet as any).rowCount < MIN_ROWS) {
                (sheet as any).rowCount = MIN_ROWS;
              }
              if (!(sheet as any).columnCount || (sheet as any).columnCount < MIN_COLS) {
                (sheet as any).columnCount = MIN_COLS;
              }
            }
          }
        }

        // Set loading flag to prevent auto-save from overwriting
        isLoadingDataRef.current = true;

        // Create workbook from imported data
        try {
          const createdWorkbook = univerAPI.createWorkbook(workbookData);
          console.log('✅ Workbook imported from XLSX:', createdWorkbook?.getId());

          // Save imported data to IndexedDB
          await saveWorkbookDataToIndexedDB(workbookData);
          console.log('✅ Imported workbook saved to IndexedDB');
        } catch (error) {
          console.error('❌ Error creating workbook from imported data:', error);
          throw error;
        } finally {
          isLoadingDataRef.current = false;
        }
      } catch (error) {
        console.error('Error importing XLSX file:', error);
        alert(`Failed to import file: ${error instanceof Error ? error.message : 'Unknown error'}`);
      } finally {
        // Reset file input
        fileInput.value = '';
      }
    });

    // Expose import function globally
    (window as any).importXLSXFile = () => {
      fileInput.click();
    };

    // Store fileInput reference for cleanup
    (univerInstanceRef.current as any).fileInput = fileInput;
  }, []);

  return <div ref={containerRef} className="h-full w-full" />;
}
