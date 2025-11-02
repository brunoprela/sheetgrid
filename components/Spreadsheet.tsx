'use client';

import { useEffect, useRef } from 'react';
import { createUniver, LocaleType, UniverInstanceType } from '@univerjs/presets';
import { mergeLocales } from '@univerjs/core';
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

interface SpreadsheetProps {}

export default function Spreadsheet({}: SpreadsheetProps) {
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!containerRef.current) return;

    // Generate a unique session ID
    const sessionId = `univer-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;

    // Create Univer instance with preset and MCP plugins
    const universerEndpoint = window.location.host;
    const collaboration = undefined;

    const { univerAPI, univer } = createUniver({
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: mergeLocales(
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
        [UniverMCPPlugin, {
          sessionId,
          ticketServerUrl: 'https://mcp.univer.ai/api/ticket',
          mcpServerUrl: 'wss://mcp.univer.ai/api/ws',
        }],
        [UniverMCPUIPlugin, {
          showDeveloperTools: true,
        }],
        UniverSheetMCPPlugin,
      ],
    });

    // Create a workbook
    univer.createUnit(UniverInstanceType.UNIVER_SHEET, {});

    // Store API and session ID globally for MCP client connection
    if (typeof window !== 'undefined') {
      (window as any).univerAPI = univerAPI;
      (window as any).univerSessionId = sessionId;
    }

    return () => {
      univerAPI.dispose();
    };
  }, []);

  return <div ref={containerRef} className="h-full w-full" />;
}
