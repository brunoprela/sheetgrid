'use client';

import { useEffect, useRef } from 'react';
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/presets/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/presets/preset-sheets-core/locales/en-US';
import { UniverMCPPlugin } from '@univerjs-pro/mcp';
import { UniverMCPUIPlugin } from '@univerjs-pro/mcp-ui';
import { UniverSheetMCPPlugin } from '@univerjs-pro/sheets-mcp';
import UniverMCPUIEnUS from '@univerjs-pro/mcp-ui/locale/en-US';
import '@univerjs-pro/mcp/facade';
import '@univerjs/presets/lib/styles/preset-sheets-core.css';
import '@univerjs-pro/mcp-ui/lib/index.css';

interface SpreadsheetProps {}

export default function Spreadsheet({}: SpreadsheetProps) {
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!containerRef.current) return;

    // Generate a unique session ID
    const sessionId = `univer-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;

    // Create Univer instance with preset and MCP plugins
    const { univerAPI, univer } = createUniver({
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: mergeLocales(
          UniverPresetSheetsCoreEnUS,
          UniverMCPUIEnUS,
        ),
      },
      presets: [
        UniverSheetsCorePreset({
          container: containerRef.current!,
          header: true,
        }),
      ],
      plugins: [
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
    univerAPI.createWorkbook({});

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
