'use client';

import { useEffect, useRef } from 'react';
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/presets/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/presets/preset-sheets-core/locales/en-US';
import '@univerjs/presets/lib/styles/preset-sheets-core.css';

interface SpreadsheetProps {
  data?: any[][];
  onCellUpdate?: (row: number, col: number, value: string) => void;
  getColumnLetter: (col: number) => string;
}

export default function Spreadsheet({ data, onCellUpdate, getColumnLetter }: SpreadsheetProps) {
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!containerRef.current) return;

    const { univerAPI } = createUniver({
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: mergeLocales(
          UniverPresetSheetsCoreEnUS,
        ),
      },
      presets: [
        UniverSheetsCorePreset({
          container: containerRef.current!,
        }),
      ],
    });

    univerAPI.createWorkbook({});

    // Store API globally for later use
    if (typeof window !== 'undefined') {
      (window as any).univerAPI = univerAPI;
    }

    return () => {
      univerAPI.dispose();
    };
  }, []);

  return <div ref={containerRef} className="h-full w-full" />;
}
