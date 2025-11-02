'use client';

import { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import dynamic from 'next/dynamic';

const Spreadsheet = dynamic(() => import('@/components/Spreadsheet'), { ssr: false });
const ChatPanel = dynamic(() => import('@/components/ChatPanel'), { ssr: false });

export default function Home() {
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [activeSheetName, setActiveSheetName] = useState<string>('Sheet1');
  const [workbookData, setWorkbookData] = useState<{ [sheetName: string]: any[][] }>({ Sheet1: [[]] });
  const [chatVisible, setChatVisible] = useState(true);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Initialize with an empty sheet
  useEffect(() => {
    if (!workbook) {
      const newWb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([[]]);
      XLSX.utils.book_append_sheet(newWb, ws, 'Sheet1');
      setWorkbook(newWb);
    }
  }, [workbook]);

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        setWorkbook(wb);

        // Convert all sheets to array format
        const sheets: { [sheetName: string]: any[][] } = {};
        wb.SheetNames.forEach((sheetName) => {
          const ws = wb.Sheets[sheetName];
          sheets[sheetName] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        });

        setWorkbookData(sheets);
        setActiveSheetName(wb.SheetNames[0] || '');
      } catch (error) {
        console.error('Error reading file:', error);
        alert('Failed to read Excel file. Please try again.');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleDownload = () => {
    // Convert workbook data back to XLSX format
    const newWb = XLSX.utils.book_new();
    Object.entries(workbookData).forEach(([sheetName, data]) => {
      const ws = XLSX.utils.aoa_to_sheet(data);
      XLSX.utils.book_append_sheet(newWb, ws, sheetName);
    });

    XLSX.writeFile(newWb, `sheetgrid_${new Date().getTime()}.xlsx`);
  };

  const updateCell = (row: number, col: number, value: string, sheetName?: string) => {
    const targetSheet = sheetName || activeSheetName;
    const newData = { ...workbookData };
    if (!newData[targetSheet][row]) {
      newData[targetSheet][row] = [];
    }
    newData[targetSheet][row][col] = value;
    setWorkbookData(newData);
  };

  const getColumnLetter = (col: number): string => {
    let result = '';
    let num = col;
    while (num >= 0) {
      result = String.fromCharCode(65 + (num % 26)) + result;
      num = Math.floor(num / 26) - 1;
    }
    return result;
  };

  return (
    <div className="flex h-screen bg-white">
      {/* Left Panel - Spreadsheet */}
      <div className="flex-1 flex flex-col overflow-hidden">
        {/* Top Toolbar */}
        <TopToolbar 
          onDownload={handleDownload} 
          fileInputRef={fileInputRef} 
          onFileUpload={handleFileUpload}
          chatVisible={chatVisible}
          onToggleChat={() => setChatVisible(!chatVisible)}
        />

        {/* Formula Bar */}
        <FormulaBar />

        {/* Spreadsheet */}
        <div className="flex-1 overflow-auto">
          <Spreadsheet
            data={workbookData[activeSheetName] || []}
            onCellUpdate={updateCell}
            getColumnLetter={getColumnLetter}
          />
        </div>

        {/* Bottom Sheet Tabs */}
        {workbook && (
          <SheetTabs
            sheets={workbook.SheetNames}
            activeSheet={activeSheetName}
            onSheetChange={setActiveSheetName}
          />
        )}
      </div>

      {/* Right Panel - AI Chat */}
      {chatVisible && (
        <div className="w-96 border-l border-gray-200 bg-white flex flex-col">
          <ChatPanel
            workbookData={workbookData}
            updateCell={updateCell}
            setWorkbookData={setWorkbookData}
            getColumnLetter={getColumnLetter}
            activeSheetName={activeSheetName}
          />
        </div>
      )}
    </div>
  );
}

function TopToolbar({ onDownload, fileInputRef, onFileUpload, chatVisible, onToggleChat }: { 
  onDownload: () => void; 
  fileInputRef: React.RefObject<HTMLInputElement | null>; 
  onFileUpload: (event: React.ChangeEvent<HTMLInputElement>) => void;
  chatVisible: boolean;
  onToggleChat: () => void;
}) {
  return (
    <>
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={onFileUpload}
        className="hidden"
      />
      <div className="border-b border-gray-200 bg-white px-4 py-2">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-4">
            {/* Logo */}
            <div className="flex items-center gap-2">
              <div className="w-8 h-8 bg-green-600 rounded flex items-center justify-center">
                <span className="text-white font-bold">X</span>
              </div>
              <span className="font-medium text-gray-900">New SheetGrid Workbook</span>
            </div>

            {/* Menu Items */}
            <div className="flex items-center gap-2 ml-8">
            <button 
              onClick={() => fileInputRef.current?.click()}
              className="px-3 py-1.5 hover:bg-gray-100 rounded text-sm text-gray-700 flex items-center gap-1"
            >
              Upload
              <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
              </svg>
            </button>
            <button className="px-3 py-1.5 hover:bg-gray-100 rounded text-sm text-gray-700 flex items-center gap-1">
              document
              <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
              </svg>
            </button>
            <button className="px-3 py-1.5 hover:bg-gray-100 rounded text-sm text-gray-700 flex items-center gap-1">
              insert
              <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
              </svg>
            </button>
            <button className="px-3 py-1.5 hover:bg-gray-100 rounded text-sm text-gray-700 flex items-center gap-1">
              view
              <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
              </svg>
            </button>
            <button className="px-3 py-1.5 hover:bg-gray-100 rounded text-sm text-gray-700">
              More
            </button>
          </div>
        </div>

        {/* Right Side */}
        <div className="flex items-center gap-2">
          <button
            onClick={onToggleChat}
            className="p-2 hover:bg-gray-100 rounded transition-colors"
            title={chatVisible ? "Hide Chat" : "Show Chat"}
          >
            <svg className="w-5 h-5 text-gray-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 10h.01M12 10h.01M16 10h.01M9 16H5a2 2 0 01-2-2V6a2 2 0 012-2h14a2 2 0 012 2v8a2 2 0 01-2 2h-5l-5 5v-5z" />
            </svg>
          </button>
          <button
            onClick={onDownload}
            className="px-4 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded"
          >
            Download
          </button>
          <div className="w-8 h-8 bg-gray-300 rounded-full"></div>
        </div>
      </div>
    </div>
    </>
  );
}

function FormulaBar() {
  return (
    <div className="border-b border-gray-200 bg-white px-4 py-2 flex items-center gap-2">
      <div className="w-12 text-sm text-gray-600 border-r border-gray-200 pr-2">A1</div>
      <div className="flex items-center gap-1 flex-1">
        <div className="text-gray-500 mr-2">X</div>
        <div className="text-green-600">✓</div>
        <div className="text-blue-600 font-bold">fx</div>
        <input
          type="text"
          className="flex-1 px-2 py-1 border border-gray-300 rounded bg-white text-sm"
          placeholder=""
        />
      </div>
    </div>
  );
}

function SheetTabs({ sheets, activeSheet, onSheetChange }: { sheets: string[]; activeSheet: string; onSheetChange: (sheet: string) => void }) {
  return (
    <div className="border-t border-gray-200 bg-gray-50 px-2 py-2 flex items-center gap-1">
      <button className="px-3 py-1 hover:bg-gray-200 rounded text-sm text-gray-700">
        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" />
        </svg>
      </button>
      <div className="flex gap-1 flex-1">
        {sheets.map((sheet) => (
          <button
            key={sheet}
            onClick={() => onSheetChange(sheet)}
            className={`px-4 py-1 rounded text-sm font-medium transition-colors ${
              activeSheet === sheet
                ? 'bg-white shadow border border-gray-200 text-gray-900'
                : 'text-gray-600 hover:bg-gray-200'
            }`}
          >
            {sheet}
          </button>
        ))}
      </div>
    </div>
  );
}
