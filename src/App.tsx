import { useState, useRef, useEffect } from 'react';
import Spreadsheet from '@/components/Spreadsheet';
import ChatPanel from '@/components/ChatPanel';

function App() {
  const [chatVisible, setChatVisible] = useState(true);
  const [chatWidth, setChatWidth] = useState(384); // 96 * 4 = 384px (w-96)
  const isResizingRef = useRef(false);

  // Handle resize
  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (!isResizingRef.current) return;
      
      const newWidth = window.innerWidth - e.clientX;
      // Constrain between min and max widths
      const minWidth = 300;
      const maxWidth = window.innerWidth * 0.5;
      setChatWidth(Math.max(minWidth, Math.min(maxWidth, newWidth)));
    };

    const handleMouseUp = () => {
      isResizingRef.current = false;
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    };

    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);

    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
  }, []);

  const handleResizeStart = () => {
    isResizingRef.current = true;
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
  };

  return (
    <div className="flex h-screen bg-white">
      {/* Left Panel - Spreadsheet */}
      <div className="flex-1 overflow-hidden">
        <Spreadsheet />
      </div>

      {/* Right Panel - AI Chat */}
      {chatVisible && (
        <>
          {/* Resize Handle */}
          <div
            onMouseDown={handleResizeStart}
            className="w-1 bg-gray-200 hover:bg-gray-300 cursor-col-resize transition-colors group relative"
            style={{ flexShrink: 0 }}
          >
            <div className="absolute inset-y-0 -inset-x-1 z-10" />
            {/* Visual indicator */}
            <div className="absolute inset-y-0 left-1/2 -translate-x-1/2 w-0.5 bg-gray-400 opacity-0 group-hover:opacity-100 transition-opacity" />
          </div>
          
          <div 
            className="border-l border-gray-200 bg-white flex flex-col"
            style={{ width: `${chatWidth}px`, flexShrink: 0 }}
          >
            <ChatPanel />
          </div>
        </>
      )}
    </div>
  );
}

export default App;
