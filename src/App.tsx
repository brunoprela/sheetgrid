import { useState, useRef, useEffect } from 'react';
import { Routes, Route, useLocation } from 'react-router-dom';
import { useUser, useStackApp, StackHandler } from '@stackframe/react';
import { stackClientApp } from './providers/StackAuthProvider';
import Spreadsheet from '@/components/Spreadsheet';
import ChatPanel from '@/components/ChatPanel';
import LandingPage from '@/components/LandingPage';
import SignInPage from '@/components/SignInPage';
import AuthButton from '@/components/AuthButton';
import UserProfile from '@/components/UserProfile';
import { getAllChats, createChat, deleteChat, type Chat } from './utils/indexeddb';

function HandlerRoutes() {
  const location = useLocation();
  return <StackHandler app={stackClientApp} location={location.pathname} fullPage />;
}

function App() {
  const user = useUser();
  const stackApp = useStackApp();
  
  // Track if we should show sign-in page
  const [showSignIn, setShowSignIn] = useState(false);

  // Show landing page only when user is NOT signed in and not showing sign-in
  const showLanding = !user && !showSignIn;

  // Load saved chat width from localStorage on mount
  const [chatVisible] = useState(true);
  const [chatWidth, setChatWidth] = useState(() => {
    if (typeof window !== 'undefined') {
      const saved = localStorage.getItem('chatPanelWidth');
      return saved ? parseInt(saved, 10) : 384; // Default to 384px if no saved value
    }
    return 384; // Default for SSR
  });
  const [activeChatId, setActiveChatId] = useState<string | null>(null);
  const [activeChatTitle, setActiveChatTitle] = useState<string>('New Chat');
  const [allChats, setAllChats] = useState<Chat[]>([]);
  const [showUserProfile, setShowUserProfile] = useState(false);
  const isResizingRef = useRef(false);

  // Save chat width to localStorage whenever it changes
  useEffect(() => {
    if (typeof window !== 'undefined') {
      localStorage.setItem('chatPanelWidth', chatWidth.toString());
    }
  }, [chatWidth]);

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

  // Load all chats and initialize
  const loadAllChats = async () => {
    const chats = await getAllChats();
    setAllChats(chats);
    return chats;
  };

  // Disable swipe back/forward navigation gestures
  useEffect(() => {
    let touchStartX = 0;
    let touchStartY = 0;

    const handleTouchStart = (e: TouchEvent) => {
      touchStartX = e.touches[0].clientX;
      touchStartY = e.touches[0].clientY;
    };

    const handleTouchMove = (e: TouchEvent) => {
      if (!touchStartX || !touchStartY) return;

      const touchEndX = e.touches[0].clientX;
      const touchEndY = e.touches[0].clientY;
      const deltaX = touchEndX - touchStartX;
      const deltaY = touchEndY - touchStartY;

      // Check if the touch is on a scrollable element
      const target = e.target as HTMLElement;
      const scrollableParent = target.closest('[class*="overflow"], [class*="scroll"]');
      const isInScrollableElement = scrollableParent && 
        (scrollableParent.scrollWidth > scrollableParent.clientWidth || 
         scrollableParent.scrollHeight > scrollableParent.clientHeight);

      // Only prevent horizontal swipe if:
      // 1. It's a horizontal swipe (deltaX > deltaY)
      // 2. The swipe is significant (> 30px)
      // 3. It's not within a scrollable element
      // 4. It starts from the edge of the screen (likely a navigation gesture)
      if (!isInScrollableElement && 
          Math.abs(deltaX) > Math.abs(deltaY) && 
          Math.abs(deltaX) > 30 &&
          (touchStartX < 20 || touchStartX > window.innerWidth - 20)) {
        e.preventDefault();
      }
    };

    // Prevent browser back/forward navigation via mousewheel swipe
    const handleWheel = (e: WheelEvent) => {
      // Only prevent if it's a clear horizontal swipe from edge
      const target = e.target as HTMLElement;
      const isScrollable = target.scrollWidth > target.clientWidth;
      
      if (!isScrollable && Math.abs(e.deltaX) > Math.abs(e.deltaY) && Math.abs(e.deltaX) > 50) {
        // Check if cursor is near screen edge
        const mouseX = e.clientX;
        if (mouseX < 50 || mouseX > window.innerWidth - 50) {
          e.preventDefault();
        }
      }
    };

    // Prevent popstate (back/forward navigation)
    const handlePopState = () => {
      // Prevent navigation if it's not a user-initiated action
      if (window.history.state === null) {
        window.history.pushState({ preventBack: true }, '');
      }
    };

    // Push initial state to prevent back navigation
    window.history.pushState({ preventBack: true }, '');

    document.addEventListener('touchstart', handleTouchStart, { passive: true });
    document.addEventListener('touchmove', handleTouchMove, { passive: false });
    window.addEventListener('popstate', handlePopState);
    document.addEventListener('wheel', handleWheel, { passive: false });

    return () => {
      document.removeEventListener('touchstart', handleTouchStart);
      document.removeEventListener('touchmove', handleTouchMove);
      window.removeEventListener('popstate', handlePopState);
      document.removeEventListener('wheel', handleWheel);
    };
  }, []);

  // Load or create initial chat - prioritize loading from IndexedDB
  useEffect(() => {
    const initializeChat = async () => {
      try {
        // Always try to load existing chats from IndexedDB first
        const chats = await loadAllChats();
        
        if (chats.length > 0) {
          // Load most recently updated chat (chats are sorted by updatedAt descending)
          const latestChat = chats[0];
          setActiveChatId(latestChat.id);
          setActiveChatTitle(latestChat.title);
          console.log('Loaded existing chat from IndexedDB:', latestChat.title);
        } else {
          // Only create a new chat if absolutely no chats exist
          console.log('No existing chats found, creating new chat');
          const newChat = await createChat();
          setActiveChatId(newChat.id);
          setActiveChatTitle(newChat.title);
          await loadAllChats(); // Reload to include new chat
        }
      } catch (error) {
        console.error('Error initializing chat:', error);
        // On error, don't create a new chat - let user see the error state
      }
    };
    initializeChat();
  }, []);

  // Update chat title and reload chats when switching chats
  useEffect(() => {
    const loadChatTitle = async () => {
      if (activeChatId) {
        await loadAllChats(); // Refresh chat list
        const chats = await getAllChats();
        const chat = chats.find(c => c.id === activeChatId);
        if (chat) {
          setActiveChatTitle(chat.title);
        }
      }
    };
    loadChatTitle();
  }, [activeChatId]);

  // Note: handleChatCreated is defined but not currently used - keeping for future use
  // TypeScript configuration allows unused functions
  const handleChatCreated = async (chat: Chat) => {
    setActiveChatId(chat.id);
    setActiveChatTitle(chat.title);
  };

  const handleCreateNewChat = async () => {
    try {
      const newChat = await createChat();
      setActiveChatId(newChat.id);
      setActiveChatTitle(newChat.title);
      await loadAllChats(); // Refresh chat list
    } catch (error) {
      console.error('Error creating new chat:', error);
    }
  };

  const handleSelectChat = async (chatId: string) => {
    setActiveChatId(chatId);
    await loadAllChats(); // Refresh to get latest titles
  };

  const handleDeleteChat = async (chatId: string) => {
    try {
      await deleteChat(chatId);
      await loadAllChats(); // Refresh chat list
      
      // If we deleted the active chat, switch to another one
      if (activeChatId === chatId) {
        const updatedChats = await getAllChats();
        if (updatedChats.length > 0) {
          setActiveChatId(updatedChats[0].id);
          setActiveChatTitle(updatedChats[0].title);
        } else {
          // No chats left, create a new one
          const newChat = await createChat();
          setActiveChatId(newChat.id);
          setActiveChatTitle(newChat.title);
          await loadAllChats();
        }
      }
    } catch (error) {
      console.error('Error deleting chat:', error);
    }
  };

  const handleResizeStart = () => {
    isResizingRef.current = true;
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
  };

  const handleGetStarted = () => {
    // Show sign-in page when button is clicked
    setShowSignIn(true);
  };

  const location = useLocation();
  
  // Check if we're on a handler route - these should be handled by StackHandler
  const isHandlerRoute = location.pathname.startsWith('/handler/');

  // Allow scrolling on landing/sign-in pages when zoomed, prevent it on main app
  useEffect(() => {
    if (showLanding || showSignIn) {
      document.documentElement.style.overflow = 'auto';
      document.body.style.overflow = 'auto';
    } else {
      document.documentElement.style.overflow = 'hidden';
      document.body.style.overflow = 'hidden';
    }
  }, [showLanding, showSignIn]);

  // Reset sign-in page when user signs in
  useEffect(() => {
    if (user && showSignIn) {
      setShowSignIn(false);
    }
  }, [user, showSignIn]);

  // Handler routes (like /handler/oauth-callback) must be handled by StackHandler
  if (isHandlerRoute) {
    return <HandlerRoutes />;
  }

  // Show sign-in page if requested
  if (showSignIn) {
    return <SignInPage />;
  }

  // Show landing page if user is not signed in
  if (showLanding) {
    return <LandingPage onGetStarted={handleGetStarted} />;
  }

  // Main app (user is signed in)
  return (
    <div className="flex flex-col h-screen bg-white">
      {/* Header Bar */}
      <div className="flex items-center justify-between px-4 py-2 border-b border-gray-200 bg-white flex-shrink-0">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2">
            <div className="text-2xl">
              📊
            </div>
            <span className="font-semibold text-gray-900">SheetGrid</span>
          </div>
        </div>
        <div className="flex items-center gap-3">
          <AuthButton onShowProfile={() => setShowUserProfile(true)} />
        </div>
      </div>

      {/* Main Content */}
      <div className="flex flex-1 overflow-hidden">
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
              {/* Chat Panel */}
              {activeChatId && (
                <ChatPanel 
                  chatId={activeChatId}
                  chatTitle={activeChatTitle}
                  onCreateNewChat={handleCreateNewChat}
                  onChatTitleChange={setActiveChatTitle}
                  onSelectChat={handleSelectChat}
                  onDeleteChat={handleDeleteChat}
                  allChats={allChats}
                />
              )}
            </div>
          </>
        )}
      </div>

      {/* User Profile Modal */}
      {showUserProfile && user && (
        <UserProfile onClose={() => setShowUserProfile(false)} />
      )}
    </div>
  );
}

export default App;
