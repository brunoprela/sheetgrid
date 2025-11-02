import { useState, useEffect, useRef } from 'react';
import { getAllChats, createChat, deleteChat, updateChat, type Chat } from '../src/utils/indexeddb';

interface ChatListProps {
  activeChatId: string | null;
  onSelectChat: (chatId: string) => void;
  onChatCreated: (chat: Chat) => void;
}

export default function ChatList({ activeChatId, onSelectChat, onChatCreated }: ChatListProps) {
  const [chats, setChats] = useState<Chat[]>([]);
  const [isCreating, setIsCreating] = useState(false);
  const [editingChatId, setEditingChatId] = useState<string | null>(null);
  const [editingTitle, setEditingTitle] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    loadChats();
  }, []);

  const loadChats = async () => {
    const allChats = await getAllChats();
    setChats(allChats);
    
    // If no chats exist and no active chat, create one
    if (allChats.length === 0 && !activeChatId) {
      handleCreateChat();
    }
  };

  const handleCreateChat = async () => {
    setIsCreating(true);
    try {
      const newChat = await createChat();
      await loadChats(); // Reload to get sorted list
      onSelectChat(newChat.id);
      onChatCreated(newChat);
    } catch (error) {
      console.error('Error creating chat:', error);
    } finally {
      setIsCreating(false);
    }
  };

  const handleDeleteChat = async (chatId: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (window.confirm('Are you sure you want to delete this chat?')) {
      try {
        await deleteChat(chatId);
        await loadChats(); // Reload to get updated list
        if (activeChatId === chatId) {
          // Select another chat or create a new one
          const updatedChats = await getAllChats();
          if (updatedChats.length > 0) {
            onSelectChat(updatedChats[0].id);
          } else {
            handleCreateChat();
          }
        }
      } catch (error) {
        console.error('Error deleting chat:', error);
      }
    }
  };

  const handleStartEdit = (chat: Chat, e: React.MouseEvent) => {
    e.stopPropagation();
    setEditingChatId(chat.id);
    setEditingTitle(chat.title);
    setTimeout(() => inputRef.current?.focus(), 0);
  };

  const handleSaveEdit = async (chatId: string) => {
    try {
      await updateChat(chatId, { title: editingTitle });
      setChats(prev => prev.map(chat => 
        chat.id === chatId ? { ...chat, title: editingTitle, updatedAt: new Date().toISOString() } : chat
      ));
      setEditingChatId(null);
      setEditingTitle('');
    } catch (error) {
      console.error('Error updating chat title:', error);
    }
  };

  const handleCancelEdit = () => {
    setEditingChatId(null);
    setEditingTitle('');
  };

  return (
    <div className="h-full flex flex-col bg-gray-50 border-r border-gray-200">
      {/* Header */}
      <div className="p-3 border-b border-gray-200 flex items-center justify-between">
        <h2 className="text-sm font-semibold text-gray-700">Chats</h2>
        <button
          onClick={handleCreateChat}
          disabled={isCreating}
          className="p-1.5 rounded hover:bg-gray-200 transition-colors disabled:opacity-50"
          title="New Chat"
        >
          <svg className="w-4 h-4 text-gray-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" />
          </svg>
        </button>
      </div>

      {/* Chat List */}
      <div className="flex-1 overflow-y-auto">
        {chats.length === 0 ? (
          <div className="p-4 text-center text-sm text-gray-500">
            No chats yet
          </div>
        ) : (
          <div className="py-2">
            {chats.map((chat) => (
              <div
                key={chat.id}
                onClick={() => onSelectChat(chat.id)}
                className={`px-3 py-2 mx-2 my-1 rounded cursor-pointer group transition-colors ${
                  activeChatId === chat.id
                    ? 'bg-blue-100 text-blue-900'
                    : 'hover:bg-gray-100 text-gray-700'
                }`}
              >
                {editingChatId === chat.id ? (
                  <input
                    ref={inputRef}
                    value={editingTitle}
                    onChange={(e) => setEditingTitle(e.target.value)}
                    onBlur={() => handleSaveEdit(chat.id)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        handleSaveEdit(chat.id);
                      } else if (e.key === 'Escape') {
                        handleCancelEdit();
                      }
                    }}
                    className="w-full px-1 py-0.5 text-sm bg-white border border-blue-300 rounded"
                    onClick={(e) => e.stopPropagation()}
                  />
                ) : (
                  <div className="flex items-center justify-between">
                    <span className="text-sm truncate flex-1">{chat.title}</span>
                    <div className="flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                      <button
                        onClick={(e) => handleStartEdit(chat, e)}
                        className="p-1 hover:bg-gray-200 rounded"
                        title="Rename"
                      >
                        <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
                        </svg>
                      </button>
                      <button
                        onClick={(e) => handleDeleteChat(chat.id, e)}
                        className="p-1 hover:bg-red-100 rounded text-red-600"
                        title="Delete"
                      >
                        <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                        </svg>
                      </button>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

