// IndexedDB utility for persisting chat history and workbook data

const DB_NAME = 'sheetgrid_db';
const DB_VERSION = 2; // Increment version for schema changes
const STORES = {
  CHATS: 'chats',
  CHAT_HISTORY: 'chat_history',
  WORKBOOK_DATA: 'workbook_data',
};

interface DBOpener {
  db: IDBDatabase;
}

// Initialize database
export async function initDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onerror = () => {
      reject(new Error('Failed to open IndexedDB'));
    };

    request.onsuccess = () => {
      resolve(request.result);
    };

    request.onupgradeneeded = (event) => {
      const db = (event.target as IDBOpenDBRequest).result;
      const oldVersion = event.oldVersion;

      // Create chats store for managing multiple chat sessions
      if (!db.objectStoreNames.contains(STORES.CHATS)) {
        const chatsStore = db.createObjectStore(STORES.CHATS, { keyPath: 'id' });
        chatsStore.createIndex('updatedAt', 'updatedAt', { unique: false });
      }

      // Create chat history store
      if (!db.objectStoreNames.contains(STORES.CHAT_HISTORY)) {
        const chatStore = db.createObjectStore(STORES.CHAT_HISTORY, { keyPath: 'id', autoIncrement: true });
        chatStore.createIndex('timestamp', 'timestamp', { unique: false });
        chatStore.createIndex('chatId', 'chatId', { unique: false });
      } else if (oldVersion < 2) {
        // Migrate existing data: add chatId index
        const chatStore = event.target?.transaction?.objectStore(STORES.CHAT_HISTORY);
        if (chatStore && !chatStore.indexNames.contains('chatId')) {
          chatStore.createIndex('chatId', 'chatId', { unique: false });
        }
      }

      // Create workbook data store
      if (!db.objectStoreNames.contains(STORES.WORKBOOK_DATA)) {
        db.createObjectStore(STORES.WORKBOOK_DATA, { keyPath: 'id' });
      }
    };
  });
}

// Chat Management Operations
export interface Chat {
  id: string;
  title: string;
  createdAt: string;
  updatedAt: string;
}

export async function createChat(title?: string): Promise<Chat> {
  const chat: Chat = {
    id: `chat-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
    title: title || 'New Chat',
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };

  try {
    const db = await initDB();
    const transaction = db.transaction([STORES.CHATS], 'readwrite');
    const store = transaction.objectStore(STORES.CHATS);
    
    await new Promise<void>((resolve, reject) => {
      const request = store.add(chat);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
    
    return chat;
  } catch (error) {
    console.error('Error creating chat:', error);
    throw error;
  }
}

export async function getAllChats(): Promise<Chat[]> {
  try {
    const db = await initDB();
    const transaction = db.transaction([STORES.CHATS], 'readonly');
    const store = transaction.objectStore(STORES.CHATS);
    const index = store.index('updatedAt');

    return new Promise((resolve, reject) => {
      const request = index.getAll();
      request.onsuccess = () => {
        const chats = request.result.sort((a: Chat, b: Chat) => 
          new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime()
        );
        resolve(chats);
      };
      request.onerror = () => reject(request.error);
    });
  } catch (error) {
    console.error('Error loading chats:', error);
    return [];
  }
}

export async function updateChat(chatId: string, updates: Partial<Chat>): Promise<void> {
  try {
    const db = await initDB();
    const transaction = db.transaction([STORES.CHATS], 'readwrite');
    const store = transaction.objectStore(STORES.CHATS);
    
    const chat = await new Promise<Chat>((resolve, reject) => {
      const request = store.get(chatId);
      request.onsuccess = () => {
        if (request.result) {
          resolve(request.result);
        } else {
          reject(new Error('Chat not found'));
        }
      };
      request.onerror = () => reject(request.error);
    });

    const updatedChat = { ...chat, ...updates, updatedAt: new Date().toISOString() };
    
    await new Promise<void>((resolve, reject) => {
      const request = store.put(updatedChat);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  } catch (error) {
    console.error('Error updating chat:', error);
    throw error;
  }
}

export async function deleteChat(chatId: string): Promise<void> {
  try {
    const db = await initDB();
    
    // Delete chat record
    const chatsTransaction = db.transaction([STORES.CHATS], 'readwrite');
    const chatsStore = chatsTransaction.objectStore(STORES.CHATS);
    await new Promise<void>((resolve, reject) => {
      const request = chatsStore.delete(chatId);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });

    // Delete all messages for this chat
    const historyTransaction = db.transaction([STORES.CHAT_HISTORY], 'readwrite');
    const historyStore = historyTransaction.objectStore(STORES.CHAT_HISTORY);
    const chatIdIndex = historyStore.index('chatId');
    const messages = await new Promise<any[]>((resolve, reject) => {
      const request = chatIdIndex.getAll(chatId);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });

    for (const message of messages) {
      await new Promise<void>((resolve, reject) => {
        const request = historyStore.delete(message.id);
        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error);
      });
    }
  } catch (error) {
    console.error('Error deleting chat:', error);
    throw error;
  }
}

// Chat History Operations (scoped by chatId)
export async function saveAllChatMessages(chatId: string, messages: any[]): Promise<void> {
  try {
    const db = await initDB();
    const transaction = db.transaction([STORES.CHAT_HISTORY], 'readwrite');
    const store = transaction.objectStore(STORES.CHAT_HISTORY);
    const chatIdIndex = store.index('chatId');
    
    // Delete existing messages for this chat
    const existingMessages = await new Promise<any[]>((resolve, reject) => {
      const request = chatIdIndex.getAll(chatId);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });

    for (const msg of existingMessages) {
      await new Promise<void>((resolve, reject) => {
        const request = store.delete(msg.id);
        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error);
      });
    }

    // Add all new messages
    for (const message of messages) {
      await new Promise<void>((resolve, reject) => {
        const request = store.add({
          ...message,
          chatId: chatId,
          timestamp: message.timestamp ? (message.timestamp instanceof Date ? message.timestamp.toISOString() : message.timestamp) : new Date().toISOString(),
        });
        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error);
      });
    }

    // Update chat's updatedAt timestamp
    await updateChat(chatId, {});
  } catch (error) {
    console.error('Error saving all chat messages:', error);
  }
}

export async function loadChatMessages(chatId: string): Promise<any[]> {
  try {
    const db = await initDB();
    const transaction = db.transaction([STORES.CHAT_HISTORY], 'readonly');
    const store = transaction.objectStore(STORES.CHAT_HISTORY);
    const chatIdIndex = store.index('chatId');

    return new Promise((resolve, reject) => {
      const request = chatIdIndex.getAll(chatId);
      request.onsuccess = () => {
        const messages = request.result.map((msg: any) => ({
          ...msg,
          timestamp: msg.timestamp ? new Date(msg.timestamp) : new Date(),
        }));
        // Sort by timestamp
        messages.sort((a: any, b: any) => 
          a.timestamp.getTime() - b.timestamp.getTime()
        );
        resolve(messages);
      };
      request.onerror = () => reject(request.error);
    });
  } catch (error) {
    console.error('Error loading chat messages:', error);
    return [];
  }
}

export async function clearChatHistory(chatId: string): Promise<void> {
  try {
    const db = await initDB();
    const transaction = db.transaction([STORES.CHAT_HISTORY], 'readwrite');
    const store = transaction.objectStore(STORES.CHAT_HISTORY);
    const chatIdIndex = store.index('chatId');
    
    const messages = await new Promise<any[]>((resolve, reject) => {
      const request = chatIdIndex.getAll(chatId);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });

    for (const message of messages) {
      await new Promise<void>((resolve, reject) => {
        const request = store.delete(message.id);
        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error);
      });
    }
  } catch (error) {
    console.error('Error clearing chat history:', error);
  }
}

// Workbook Data Operations
export async function saveWorkbookData(data: any): Promise<void> {
  try {
    const db = await initDB();
    const transaction = db.transaction([STORES.WORKBOOK_DATA], 'readwrite');
    const store = transaction.objectStore(STORES.WORKBOOK_DATA);
    
    await new Promise<void>((resolve, reject) => {
      const request = store.put({
        id: 'current',
        data: data,
        savedAt: new Date().toISOString(),
      });
      request.onsuccess = () => {
        console.log('Workbook data saved to IndexedDB');
        resolve();
      };
      request.onerror = () => reject(request.error);
    });
  } catch (error) {
    console.error('Error saving workbook data:', error);
  }
}

export async function loadWorkbookData(): Promise<any | null> {
  try {
    const db = await initDB();
    const transaction = db.transaction([STORES.WORKBOOK_DATA], 'readonly');
    const store = transaction.objectStore(STORES.WORKBOOK_DATA);
    
    return new Promise((resolve, reject) => {
      const request = store.get('current');
      request.onsuccess = () => {
        const result = request.result;
        if (result && result.data) {
          console.log('Workbook data loaded from IndexedDB');
          resolve(result.data);
        } else {
          resolve(null);
        }
      };
      request.onerror = () => reject(request.error);
    });
  } catch (error) {
    console.error('Error loading workbook data:', error);
    return null;
  }
}

export async function clearWorkbookData(): Promise<void> {
  try {
    const db = await initDB();
    const transaction = db.transaction([STORES.WORKBOOK_DATA], 'readwrite');
    const store = transaction.objectStore(STORES.WORKBOOK_DATA);
    
    await new Promise<void>((resolve, reject) => {
      const request = store.delete('current');
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  } catch (error) {
    console.error('Error clearing workbook data:', error);
  }
}

