import { useState, useEffect } from 'react';
import { useUser, useStackApp } from '@stackframe/react';
import { AccountSettings } from '@stackframe/react';

export default function UserProfile({ onClose }: { onClose: () => void }) {
  const user = useUser();
  const [openRouterKey, setOpenRouterKey] = useState('');
  const [univerMcpKey, setUniverMcpKey] = useState('');
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    // Load user's API keys from NeonDB
    const loadKeys = async () => {
      if (!user) return;
      
      try {
        const { getUserApiKeys } = await import('@/src/utils/neondb');
        const keys = await getUserApiKeys(user.id);
        
        if (keys) {
          setOpenRouterKey(keys.openRouterKey || '');
          setUniverMcpKey(keys.univerMcpKey || '');
        }
      } catch (error) {
        console.error('Error loading API keys:', error);
      } finally {
        setLoading(false);
      }
    };

    loadKeys();
  }, [user]);

  const handleSaveKeys = async () => {
    if (!user) return;
    
    setSaving(true);
    try {
      const { saveUserApiKeys } = await import('@/src/utils/neondb');
      await saveUserApiKeys(user.id, {
        openRouterKey: openRouterKey.trim() || undefined,
        univerMcpKey: univerMcpKey.trim() || undefined,
      });
      
      // Update localStorage/cache for immediate use
      if (openRouterKey.trim()) {
        localStorage.setItem(`openrouter_key_${user.id}`, openRouterKey.trim());
      } else {
        localStorage.removeItem(`openrouter_key_${user.id}`);
      }
      if (univerMcpKey.trim()) {
        localStorage.setItem(`univer_mcp_key_${user.id}`, univerMcpKey.trim());
      } else {
        localStorage.removeItem(`univer_mcp_key_${user.id}`);
      }
      
      // Trigger a reload to apply new keys
      alert('API keys saved successfully! The page will reload to apply the changes.');
      window.location.reload();
    } catch (error) {
      console.error('Error saving API keys:', error);
      alert('Failed to save API keys. Please try again.');
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 z-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-lg shadow-xl max-w-4xl w-full max-h-[90vh] overflow-y-auto">
        <div className="sticky top-0 bg-white border-b border-gray-200 px-6 py-4 flex items-center justify-between">
          <h2 className="text-2xl font-bold text-gray-900">User Profile & Settings</h2>
          <button
            onClick={onClose}
            className="text-gray-400 hover:text-gray-600 transition-colors"
          >
            <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>

        <div className="p-6 space-y-8">
          {/* Stack Auth Account Settings */}
          <div>
            <h3 className="text-lg font-semibold text-gray-900 mb-4">Account Settings</h3>
            <div className="border border-gray-200 rounded-lg p-4">
              <AccountSettings />
            </div>
          </div>

          {/* API Keys Section */}
          <div>
            <h3 className="text-lg font-semibold text-gray-900 mb-4">API Keys</h3>
            <div className="space-y-4">
              {/* OpenRouter API Key */}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  OpenRouter API Key
                </label>
                <input
                  type="password"
                  value={openRouterKey}
                  onChange={(e) => setOpenRouterKey(e.target.value)}
                  placeholder="sk-or-v1-..."
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  disabled={loading || saving}
                />
                <p className="mt-1 text-xs text-gray-500">
                  Your OpenRouter API key for AI chat functionality
                </p>
              </div>

              {/* Univer MCP API Key */}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Univer MCP API Key
                </label>
                <input
                  type="password"
                  value={univerMcpKey}
                  onChange={(e) => setUniverMcpKey(e.target.value)}
                  placeholder="sk-..."
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  disabled={loading || saving}
                />
                <p className="mt-1 text-xs text-gray-500">
                  Your Univer MCP API key for spreadsheet functionality
                </p>
              </div>

              <button
                onClick={handleSaveKeys}
                disabled={loading || saving}
                className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {saving ? 'Saving...' : 'Save API Keys'}
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

