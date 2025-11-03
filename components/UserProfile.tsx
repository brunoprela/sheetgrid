import { useState, useEffect } from 'react';
import { useUser } from '@stackframe/react';
import { AccountSettings } from '@stackframe/react';
import { useNavigate } from 'react-router-dom';

export default function ProfilePage() {
  const user = useUser();
  const navigate = useNavigate();
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

  // Inject global CSS first to fix tooltips
  useEffect(() => {
    const styleId = 'global-account-settings-fix';
    if (!document.getElementById(styleId)) {
      const style = document.createElement('style');
      style.id = styleId;
      style.textContent = `
        /* Aggressive fix for all dark backgrounds in the profile page */
        body [style*="background-color: rgb(0"],
        body [style*="background-color:rgba(0"],
        body [style*="background: rgb(0"],
        body [style*="background:rgba(0"],
        body [style*="background-color:#000"],
        body [style*="background:#000"],
        body [style*="background-color:#0a0a0a"],
        body [style*="background:#0a0a0a"] {
          background-color: white !important;
          background: white !important;
          color: #000000 !important;
        }
        
        body [style*="background-color: rgb(0"] *,
        body [style*="background-color:rgba(0"] *,
        body [style*="background: rgb(0"] *,
        body [style*="background:rgba(0"] *,
        body [style*="background-color:#000"] *,
        body [style*="background:#000"] *,
        body [style*="background-color:#0a0a0a"] *,
        body [style*="background:#0a0a0a"] * {
          color: #000000 !important;
        }
      `;
      document.head.appendChild(style);
    }

    return () => {
      const existingStyle = document.getElementById(styleId);
      if (existingStyle) {
        existingStyle.remove();
      }
    };
  }, []);

  // Fix black on black text in AccountSettings with MutationObserver
  useEffect(() => {
    const fixAccountSettingsStyles = () => {
      // Find all elements with inline background styles
      const container = document.getElementById('account-settings-container');
      if (!container) return;

      // Use a more aggressive approach: query all elements and check their computed styles
      const allElements = container.querySelectorAll('*');
      allElements.forEach((element: Element) => {
        const htmlElement = element as HTMLElement;
        const computedStyle = window.getComputedStyle(htmlElement);
        const backgroundColor = computedStyle.backgroundColor;

        // Check if background is dark (not white/transparent)
        if (backgroundColor && backgroundColor !== 'rgba(0, 0, 0, 0)' && backgroundColor !== 'transparent') {
          // Parse RGB to check if it's dark
          const rgbMatch = backgroundColor.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
          if (rgbMatch) {
            const r = parseInt(rgbMatch[1]);
            const g = parseInt(rgbMatch[2]);
            const b = parseInt(rgbMatch[3]);
            // If it's a dark color (sum is less than 400), make it white
            if (r + g + b < 400) {
              htmlElement.style.setProperty('background-color', 'white', 'important');
              htmlElement.style.setProperty('background', 'white', 'important');
              htmlElement.style.setProperty('color', '#000000', 'important');
            }
          }
        }
      });

      // Also fix any elements with style attributes containing dark colors
      const elementsWithInlineStyles = container.querySelectorAll('[style*="background"]');
      elementsWithInlineStyles.forEach((element: Element) => {
        const htmlElement = element as HTMLElement;
        const inlineStyle = htmlElement.getAttribute('style') || '';

        // Check for dark colors in inline style
        if (inlineStyle.includes('rgba(0') || inlineStyle.includes('rgb(0') ||
          inlineStyle.includes('#000') || inlineStyle.includes('#0a0a0a')) {
          htmlElement.style.setProperty('background-color', 'white', 'important');
          htmlElement.style.setProperty('background', 'white', 'important');
          htmlElement.style.setProperty('color', '#000000', 'important');
        }
      });

      // Fix tooltips specifically - they might be outside the AccountSettings container
      // Use very aggressive selector to catch any tooltip-like element
      const tooltips = document.querySelectorAll('[role="tooltip"], .tooltip, [class*="Tooltip"], [class*="Popover"], [class*="tooltip"], [class*="popover"], [data-radix-portal], [class*="radix"]');
      tooltips.forEach((tooltip: Element) => {
        const htmlElement = tooltip as HTMLElement;
        const computedStyle = window.getComputedStyle(htmlElement);
        const backgroundColor = computedStyle.backgroundColor;

        // Check if background is dark
        if (backgroundColor && backgroundColor !== 'rgba(0, 0, 0, 0)' && backgroundColor !== 'transparent') {
          const rgbMatch = backgroundColor.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
          if (rgbMatch) {
            const r = parseInt(rgbMatch[1]);
            const g = parseInt(rgbMatch[2]);
            const b = parseInt(rgbMatch[3]);
            if (r + g + b < 400) {
              htmlElement.style.setProperty('background-color', 'white', 'important');
              htmlElement.style.setProperty('background', 'white', 'important');
              htmlElement.style.setProperty('color', '#000000', 'important');
              // Also fix all children
              const children = htmlElement.querySelectorAll('*');
              children.forEach((child: Element) => {
                (child as HTMLElement).style.setProperty('color', '#000000', 'important');
              });
            }
          }
        }

        // Also check inline styles directly
        const inlineStyle = htmlElement.getAttribute('style') || '';
        if (inlineStyle.includes('rgba(0') || inlineStyle.includes('rgb(0') ||
          inlineStyle.includes('#000') || inlineStyle.includes('#0a0a0a')) {
          htmlElement.style.setProperty('background-color', 'white', 'important');
          htmlElement.style.setProperty('background', 'white', 'important');
          htmlElement.style.setProperty('color', '#000000', 'important');
          const children = htmlElement.querySelectorAll('*');
          children.forEach((child: Element) => {
            (child as HTMLElement).style.setProperty('color', '#000000', 'important');
          });
        }
      });
    };

    // Run immediately and on mutations
    const observer = new MutationObserver(() => {
      fixAccountSettingsStyles();
    });

    // Observe the container
    const container = document.getElementById('account-settings-container');
    if (container) {
      observer.observe(container, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ['style', 'class']
      });
    }

    // Also observe body for tooltips that might be rendered outside the container
    const bodyObserver = new MutationObserver(() => {
      fixAccountSettingsStyles();
    });

    bodyObserver.observe(document.body, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ['style', 'class']
    });

    // Initial fix with multiple attempts
    setTimeout(fixAccountSettingsStyles, 100);
    setTimeout(fixAccountSettingsStyles, 500);
    setTimeout(fixAccountSettingsStyles, 1000);
    setTimeout(fixAccountSettingsStyles, 2000);

    // Also use interval as a fallback to catch any tooltips that appear
    const intervalId = setInterval(fixAccountSettingsStyles, 500);

    return () => {
      observer.disconnect();
      bodyObserver.disconnect();
      clearInterval(intervalId);
    };
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
    <div className="flex flex-col h-screen bg-white">
      {/* Header Bar */}
      <div className="flex items-center justify-between px-6 py-4 border-b border-gray-200 bg-white flex-shrink-0">
        <div className="flex items-center gap-3">
          <button
            onClick={() => navigate('/')}
            className="text-gray-600 hover:text-gray-900 transition-colors"
          >
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10 19l-7-7m0 0l7-7m-7 7h18" />
            </svg>
          </button>
          <h1 className="text-xl font-semibold text-gray-900">User Profile & Settings</h1>
        </div>
      </div>

      {/* Main Content */}
      <div className="flex-1 overflow-y-auto bg-white">
        <div className="max-w-4xl mx-auto px-6 py-8 space-y-8">
          {/* Stack Auth Account Settings */}
          <div>
            <h2 className="text-lg font-semibold text-gray-900 mb-4">Account Settings</h2>
            <div className="border border-gray-200 rounded-lg p-4 bg-white" id="account-settings-container">
              <AccountSettings />
            </div>
          </div>

          {/* API Keys Section */}
          <div>
            <h2 className="text-lg font-semibold text-gray-900 mb-4">API Keys</h2>
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
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
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
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
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

