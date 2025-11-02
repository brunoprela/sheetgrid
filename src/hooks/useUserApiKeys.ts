import { useState, useEffect } from 'react';
import { useUser } from '@stackframe/react';
import { getUserApiKeys } from '@/src/utils/neondb';

export function useUserApiKeys() {
  const user = useUser();
  const [openRouterKey, setOpenRouterKey] = useState<string | null>(null);
  const [univerMcpKey, setUniverMcpKey] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    const loadKeys = async () => {
      if (!user) {
        setOpenRouterKey(null);
        setUniverMcpKey(null);
        setLoading(false);
        return;
      }

      setLoading(true);

      // Check localStorage cache first for immediate response
      try {
        const cachedOpenRouter = localStorage.getItem(`openrouter_key_${user.id}`);
        const cachedUniverMcp = localStorage.getItem(`univer_mcp_key_${user.id}`);

        if (cachedOpenRouter || cachedUniverMcp) {
          setOpenRouterKey(cachedOpenRouter);
          setUniverMcpKey(cachedUniverMcp);
          setLoading(false);
        }
      } catch (error) {
        console.error('Error reading from localStorage:', error);
      }
      
      // Load from NeonDB in background (non-blocking)
      try {
        const keys = await getUserApiKeys(user.id);
        if (keys) {
          const orKey = keys.openRouterKey || null;
          const univerKey = keys.univerMcpKey || null;
          
          setOpenRouterKey(orKey);
          setUniverMcpKey(univerKey);
          
          // Update cache
          try {
            if (orKey) {
              localStorage.setItem(`openrouter_key_${user.id}`, orKey);
            } else {
              localStorage.removeItem(`openrouter_key_${user.id}`);
            }
            if (univerKey) {
              localStorage.setItem(`univer_mcp_key_${user.id}`, univerKey);
            } else {
              localStorage.removeItem(`univer_mcp_key_${user.id}`);
            }
          } catch (storageError) {
            console.error('Error updating localStorage:', storageError);
          }
        }
      } catch (error) {
        console.error('Error loading API keys:', error);
        // Don't block the app if database fails - fallback to env vars
      } finally {
        setLoading(false);
      }
    };

    loadKeys();
  }, [user]);

  return {
    openRouterKey,
    univerMcpKey,
    loading,
  };
}

