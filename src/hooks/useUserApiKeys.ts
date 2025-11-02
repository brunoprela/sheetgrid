import { useState, useEffect } from 'react';
import { useUser } from '@stackframe/react';
import { getUserApiKeys } from '@/src/utils/neondb';

export function useUserApiKeys() {
  const user = useUser();
  const [openRouterKey, setOpenRouterKey] = useState<string | null>(null);
  const [univerMcpKey, setUniverMcpKey] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const loadKeys = async () => {
      if (!user) {
        setOpenRouterKey(null);
        setUniverMcpKey(null);
        setLoading(false);
        return;
      }

      // Check localStorage cache first
      const cachedOpenRouter = localStorage.getItem(`openrouter_key_${user.id}`);
      const cachedUniverMcp = localStorage.getItem(`univer_mcp_key_${user.id}`);

      if (cachedOpenRouter || cachedUniverMcp) {
        setOpenRouterKey(cachedOpenRouter);
        setUniverMcpKey(cachedUniverMcp);
        setLoading(false);
      }

      // Load from NeonDB
      try {
        const keys = await getUserApiKeys(user.id);
        if (keys) {
          const orKey = keys.openRouterKey || null;
          const univerKey = keys.univerMcpKey || null;
          
          setOpenRouterKey(orKey);
          setUniverMcpKey(univerKey);
          
          // Update cache
          if (orKey) {
            localStorage.setItem(`openrouter_key_${user.id}`, orKey);
          }
          if (univerKey) {
            localStorage.setItem(`univer_mcp_key_${user.id}`, univerKey);
          }
        }
      } catch (error) {
        console.error('Error loading API keys:', error);
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

