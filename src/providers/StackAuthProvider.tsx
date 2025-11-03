import { StackProvider, StackClientApp, StackTheme } from '@stackframe/react';
import { ReactNode } from 'react';

const projectId = import.meta.env.VITE_STACK_PROJECT_ID;
const publishableClientKey = import.meta.env.VITE_STACK_PUBLISHABLE_CLIENT_KEY;

if (!projectId || !publishableClientKey) {
  console.warn('Neon Auth environment variables not set. Auth features may not work.');
}

export const stackClientApp = new StackClientApp({
  projectId: projectId || '',
  publishableClientKey: publishableClientKey || '',
  tokenStore: 'cookie',
  redirectMethod: 'window',
});

export function StackAuthProvider({ children }: { children: ReactNode }) {
  return (
    <StackProvider app={stackClientApp}>
      <StackTheme theme="light">
        {children}
      </StackTheme>
    </StackProvider>
  );
}

