import { StackProvider, StackClientApp } from '@stackframe/stack';
import { ReactNode } from 'react';

const projectId = import.meta.env.VITE_STACK_PROJECT_ID;
const publishableClientKey = import.meta.env.VITE_STACK_PUBLISHABLE_CLIENT_KEY;

if (!projectId || !publishableClientKey) {
  console.warn('Stack Auth environment variables not set. Auth features may not work.');
}

const stackApp = new StackClientApp({
  projectId: projectId || '',
  publishableClientKey: publishableClientKey || '',
  tokenStore: 'memory',
  redirectMethod: 'window',
  urls: {
    signIn: '/sign-in',
    afterSignIn: '/',
    afterSignUp: '/',
    afterSignOut: '/',
  },
});

export function StackAuthProvider({ children }: { children: ReactNode }) {
  return (
    <StackProvider app={stackApp}>
      {children}
    </StackProvider>
  );
}

