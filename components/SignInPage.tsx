import { SignIn } from '@stackframe/react';
import { useEffect } from 'react';

export default function SignInPage() {
  useEffect(() => {
    // Ensure background is white
    document.body.style.background = '#ffffff';
    document.documentElement.style.background = '#ffffff';
    return () => {
      // Reset on unmount if needed
    };
  }, []);

  return (
    <div className="flex items-center justify-center min-h-screen bg-white" style={{ background: '#ffffff' }}>
      <div className="w-full max-w-md p-8">
        <div className="mb-8 text-center">
          <div className="text-5xl mb-4">📊</div>
          <h1 className="text-3xl font-bold text-gray-900 mb-2">SheetGrid</h1>
          <p className="text-gray-600">Sign in to continue</p>
        </div>
        <div style={{ background: '#ffffff' }}>
          <SignIn />
        </div>
      </div>
    </div>
  );
}

