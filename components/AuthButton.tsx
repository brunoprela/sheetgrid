import { useStackApp, useUser } from '@stackframe/react';

export default function AuthButton({ onShowProfile }: { onShowProfile: () => void }) {
  const stackApp = useStackApp();
  const user = useUser();

  if (user) {
    return (
      <div className="flex items-center gap-3">
        <button
          onClick={onShowProfile}
          className="flex items-center gap-2 px-3 py-1.5 text-sm text-gray-600 hover:text-gray-900 transition-colors"
        >
          <div className="w-8 h-8 bg-blue-600 rounded-full flex items-center justify-center text-white text-xs font-semibold">
            {(user.displayName || user.primaryEmail || 'U').charAt(0).toUpperCase()}
          </div>
          <span>{user.displayName || user.primaryEmail || 'User'}</span>
        </button>
      </div>
    );
  }

  return (
    <button
      onClick={async () => {
        await stackApp.redirectToSignIn();
      }}
      className="px-3 py-1.5 text-sm text-white bg-blue-600 hover:bg-blue-700 rounded-md transition-colors"
    >
      Sign In
    </button>
  );
}

