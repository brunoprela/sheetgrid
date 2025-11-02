import { useStackApp, useUser } from '@stackframe/react';

export default function AuthButton() {
  const stackApp = useStackApp();
  const user = useUser();

  if (user) {
    return (
      <div className="flex items-center gap-3">
        <span className="text-sm text-gray-600">
          {user.displayName || user.primaryEmail || 'User'}
        </span>
        <button
          onClick={async () => {
            await user.signOut();
            window.location.href = '/';
          }}
          className="px-3 py-1.5 text-sm text-gray-600 hover:text-gray-900 transition-colors border border-gray-300 rounded-md hover:border-gray-400"
        >
          Sign Out
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

