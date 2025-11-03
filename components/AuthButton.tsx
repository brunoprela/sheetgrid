import { useState, useRef, useEffect } from 'react';
import { useStackApp, useUser } from '@stackframe/react';
import { useNavigate } from 'react-router-dom';

export default function AuthButton() {
  const stackApp = useStackApp();
  const user = useUser();
  const navigate = useNavigate();
  const [showDropdown, setShowDropdown] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  // Close dropdown when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setShowDropdown(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  if (user) {
    return (
      <div className="relative" ref={dropdownRef}>
        <button
          onClick={() => setShowDropdown(!showDropdown)}
          className="flex items-center gap-2 px-3 py-1.5 rounded-lg bg-gray-100 hover:bg-gray-200 transition-colors"
        >
          <div className="w-8 h-8 bg-purple-600 rounded-full flex items-center justify-center text-white text-xs font-semibold">
            {(user.displayName || user.primaryEmail || 'U').charAt(0).toUpperCase()}
          </div>
          <span className="text-sm font-medium text-gray-900">{user.displayName || user.primaryEmail || 'User'}</span>
        </button>

        {showDropdown && (
          <div className="absolute right-0 mt-2 w-64 bg-white rounded-lg shadow-lg border border-gray-200 overflow-hidden z-50">
            {/* User Info */}
            <div className="px-4 py-3 border-b border-gray-200">
              <div className="font-medium text-gray-900">{user.displayName || user.primaryEmail}</div>
              <div className="text-sm text-gray-500 mt-0.5">{user.primaryEmail}</div>
            </div>

            {/* Account Settings */}
            <button
              onClick={() => {
                setShowDropdown(false);
                navigate('/profile');
              }}
              className="w-full px-4 py-2 text-left text-sm text-gray-900 hover:bg-gray-50 transition-colors"
            >
              Account Settings
            </button>

            {/* Sign Out */}
            <button
              onClick={async () => {
                setShowDropdown(false);
                await stackApp.redirectToSignOut();
              }}
              className="w-full px-4 py-2 text-left text-sm text-red-600 hover:bg-red-50 transition-colors border-t border-gray-200"
            >
              Sign Out
            </button>
          </div>
        )}
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

