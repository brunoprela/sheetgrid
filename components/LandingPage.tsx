import { useStackApp } from '@stackframe/stack';

export default function LandingPage({ onGetStarted }: { onGetStarted: () => void }) {
  const stackApp = useStackApp();
  return (
    <div className="flex items-center justify-center min-h-screen bg-white overflow-y-auto overflow-x-hidden">
      <div className="max-w-4xl mx-auto px-8 py-16 text-center w-full">
        {/* Logo/Icon */}
        <div className="mb-8 flex justify-center">
          <div className="text-7xl">
            📊
          </div>
        </div>

        {/* Main Title */}
        <h1 className="text-5xl font-bold text-gray-900 mb-4">
          SheetGrid
        </h1>
        
        {/* Subtitle */}
        <p className="text-2xl text-gray-600 mb-8">
          AI-Powered Excel Editor
        </p>

        {/* Description */}
        <div className="max-w-2xl mx-auto mb-12">
          <p className="text-lg text-gray-700 leading-relaxed mb-6">
            Upload Excel files, edit them in a beautiful interface, and interact with your data using natural language AI commands.
          </p>
          
          <p className="text-base text-gray-600 leading-relaxed">
            SheetGrid combines the power of professional spreadsheet editing with AI assistance, 
            allowing you to analyze, modify, and work with your data through conversational interactions.
          </p>
        </div>

        {/* Development Notice */}
        <div className="bg-blue-50 border border-blue-200 rounded-lg p-6 mb-12 max-w-2xl mx-auto">
          <p className="text-sm text-blue-900 mb-3 font-semibold">
            🚧 Development Project
          </p>
          <p className="text-sm text-blue-800 mb-4">
            SheetGrid is currently in active development. We're continuously improving features and fixing bugs.
          </p>
          <a
            href="https://github.com/brunoprela/sheetgrid"
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center gap-2 text-sm text-blue-600 hover:text-blue-700 font-medium transition-colors"
          >
            <svg className="w-5 h-5" fill="currentColor" viewBox="0 0 24 24" aria-hidden="true">
              <path fillRule="evenodd" d="M12 2C6.477 2 2 6.484 2 12.017c0 4.425 2.865 8.18 6.839 9.504.5.092.682-.217.682-.483 0-.237-.008-.868-.013-1.703-2.782.605-3.369-1.343-3.369-1.343-.454-1.158-1.11-1.466-1.11-1.466-.908-.62.069-.608.069-.608 1.003.07 1.531 1.032 1.531 1.032.892 1.532 2.341 1.088 2.91.832.092-.647.35-1.088.636-1.338-2.22-.253-4.555-1.113-4.555-4.951 0-1.093.39-1.988 1.029-2.688-.103-.253-.446-1.272.098-2.65 0 0 .84-.27 2.75 1.026A9.564 9.564 0 0112 6.844c.85.004 1.705.115 2.504.337 1.909-1.296 2.747-1.027 2.747-1.027.546 1.379.202 2.398.1 2.651.64.7 1.028 1.595 1.028 2.688 0 3.848-2.339 4.695-4.566 4.943.359.309.678.92.678 1.855 0 1.338-.012 2.419-.012 2.747 0 .268.18.58.688.482A10.019 10.019 0 0022 12.017C22 6.484 17.522 2 12 2z" clipRule="evenodd" />
            </svg>
            View on GitHub
          </a>
        </div>

        {/* Features Grid */}
        <div className="grid md:grid-cols-3 gap-6 mb-12 max-w-3xl mx-auto">
          <div className="bg-gray-50 rounded-lg p-6">
            <div className="text-3xl mb-3">📊</div>
            <h3 className="font-semibold text-gray-900 mb-2">Excel Upload</h3>
            <p className="text-sm text-gray-600">
              Upload and edit .xlsx or .xls files with full spreadsheet functionality
            </p>
          </div>
          
          <div className="bg-gray-50 rounded-lg p-6">
            <div className="text-3xl mb-3">🤖</div>
            <h3 className="font-semibold text-gray-900 mb-2">AI Assistant</h3>
            <p className="text-sm text-gray-600">
              Natural language commands to analyze, modify, and work with your data
            </p>
          </div>
          
          <div className="bg-gray-50 rounded-lg p-6">
            <div className="text-3xl mb-3">⚡</div>
            <h3 className="font-semibold text-gray-900 mb-2">Powerful Editing</h3>
            <p className="text-sm text-gray-600">
              Full-featured spreadsheet editing with formatting, formulas, and more
            </p>
          </div>
        </div>

        {/* Get Started Button */}
        <button
          onClick={onGetStarted}
          className="inline-flex items-center gap-2 px-8 py-4 bg-blue-600 text-white font-semibold rounded-lg hover:bg-blue-700 transition-colors shadow-lg hover:shadow-xl transform hover:scale-105 transition-transform"
        >
          Sign In to Get Started
          <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 7l5 5m0 0l-5 5m5-5H6" />
          </svg>
        </button>
      </div>
    </div>
  );
}

