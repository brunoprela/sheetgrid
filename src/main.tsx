import React, { Suspense } from 'react';
import ReactDOM from 'react-dom/client';
import { BrowserRouter } from 'react-router-dom';
import { ErrorBoundary } from './components/ErrorBoundary';
import App from './App';
import { StackAuthProvider } from './providers/StackAuthProvider';
import './index.css';

function main() {
  ReactDOM.createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <ErrorBoundary>
        <Suspense fallback={null}>
          <BrowserRouter>
            <StackAuthProvider>
              <App />
            </StackAuthProvider>
          </BrowserRouter>
        </Suspense>
      </ErrorBoundary>
    </React.StrictMode>
  );
}

main();
