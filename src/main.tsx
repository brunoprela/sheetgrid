import React, { Suspense } from 'react';
import ReactDOM from 'react-dom/client';
import { BrowserRouter } from 'react-router-dom';
import App from './App';
import { StackAuthProvider } from './providers/StackAuthProvider';
import './index.css';

function main() {
  ReactDOM.createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <Suspense fallback={null}>
        <BrowserRouter>
          <StackAuthProvider>
            <App />
          </StackAuthProvider>
        </BrowserRouter>
      </Suspense>
    </React.StrictMode>
  );
}

main();
