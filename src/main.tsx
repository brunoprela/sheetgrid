import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import { StackAuthProvider } from './providers/StackAuthProvider';
import './index.css';

function main() {
  ReactDOM.createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <StackAuthProvider>
        <App />
      </StackAuthProvider>
    </React.StrictMode>
  );
}

main();
