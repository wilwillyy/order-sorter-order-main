import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { loadEnv } from 'vite';

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '');
  const getFirebaseValue = (key: string) => env[`firebase_${key.toLowerCase()}`] || env[`FIREBASE_${key}`] || env[`VITE_FIREBASE_${key}`] || '';

  return {
    plugins: [react()],
    define: {
      __FIREBASE_CONFIG__: JSON.stringify({
        apiKey: getFirebaseValue('API_KEY'),
        authDomain: getFirebaseValue('AUTH_DOMAIN'),
        projectId: getFirebaseValue('PROJECT_ID'),
        storageBucket: getFirebaseValue('STORAGE_BUCKET'),
        messagingSenderId: getFirebaseValue('MESSAGING_SENDER_ID'),
        appId: getFirebaseValue('APP_ID'),
      }),
    },
  };
});
