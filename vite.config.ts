import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { loadEnv } from 'vite';

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '');
  const getFirebaseValue = (key: string) => {
    const normalizedKey = key.toLowerCase();
    const aliases = key === 'PROJECT_ID'
      ? ['firebase_project_id', 'firebase_projectid', 'FIREBASE_PROJECT_ID', 'VITE_FIREBASE_PROJECT_ID']
      : [`firebase_${normalizedKey}`, `FIREBASE_${key}`, `VITE_FIREBASE_${key}`];

    return aliases.map(alias => env[alias]).find(Boolean) || '';
  };

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
