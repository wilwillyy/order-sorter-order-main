import { initializeApp, getApps } from 'firebase/app';
import { getAuth, signInAnonymously, onAuthStateChanged } from 'firebase/auth';
import { getFirestore } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
};

const hasFirebaseConfig = Object.values(firebaseConfig).every(value => Boolean(value));

export const firebaseApp = hasFirebaseConfig && getApps().length === 0 ? initializeApp(firebaseConfig) : getApps()[0] ?? null;
export const auth = firebaseApp ? getAuth(firebaseApp) : null;
export const db = firebaseApp ? getFirestore(firebaseApp) : null;

export const ensureAnonymousSession = async () => {
  if (!auth) return;

  if (auth.currentUser) return;

  try {
    await signInAnonymously(auth);
  } catch (error) {
    console.error('Gagal membuat sesi anonim Firebase:', error);
  }
};

declare global {
  interface Window {
    __ORDER_SORTER_FIREBASE__?: {
      auth: typeof auth;
      db: typeof db;
    };
  }
}

if (typeof window !== 'undefined' && auth && db) {
  window.__ORDER_SORTER_FIREBASE__ = { auth, db };

  onAuthStateChanged(auth, user => {
    console.log('[Firebase Debug] auth state:', user ? 'anonymous user ready' : 'not authenticated');
  });
}
