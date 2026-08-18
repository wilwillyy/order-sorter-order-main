import { initializeApp, getApps } from 'firebase/app';
import { getAuth, signInAnonymously, onAuthStateChanged } from 'firebase/auth';
import { getFirestore } from 'firebase/firestore';

const firebaseConfig = __FIREBASE_CONFIG__;

const missingFirebaseConfigKeys = Object.entries(firebaseConfig)
  .filter(([, value]) => !value)
  .map(([key]) => key);
const hasFirebaseConfig = missingFirebaseConfigKeys.length === 0;

if (typeof window !== 'undefined' && missingFirebaseConfigKeys.length > 0) {
  console.error('[Firebase Debug] Missing Vercel build variables:', missingFirebaseConfigKeys);
}

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
