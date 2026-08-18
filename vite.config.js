import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { loadEnv } from 'vite';
export default defineConfig(function (_a) {
    var mode = _a.mode;
    var env = loadEnv(mode, process.cwd(), '');
    var getFirebaseValue = function (key) { return env["firebase_".concat(key.toLowerCase())] || env["FIREBASE_".concat(key)] || env["VITE_FIREBASE_".concat(key)] || ''; };
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
