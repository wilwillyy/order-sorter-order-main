import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { loadEnv } from 'vite';
export default defineConfig(function (_a) {
    var mode = _a.mode;
    var env = loadEnv(mode, process.cwd(), '');
    var getFirebaseValue = function (key) {
        var normalizedKey = key.toLowerCase();
        var aliases = key === 'PROJECT_ID'
            ? ['firebase_project_id', 'firebase_projectid', 'FIREBASE_PROJECT_ID', 'VITE_FIREBASE_PROJECT_ID']
            : ["firebase_".concat(normalizedKey), "FIREBASE_".concat(key), "VITE_FIREBASE_".concat(key)];
        return aliases.map(function (alias) { return env[alias]; }).find(Boolean) || '';
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
