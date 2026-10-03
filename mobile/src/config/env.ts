// ── API environments ────────────────────────────────────────────────────────
// Pick which backend a DEBUG build talks to by changing API_ENV below.
// Release builds always use `prod` regardless of this setting.
type ApiEnv = 'local' | 'dev' | 'prod';
const API_ENV: ApiEnv = 'local';

// 10.0.2.2 = host machine's localhost as seen from the Android emulator.
// On a physical device use `adb reverse tcp:8010 tcp:8010` so localhost maps to your PC,
// or set LOCAL to your machine's LAN IP (e.g. http://192.168.1.50:8010).
const isEmulator = true; // Android emulator → host machine; set false for physical device + adb reverse
// Port 8010 — local Freightflex API (do not point at remote hosts while developing)
const LOCAL = isEmulator ? 'http://10.0.2.2:8010' : 'http://localhost:8010';

// Kept as local so nothing accidentally calls the remote server
const DEV = LOCAL;
const PROD = LOCAL;

const ORIGINS: Record<ApiEnv, string> = {local: LOCAL, dev: DEV, prod: PROD};

export const API_ORIGIN = ORIGINS[API_ENV].replace(/\/+$/, '');
export const API_BASE_URL = `${API_ORIGIN}/api/v1`;
export const WS_BASE_URL = API_ORIGIN.replace(/^https/, 'ws').replace(/^http/, 'ws');

// Same key as AndroidManifest — set when you have a local/dev Maps key
export const GOOGLE_MAPS_API_KEY = '';
