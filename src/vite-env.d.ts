/// <reference types="vite/client" />

/** Version from package.json, injected by Vite at build time. */
declare const __APP_VERSION__: string;
/** Git commit and build date, injected by Vite at build time. */
declare const __BUILD_INFO__: string;

/** True in the browser build (`vite --mode web`); false in the Tauri app build. */
declare const __WEB_BUILD__: boolean;
