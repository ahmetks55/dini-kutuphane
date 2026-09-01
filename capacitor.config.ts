import type { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'com.dinikutuphane.app',
  appName: 'Dini Kutuphane',
  webDir: 'public',
  server: {
    url: 'https://dini-kutuphane.onrender.com',
    cleartext: false
  },
  android: {
    allowMixedContent: false
  }
};

export default config;
