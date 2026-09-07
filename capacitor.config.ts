import type { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'com.dinikutuphane.app',
  appName: 'Dini Kutuphane',
  webDir: 'public',
  android: {
    allowMixedContent: false
  }
};

export default config;