import { contextBridge } from 'electron';

contextBridge.exposeInMainWorld('electron', {
  isElectron: true,
  platform: process.platform,
});
