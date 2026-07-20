import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('electron', {
  isElectron: true,
  platform: process.platform,
  tally: {
    testConnection: (config: any, xml?: string) => ipcRenderer.invoke('tally:testConnection', { config, xml }),
    fetchCompany: (config: any, xml?: string) => ipcRenderer.invoke('tally:fetchCompany', { config, xml }),
    fetchMasters: (config: any, xml?: string) => ipcRenderer.invoke('tally:fetchMasters', { config, xml }),
    fetchLedgers: (config: any, xml: string) => ipcRenderer.invoke('tally:fetchLedgers', { config, xml }),
    fetchGroups: (config: any, xml: string) => ipcRenderer.invoke('tally:fetchGroups', { config, xml }),
    fetchStockItems: (config: any, xml: string) => ipcRenderer.invoke('tally:fetchStockItems', { config, xml }),
    fetchUnits: (config: any, xml: string) => ipcRenderer.invoke('tally:fetchUnits', { config, xml }),
    fetchDaybook: (config: any, xml: string) => ipcRenderer.invoke('tally:fetchDaybook', { config, xml }),
    pushXml: (config: any, xml: string) => ipcRenderer.invoke('tally:pushXml', { config, xml }),
    fetchCompanies: (config: any, xml: string) => ipcRenderer.invoke('tally:fetchCompanies', { config, xml }),
    fetchMastersForCompany: (config: any, xml: string) => ipcRenderer.invoke('tally:fetchMastersForCompany', { config, xml }),
    fetchDaybookForCompany: (config: any, xml: string) => ipcRenderer.invoke('tally:fetchDaybookForCompany', { config, xml }),
  },
});
