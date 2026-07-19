export interface ElectronTallyConfig {
  host: string;
  port: number;
}

export interface ElectronTallyResponse {
  success: boolean;
  response?: string;
  error?: string;
}

export interface ElectronAPI {
  isElectron: boolean;
  platform: string;
  tally: {
    testConnection: (config: ElectronTallyConfig, xml?: string) => Promise<ElectronTallyResponse>;
    fetchCompany: (config: ElectronTallyConfig, xml?: string) => Promise<ElectronTallyResponse>;
    fetchMasters: (config: ElectronTallyConfig, xml?: string) => Promise<ElectronTallyResponse>;
    fetchLedgers: (config: ElectronTallyConfig, xml: string) => Promise<ElectronTallyResponse>;
    fetchGroups: (config: ElectronTallyConfig, xml: string) => Promise<ElectronTallyResponse>;
    fetchStockItems: (config: ElectronTallyConfig, xml: string) => Promise<ElectronTallyResponse>;
    fetchUnits: (config: ElectronTallyConfig, xml: string) => Promise<ElectronTallyResponse>;
    fetchDaybook: (config: ElectronTallyConfig, xml: string) => Promise<ElectronTallyResponse>;
    pushXml: (config: ElectronTallyConfig, xml: string) => Promise<ElectronTallyResponse>;
  };
}

declare global {
  interface Window {
    electron: ElectronAPI;
  }
}
