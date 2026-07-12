import { 
  doc, 
  getDoc, 
  setDoc, 
  deleteDoc, 
  collection, 
  addDoc, 
  updateDoc, 
  onSnapshot, 
  query, 
  where 
} from 'firebase/firestore';
import { db } from './firebase';

export function isElectron(): boolean {
  return typeof window !== 'undefined' && (window as any).electron?.isElectron === true;
}

export type AppMode = 'web' | 'desktop-offline';

export function getAppMode(): AppMode {
  return isElectron() ? 'desktop-offline' : 'web';
}

export interface OfflineUser {
  uid: string;
  displayName: string;
  email: string;
}

export const OFFLINE_USER: OfflineUser = {
  uid: 'local-desktop-user',
  displayName: 'Offline User',
  email: 'offline@tallygen.local'
};

// --- IndexedDB Setup for Desktop Offline ---
const DB_NAME = 'tallygen_offline_db';
const DB_VERSION = 1;

function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onerror = () => reject(request.error);
    request.onsuccess = () => resolve(request.result);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains('tally_context')) {
        db.createObjectStore('tally_context', { keyPath: 'userId' });
      }
      if (!db.objectStoreNames.contains('conversions')) {
        db.createObjectStore('conversions', { keyPath: 'id' });
      }
    };
  });
}

async function idbGet(storeName: string, key: string): Promise<any> {
  try {
    const dbInstance = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = dbInstance.transaction(storeName, 'readonly');
      const store = transaction.objectStore(storeName);
      const request = store.get(key);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error);
    });
  } catch (err) {
    console.error(`IndexedDB get error from ${storeName}:`, err);
    return null;
  }
}

async function idbGetAll(storeName: string): Promise<any[]> {
  try {
    const dbInstance = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = dbInstance.transaction(storeName, 'readonly');
      const store = transaction.objectStore(storeName);
      const request = store.getAll();
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error);
    });
  } catch (err) {
    console.error(`IndexedDB getAll error from ${storeName}:`, err);
    return [];
  }
}

async function idbPut(storeName: string, item: any): Promise<void> {
  try {
    const dbInstance = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = dbInstance.transaction(storeName, 'readwrite');
      const store = transaction.objectStore(storeName);
      const request = store.put(item);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  } catch (err) {
    console.error(`IndexedDB put error in ${storeName}:`, err);
  }
}

async function idbDelete(storeName: string, key: string): Promise<void> {
  try {
    const dbInstance = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = dbInstance.transaction(storeName, 'readwrite');
      const store = transaction.objectStore(storeName);
      const request = store.delete(key);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  } catch (err) {
    console.error(`IndexedDB delete error from ${storeName}:`, err);
  }
}

// In-Memory listeners to emulate onSnapshot for local IndexedDB state changes
type TallyContextListener = (context: any) => void;
type ConversionsListener = (conversions: any[]) => void;

const tallyContextListeners = new Map<string, Set<TallyContextListener>>();
const conversionsListeners = new Map<string, Set<ConversionsListener>>();

function notifyTallyContext(userId: string, context: any) {
  const listeners = tallyContextListeners.get(userId);
  if (listeners) {
    listeners.forEach(fn => fn(context));
  }
}

async function notifyConversions(userId: string) {
  const listeners = conversionsListeners.get(userId);
  if (listeners) {
    const allConversions = await getOfflineConversions(userId);
    listeners.forEach(fn => fn(allConversions));
  }
}

async function getOfflineConversions(userId: string): Promise<any[]> {
  const list = await idbGetAll('conversions');
  return list
    .filter((item: any) => item.uid === userId)
    .sort((a, b) => {
      const tA = a.timestamp?.seconds || 0;
      const tB = b.timestamp?.seconds || 0;
      return tB - tA;
    });
}

// --- Unified Storage Adapter API ---

/**
 * Subscribe to Tally Context updates.
 * Unsubscribe function is returned.
 */
export function subscribeToTallyContext(userId: string, onUpdate: TallyContextListener): () => void {
  if (getAppMode() === 'web') {
    return onSnapshot(doc(db, 'tally_context', userId), (snapshot) => {
      onUpdate(snapshot.exists() ? snapshot.data() : null);
    }, (err) => {
      console.error('Tally context snapshot error:', err);
    });
  } else {
    // Desktop Offline: read from IndexedDB, add to listeners
    let active = true;
    
    if (!tallyContextListeners.has(userId)) {
      tallyContextListeners.set(userId, new Set());
    }
    tallyContextListeners.get(userId)!.add(onUpdate);

    idbGet('tally_context', userId).then((data) => {
      if (active) {
        onUpdate(data);
      }
    });

    return () => {
      active = false;
      const set = tallyContextListeners.get(userId);
      if (set) {
        set.delete(onUpdate);
        if (set.size === 0) {
          tallyContextListeners.delete(userId);
        }
      }
    };
  }
}

/**
 * Subscribe to Conversions list.
 * Unsubscribe function is returned.
 */
export function subscribeToConversions(userId: string, onUpdate: ConversionsListener): () => void {
  if (getAppMode() === 'web') {
    const q = query(collection(db, 'conversions'), where('uid', '==', userId));
    return onSnapshot(q, (snapshot) => {
      const list = snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
      // Sort in descending order of timestamp seconds
      const sorted = list.sort((a: any, b: any) => {
        const tA = a.timestamp?.seconds || 0;
        const tB = b.timestamp?.seconds || 0;
        return tB - tA;
      });
      onUpdate(sorted);
    }, (err) => {
      console.error('Conversions snapshot error:', err);
    });
  } else {
    let active = true;

    if (!conversionsListeners.has(userId)) {
      conversionsListeners.set(userId, new Set());
    }
    conversionsListeners.get(userId)!.add(onUpdate);

    getOfflineConversions(userId).then((list) => {
      if (active) {
        onUpdate(list);
      }
    });

    return () => {
      active = false;
      const set = conversionsListeners.get(userId);
      if (set) {
        set.delete(onUpdate);
        if (set.size === 0) {
          conversionsListeners.delete(userId);
        }
      }
    };
  }
}

/**
 * Saves Tally Context.
 */
export async function saveTallyContext(userId: string, context: any): Promise<void> {
  if (getAppMode() === 'web') {
    await setDoc(doc(db, 'tally_context', userId), context);
  } else {
    const item = { userId, ...context };
    await idbPut('tally_context', item);
    notifyTallyContext(userId, item);
  }
}

/**
 * Deletes Tally Context.
 */
export async function deleteTallyContext(userId: string): Promise<void> {
  if (getAppMode() === 'web') {
    await deleteDoc(doc(db, 'tally_context', userId));
  } else {
    await idbDelete('tally_context', userId);
    notifyTallyContext(userId, null);
  }
}

/**
 * Creates/Saves a new Conversion record.
 * Returns the record's ID.
 */
export async function saveConversion(userId: string, conversion: any): Promise<string> {
  if (getAppMode() === 'web') {
    // If we're on the web, use firestore
    const docRef = await addDoc(collection(db, 'conversions'), {
      ...conversion,
      uid: userId
    });
    return docRef.id;
  } else {
    const id = 'conv_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    const timestamp = { seconds: Math.floor(Date.now() / 1000) };
    const item = {
      id,
      ...conversion,
      uid: userId,
      timestamp
    };
    await idbPut('conversions', item);
    await notifyConversions(userId);
    return id;
  }
}

/**
 * Updates an existing Conversion record.
 */
export async function updateConversion(userId: string, conversionId: string, updates: any): Promise<void> {
  if (getAppMode() === 'web') {
    await updateDoc(doc(db, 'conversions', conversionId), updates);
  } else {
    const existing = await idbGet('conversions', conversionId);
    if (existing) {
      const updated = {
        ...existing,
        ...updates
      };
      await idbPut('conversions', updated);
      await notifyConversions(userId);
    }
  }
}

/**
 * Clears the offline workspace (Tally context) from IndexedDB.
 */
export async function clearOfflineWorkspace(userId: string): Promise<void> {
  if (getAppMode() === 'desktop-offline') {
    await idbDelete('tally_context', userId);
    notifyTallyContext(userId, null);
  }
}
