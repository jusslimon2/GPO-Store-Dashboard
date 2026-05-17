import { initializeApp } from 'firebase/app';
import { getFirestore, collection, addDoc, getDocs, doc, updateDoc, deleteDoc, query, orderBy, writeBatch } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: "AIzaSyDuz2AidAjBqEkoKvUOvEFXPvOEEXY5mps",
  authDomain: "gpo-dashboard-d71da.firebaseapp.com",
  projectId: "gpo-dashboard-d71da",
  storageBucket: "gpo-dashboard-d71da.firebasestorage.app",
  messagingSenderId: "526678782097",
  appId: "1:526678782097:web:eacc6cbe6e2d675f39cf30"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);

export interface GpoRecord {
  gpo: string;
  interestRate: number;
  profit: number;
  usedLimit: number;
}

export interface GpoDoc extends GpoRecord {
  id: string;
  batchId: string;
  createdAt: string;
}

const COLLECTION = 'gpo_statistics';

export async function fetchAllRecords(): Promise<GpoDoc[]> {
  const q = query(collection(db, COLLECTION), orderBy('createdAt', 'desc'));
  const snap = await getDocs(q);
  return snap.docs.map((d) => {
    const data = d.data();
    return {
      id: d.id,
      batchId: data.batchId || '',
      gpo: data.gpo || '',
      interestRate: data.interestRate ?? 0.003,
      profit: data.profit ?? 0,
      usedLimit: data.usedLimit ?? 0,
      createdAt: data.createdAt || '',
    };
  });
}

export async function saveRecords(records: GpoRecord[]): Promise<void> {
  const batchId = crypto.randomUUID();
  const now = new Date().toISOString();
  const col = collection(db, COLLECTION);
  const batch = writeBatch(db);

  // Remove all existing records first so the new upload replaces everything.
  const existingDocs = await getDocs(col);
  existingDocs.docs.forEach((existingDoc) => batch.delete(existingDoc.ref));

  for (const r of records) {
    const docRef = doc(col);
    batch.set(docRef, { ...r, batchId, createdAt: now });
  }

  await batch.commit();
}

export async function updateRecord(id: string, updates: Partial<GpoRecord>): Promise<void> {
  const ref = doc(db, COLLECTION, id);
  await updateDoc(ref, updates);
}

export async function removeRecord(id: string): Promise<void> {
  const ref = doc(db, COLLECTION, id);
  await deleteDoc(ref);
}
