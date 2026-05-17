import { useState, useRef, useEffect, useCallback } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import {
  Upload,
  Save,
  BarChart3,
  FileSpreadsheet,
  CheckCircle,
  AlertCircle,
  Loader2,
  ArrowLeft,
  Lock,
  Pencil,
  X,
  Check,
  Trash2,
} from 'lucide-react';
import {
  fetchAllRecords,
  saveRecords,
  updateRecord,
  removeRecord,
  type GpoDoc,
  type GpoRecord,
} from '../lib/firebase';

const INTEREST_RATE = 0.003;
const ADMIN_PASSWORD = import.meta.env.VITE_ADMIN_PASSWORD || 'gpo@2024';

function maskGpo(gpo: unknown): string {
  if (!gpo) return '';
  const str = String(gpo);
  if (str.length <= 6) return str;
  return str.slice(0, 2) + '****' + str.slice(-4);
}

function formatNumber(n: number): string {
  return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 3 });
}

type ToastType = 'success' | 'error';

interface Toast {
  type: ToastType;
  message: string;
}

interface EditState {
  id: string;
  field: keyof GpoRecord;
  value: string;
}

interface PreviewState {
  isOpen: boolean;
  data: GpoRecord[];
}

interface ConfirmSaveState {
  isOpen: boolean;
  data: GpoRecord[];
}

export default function Admin() {
  const [authenticated, setAuthenticated] = useState(false);
  const [password, setPassword] = useState('');
  const [passwordError, setPasswordError] = useState('');

  const [records, setRecords] = useState<GpoDoc[]>([]);
  const [uploadedRecords, setUploadedRecords] = useState<GpoRecord[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [fileName, setFileName] = useState('');
  const [saving, setSaving] = useState(false);
  const [loading, setLoading] = useState(false);
  const [hasIcon, setHasIcon] = useState(false);
  const [toast, setToast] = useState<Toast | null>(null);
  const [editState, setEditState] = useState<EditState | null>(null);
  const [deleteConfirm, setDeleteConfirm] = useState<string | null>(null);
  const [pendingEditConfirm, setPendingEditConfirm] = useState<EditState | null>(null);
  const [preview, setPreview] = useState<PreviewState>({ isOpen: false, data: [] });
  const [confirmSave, setConfirmSave] = useState<ConfirmSaveState>({ isOpen: false, data: [] });
  const fileInputRef = useRef<HTMLInputElement>(null);
  const editInputRef = useRef<HTMLInputElement>(null);

  function showToast(type: ToastType, message: string) {
    setToast({ type, message });
    setTimeout(() => setToast(null), 4000);
  }

  function handleLogin(e: React.FormEvent) {
    e.preventDefault();
    if (password === ADMIN_PASSWORD) {
      setAuthenticated(true);
      setPasswordError('');
    } else {
      setPasswordError('Incorrect password');
    }
  }

  const loadRecords = useCallback(async () => {
    setLoading(true);
    try {
      const data = await fetchAllRecords();
      setRecords(data);
    } catch (err: unknown) {
      showToast('error', (err as Error).message);
    }
    setLoading(false);
  }, []);

  useEffect(() => {
    if (authenticated) loadRecords();
  }, [authenticated, loadRecords]);

  useEffect(() => {
    if (editState && editInputRef.current) {
      editInputRef.current.focus();
      editInputRef.current.select();
    }
  }, [editState]);

  function handleFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1 });

      const parsed: GpoRecord[] = [];
      for (let i = 1; i < rows.length; i++) {
        const row = (rows[i] as unknown[]) || [];
        const gpoRaw = row[1];
        if (!gpoRaw) continue;
        const usedLimit = parseFloat(String(row[7])) || 0;
        const profit = parseFloat((usedLimit * INTEREST_RATE).toFixed(3));
        parsed.push({
          gpo: maskGpo(gpoRaw),
          interestRate: INTEREST_RATE,
          profit,
          usedLimit,
        });
      }
      setUploadedRecords(parsed);
      setPreview({ isOpen: true, data: parsed });
    };
    reader.readAsArrayBuffer(file);
  }

  function openSaveConfirmation() {
    if (!uploadedRecords.length) {
      showToast('error', 'No data to save. Please upload a file first.');
      return;
    }
    setConfirmSave({ isOpen: true, data: uploadedRecords });
  }

  async function handleConfirmSave() {
    setSaving(true);
    try {
      await saveRecords(confirmSave.data);
      showToast('success', `${confirmSave.data.length} records saved successfully.`);
      setUploadedRecords([]);
      setFileName('');
      setPreview({ isOpen: false, data: [] });
      setConfirmSave({ isOpen: false, data: [] });
      if (fileInputRef.current) fileInputRef.current.value = '';
      await loadRecords();
    } catch (err: unknown) {
      showToast('error', 'Error saving: ' + (err as Error).message);
    }
    setSaving(false);
  }

  function startEdit(id: string, field: keyof GpoRecord, value: number | string) {
    setEditState({ id, field, value: String(value) });
  }

  function requestCommitEdit() {
    if (editState) {
      setPendingEditConfirm(editState);
    }
  }

  async function performEdit(edit: EditState) {
    const { id, field, value } = edit;
    const numValue = parseFloat(value);
    if (field !== 'gpo' && isNaN(numValue)) {
      showToast('error', 'Please enter a valid number.');
      return false;
    }

    const updates: Partial<GpoRecord> = {};
    if (field === 'gpo') {
      updates.gpo = value;
    } else {
      updates[field] = numValue;
    }

    try {
      await updateRecord(id, updates);
      setRecords((prev) =>
        prev.map((r) => (r.id === id ? { ...r, ...updates } : r))
      );
      showToast('success', 'Record updated.');
      return true;
    } catch (err: unknown) {
      showToast('error', 'Error updating: ' + (err as Error).message);
      return false;
    }
  }

  async function confirmEdit() {
    if (!pendingEditConfirm) return;
    const success = await performEdit(pendingEditConfirm);
    if (success) {
      setEditState(null);
    }
    setPendingEditConfirm(null);
  }

  function cancelEdit() {
    setEditState(null);
    setPendingEditConfirm(null);
  }

  async function confirmDelete(id: string) {
    try {
      await removeRecord(id);
      setRecords((prev) => prev.filter((r) => r.id !== id));
      showToast('success', 'Record deleted.');
    } catch (err: unknown) {
      showToast('error', 'Error deleting: ' + (err as Error).message);
    }
    setDeleteConfirm(null);
  }

  const filteredRecords = records.filter((item) => {
    const query = searchQuery.trim().toLowerCase();
    if (!query) return true;
    return item.gpo.toLowerCase().includes(query);
  });

  useEffect(() => {
    const img = new Image();
    img.src = '/icon.png';
    img.onload = () => setHasIcon(true);
    img.onerror = () => setHasIcon(false);
  }, []);

  // Password gate
  if (!authenticated) {
    return (
      <div className="min-h-screen bg-white flex items-center justify-center px-4">
        <div className="w-full max-w-sm">
          <div className="bg-white rounded border border-slate-200 p-6">
            <div className="flex flex-col items-center mb-6">
              <div className="p-2 bg-slate-800 rounded mb-3">
                <Lock className="w-5 h-5 text-white" />
              </div>
              <h1 className="text-lg font-bold text-slate-800">Admin Access</h1>
              <p className="text-xs text-slate-500 mt-1">Enter password to continue</p>
            </div>
            <form onSubmit={handleLogin} className="space-y-3">
              <div>
                <input
                  type="password"
                  value={password}
                  onChange={(e) => {
                    setPassword(e.target.value);
                    setPasswordError('');
                  }}
                  placeholder="Password"
                  className="w-full px-3 py-2 border border-slate-300 rounded text-xs focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent transition-all"
                  autoFocus
                />
                {passwordError && (
                  <p className="text-xs text-red-600 mt-1 flex items-center gap-1">
                    <AlertCircle className="w-3 h-3" />
                    {passwordError}
                  </p>
                )}
              </div>
              <button
                type="submit"
                className="w-full py-2 px-3 bg-slate-800 hover:bg-slate-900 text-white font-semibold text-xs rounded transition-colors"
              >
                Unlock
              </button>
            </form>
            <Link
              to="/"
              className="flex items-center justify-center gap-1.5 text-xs text-slate-500 hover:text-slate-700 mt-4 transition-colors"
            >
              <ArrowLeft className="w-3 h-3" />
              Back to Dashboard
            </Link>
          </div>
        </div>
      </div>
    );
  }

  // Admin panel
  const totalProfit = records.reduce((s, r) => s + r.profit, 0);
  const totalLimit = records.reduce((s, r) => s + r.usedLimit, 0);

  const uploadProfit = uploadedRecords.reduce((s, r) => s + r.profit, 0);
  const uploadLimit = uploadedRecords.reduce((s, r) => s + r.usedLimit, 0);

  return (
    <div className="min-h-screen bg-white">
      {/* Toast */}
      {toast && (
        <div
          className={`fixed top-3 right-3 z-50 flex items-center gap-2 px-4 py-2 rounded text-xs font-medium transition-all duration-300 ${
            toast.type === 'success' ? 'bg-emerald-600 text-white' : 'bg-red-600 text-white'
          }`}
        >
          {toast.type === 'success' ? (
            <CheckCircle className="w-3.5 h-3.5 shrink-0" />
          ) : (
            <AlertCircle className="w-3.5 h-3.5 shrink-0" />
          )}
          {toast.message}
        </div>
      )}

      <div className="max-w-6xl mx-auto px-4 py-6">
        {/* Header */}
        <div className="flex items-center justify-between mb-6">
          <div className="flex items-center gap-3">
            {hasIcon ? (
              <div className="flex items-center">
                <img src="/icon.png" alt="GPO" className="w-5 h-5 object-contain" />
              </div>
            ) : (
              <div className="p-2 bg-slate-800 rounded">
                <BarChart3 className="w-5 h-5 text-white" />
              </div>
            )}
            <div>
              <h1 className="text-xl font-bold text-slate-800">GPO Statistics</h1>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={() => setAuthenticated(false)}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-red-50 hover:bg-red-100 text-red-700 text-xs font-semibold rounded border border-red-200 transition-colors"
            >
              <Lock className="w-3.5 h-3.5" />
              Lock
            </button>
            <Link
              to="/"
              className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded border border-slate-200 transition-colors"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              Dashboard
            </Link>
          </div>
        </div>

        {/* Upload Section */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mb-6">
          <div className="md:col-span-2 bg-slate-50 rounded border border-slate-300 p-4">
            <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">
              Upload File
            </p>
            <button
              onClick={() => fileInputRef.current?.click()}
              className="w-full flex items-center gap-2 px-3 py-2.5 border-2 border-dashed border-slate-300 hover:border-blue-400 hover:bg-blue-50 rounded transition-all duration-200 text-left group"
            >
              <div className="p-1.5 bg-slate-100 group-hover:bg-blue-100 rounded transition-colors">
                <FileSpreadsheet className="w-4 h-4 text-slate-600 group-hover:text-blue-600 transition-colors" />
              </div>
              <div className="flex-1">
                {fileName ? (
                  <>
                    <p className="text-xs font-semibold text-slate-700">{fileName}</p>
                    <p className="text-xs text-slate-500">{uploadedRecords.length} records</p>
                  </>
                ) : (
                  <>
                    <p className="text-xs font-semibold text-slate-700">Choose Excel or CSV</p>
                    <p className="text-xs text-slate-500">.xlsx, .xls, .csv</p>
                  </>
                )}
              </div>
              <Upload className="w-3.5 h-3.5 text-slate-400 group-hover:text-blue-500 transition-colors" />
            </button>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={handleFile}
            />
          </div>

          <div className="bg-slate-50 rounded border border-slate-300 p-4 flex flex-col justify-between">
            <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">
              Action
            </p>
            <button
              onClick={openSaveConfirmation}
              disabled={saving || uploadedRecords.length === 0}
              className="flex items-center justify-center gap-1.5 w-full py-2 px-3 bg-blue-600 hover:bg-blue-700 disabled:bg-slate-300 disabled:text-slate-500 text-white font-semibold text-xs rounded transition-all duration-200"
            >
              {saving ? <Loader2 className="w-3 h-3 animate-spin" /> : <Save className="w-3 h-3" />}
              {saving ? 'Saving...' : 'Review & Save'}
            </button>
          </div>
        </div>

        {/* Upload Preview Summary */}
        {uploadedRecords.length > 0 && (
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-6">
            <div className="bg-slate-100 rounded p-3">
              <p className="text-xs text-slate-500 font-medium mb-1">Records to Upload</p>
              <p className="text-lg font-bold text-slate-800">{uploadedRecords.length}</p>
            </div>
            <div className="bg-slate-100 rounded p-3">
              <p className="text-xs text-slate-500 font-medium mb-1">Upload Profit</p>
              <p className="text-lg font-bold text-slate-800">{formatNumber(uploadProfit)}</p>
            </div>
            <div className="bg-slate-100 rounded p-3">
              <p className="text-xs text-slate-500 font-medium mb-1">Upload Used Limit</p>
              <p className="text-lg font-bold text-slate-800">{formatNumber(uploadLimit)}</p>
            </div>
          </div>
        )}

        <div className="mb-4">
          <input
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search GPO by last 4 digits"
            className="w-full max-w-sm px-3 py-2 border border-slate-300 rounded text-xs focus:outline-none focus:ring-1 focus:ring-blue-500"
          />
        </div>

        {/* Saved Records Table */}
        <div className="bg-white rounded border border-slate-200 overflow-hidden mb-6">
          <div className="px-4 py-3 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
            <h2 className="text-xs font-semibold text-slate-700 uppercase">
              Saved Records ({records.length})
            </h2>
            <button
              onClick={loadRecords}
              disabled={loading}
              className="text-xs text-blue-600 hover:text-blue-800 font-semibold transition-colors"
            >
              {loading ? 'Loading...' : 'Refresh'}
            </button>
          </div>

          {loading && records.length === 0 ? (
            <div className="flex items-center justify-center py-12">
              <Loader2 className="w-5 h-5 text-blue-500 animate-spin" />
            </div>
          ) : records.length === 0 ? (
            <div className="flex flex-col items-center justify-center py-12 text-slate-400">
              <FileSpreadsheet className="w-8 h-8 mb-2 text-slate-200" />
              <p className="text-xs font-medium">No saved records</p>
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-slate-200 border-b border-slate-300">
                    <th className="px-3 py-2 text-left font-bold text-slate-700 w-10">#</th>
                    <th className="px-3 py-2 text-left font-bold text-slate-700">GPO Number</th>
                    <th className="px-3 py-2 text-center font-bold text-slate-700">Interest Rate</th>
                    <th className="px-3 py-2 text-right font-bold text-slate-700">Profit</th>
                    <th className="px-3 py-2 text-right font-bold text-slate-700">Used Limit</th>
                    <th className="px-3 py-2 text-center font-bold text-slate-700 w-16">Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRecords.map((item, i) => {
                    const isEditingThis = editState?.id === item.id;
                    return (
                      <tr key={item.id} className="border-b border-slate-100 hover:bg-slate-50 group">
                        <td className="px-3 py-2 text-slate-500 text-xs">{i + 1}</td>
                        <td className="px-3 py-2 bg-blue-100">
                          {isEditingThis && editState.field === 'gpo' ? (
                            <input
                              ref={editInputRef}
                              type="text"
                              value={editState.value}
                              onChange={(e) => setEditState({ ...editState, value: e.target.value })}
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') requestCommitEdit();
                                if (e.key === 'Escape') cancelEdit();
                              }}
                              className="w-full px-1.5 py-0.5 text-xs border border-blue-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500"
                            />
                          ) : (
                            <span
                              className="text-slate-800 font-medium cursor-pointer hover:underline"
                              onClick={() => startEdit(item.id, 'gpo', item.gpo)}
                            >
                              {item.gpo}
                            </span>
                          )}
                        </td>
                        <td className="px-3 py-2 bg-pink-100 text-center">
                          <span className="text-slate-800 font-medium">
                            {(item.interestRate * 100).toFixed(1)}%
                          </span>
                        </td>
                        <td className="px-3 py-2 bg-pink-100 text-right">
                          {isEditingThis && editState.field === 'profit' ? (
                            <input
                              ref={editInputRef}
                              type="number"
                              step="0.001"
                              value={editState.value}
                              onChange={(e) => setEditState({ ...editState, value: e.target.value })}
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') requestCommitEdit();
                                if (e.key === 'Escape') cancelEdit();
                              }}
                              className="w-20 px-1.5 py-0.5 text-xs text-right border border-blue-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500"
                            />
                          ) : (
                            <span
                              className="text-slate-800 font-medium cursor-pointer hover:underline"
                              onClick={() => startEdit(item.id, 'profit', item.profit)}
                            >
                              {formatNumber(item.profit)}
                            </span>
                          )}
                        </td>
                        <td className="px-3 py-2 bg-green-100 text-right">
                          {isEditingThis && editState.field === 'usedLimit' ? (
                            <input
                              ref={editInputRef}
                              type="number"
                              step="0.01"
                              value={editState.value}
                              onChange={(e) => setEditState({ ...editState, value: e.target.value })}
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') requestCommitEdit();
                                if (e.key === 'Escape') cancelEdit();
                              }}
                              className="w-24 px-1.5 py-0.5 text-xs text-right border border-blue-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500"
                            />
                          ) : (
                            <span
                              className="text-slate-800 font-medium cursor-pointer hover:underline"
                              onClick={() => startEdit(item.id, 'usedLimit', item.usedLimit)}
                            >
                              {formatNumber(item.usedLimit)}
                            </span>
                          )}
                        </td>
                        <td className="px-3 py-2 text-center">
                          {isEditingThis ? (
                            <div className="flex items-center justify-center gap-0.5">
                              <button
                                onClick={requestCommitEdit}
                                className="p-1 rounded bg-emerald-100 text-emerald-700 hover:bg-emerald-200 transition-colors"
                              >
                                <Check className="w-3 h-3" />
                              </button>
                              <button
                                onClick={cancelEdit}
                                className="p-1 rounded bg-slate-100 text-slate-500 hover:bg-slate-200 transition-colors"
                              >
                                <X className="w-3 h-3" />
                              </button>
                            </div>
                          ) : (
                            <div className="flex items-center justify-center gap-0.5 opacity-0 group-hover:opacity-100 transition-opacity">
                              <button
                                onClick={() => startEdit(item.id, 'gpo', item.gpo)}
                                className="p-1 rounded bg-blue-100 text-blue-600 hover:bg-blue-200 transition-colors"
                              >
                                <Pencil className="w-3 h-3" />
                              </button>
                              <button
                                onClick={() => setDeleteConfirm(item.id)}
                                className="p-1 rounded bg-red-100 text-red-600 hover:bg-red-200 transition-colors"
                              >
                                <Trash2 className="w-3 h-3" />
                              </button>
                            </div>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
                {records.length > 1 && (
                  <tfoot>
                    <tr className="bg-slate-100 border-t-2 border-slate-300">
                      <td colSpan={2} className="px-3 py-2 text-xs font-bold text-slate-700 uppercase">
                        Totals
                      </td>
                      <td className="px-3 py-2" />
                      <td className="px-3 py-2 bg-pink-100 text-right font-bold text-slate-800">
                        {formatNumber(totalProfit)}
                      </td>
                      <td className="px-3 py-2 bg-green-100 text-right font-bold text-slate-800">
                        {formatNumber(totalLimit)}
                      </td>
                      <td className="px-3 py-2" />
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>
          )}
        </div>

        {/* Preview Modal */}
        {preview.isOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4">
            <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full max-h-[80vh] flex flex-col">
              <div className="px-4 py-3 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
                <h2 className="text-sm font-bold text-slate-700">Preview Upload Data</h2>
                <button
                  onClick={() => setPreview({ isOpen: false, data: [] })}
                  className="text-slate-400 hover:text-slate-600 transition-colors"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>

              <div className="flex-1 overflow-auto">
                <table className="w-full text-xs border-collapse">
                  <thead className="sticky top-0 bg-slate-200 border-b border-slate-300">
                    <tr>
                      <th className="px-3 py-2 text-left font-bold text-slate-700 w-10">#</th>
                      <th className="px-3 py-2 text-left font-bold text-slate-700">GPO Number</th>
                      <th className="px-3 py-2 text-center font-bold text-slate-700">Interest Rate</th>
                      <th className="px-3 py-2 text-right font-bold text-slate-700">Profit</th>
                      <th className="px-3 py-2 text-right font-bold text-slate-700">Used Limit</th>
                    </tr>
                  </thead>
                  <tbody>
                    {preview.data.map((item, i) => (
                      <tr key={i} className="border-b border-slate-100 hover:bg-slate-50">
                        <td className="px-3 py-2 text-slate-500">{i + 1}</td>
                        <td className="px-3 py-2 bg-blue-100 font-medium text-slate-800">{item.gpo}</td>
                        <td className="px-3 py-2 bg-pink-100 text-center font-medium text-slate-800">
                          {(item.interestRate * 100).toFixed(1)}%
                        </td>
                        <td className="px-3 py-2 bg-pink-100 text-right font-medium text-slate-800">
                          {formatNumber(item.profit)}
                        </td>
                        <td className="px-3 py-2 bg-green-100 text-right font-medium text-slate-800">
                          {formatNumber(item.usedLimit)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="px-4 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-between gap-2">
                <p className="text-xs text-slate-600">
                  <span className="font-semibold">{preview.data.length}</span> records to upload
                </p>
                <div className="flex gap-2">
                  <button
                    onClick={() => setPreview({ isOpen: false, data: [] })}
                    className="px-3 py-1.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded transition-colors"
                  >
                    Back
                  </button>
                  <button
                    onClick={() => {
                      setPreview({ isOpen: false, data: [] });
                      openSaveConfirmation();
                    }}
                    className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 rounded transition-colors"
                  >
                    <Check className="w-3 h-3" />
                    Proceed to Confirm
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {pendingEditConfirm && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4">
            <div className="bg-white rounded-lg shadow-xl max-w-md w-full">
              <div className="px-4 py-3 border-b border-slate-200 bg-slate-50">
                <h2 className="text-sm font-bold text-slate-700 flex items-center gap-2">
                  <AlertCircle className="w-4 h-4 text-amber-600" />
                  Confirm Edit
                </h2>
              </div>
              <div className="px-4 py-4 space-y-2">
                <p className="text-xs text-slate-600">
                  Are you sure you want to update <span className="font-semibold text-slate-800">{pendingEditConfirm.field}</span> for this record?
                </p>
                <p className="text-xs text-slate-600">
                  New value: <span className="font-semibold text-slate-800">{pendingEditConfirm.value}</span>
                </p>
              </div>
              <div className="px-4 py-3 border-t border-slate-200 bg-slate-50 flex gap-2 justify-end">
                <button
                  onClick={() => setPendingEditConfirm(null)}
                  className="px-3 py-1.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded transition-colors"
                >
                  Cancel
                </button>
                <button
                  onClick={confirmEdit}
                  className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 rounded transition-colors"
                >
                  <Check className="w-3 h-3" />
                  Confirm Edit
                </button>
              </div>
            </div>
          </div>
        )}

        {deleteConfirm && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4">
            <div className="bg-white rounded-lg shadow-xl max-w-md w-full">
              <div className="px-4 py-3 border-b border-slate-200 bg-slate-50">
                <h2 className="text-sm font-bold text-slate-700 flex items-center gap-2">
                  <AlertCircle className="w-4 h-4 text-red-600" />
                  Confirm Delete
                </h2>
              </div>
              <div className="px-4 py-4 space-y-2">
                <p className="text-xs text-slate-600">
                  Are you sure you want to delete this record? This action cannot be undone.
                </p>
              </div>
              <div className="px-4 py-3 border-t border-slate-200 bg-slate-50 flex gap-2 justify-end">
                <button
                  onClick={() => setDeleteConfirm(null)}
                  className="px-3 py-1.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded transition-colors"
                >
                  Cancel
                </button>
                <button
                  onClick={() => confirmDelete(deleteConfirm)}
                  className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-red-600 hover:bg-red-700 rounded transition-colors"
                >
                  <Trash2 className="w-3 h-3" />
                  Delete
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Confirmation Modal */}
        {confirmSave.isOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4">
            <div className="bg-white rounded-lg shadow-xl max-w-md w-full">
              <div className="px-4 py-3 border-b border-slate-200 bg-slate-50">
                <h2 className="text-sm font-bold text-slate-700 flex items-center gap-2">
                  <AlertCircle className="w-4 h-4 text-amber-600" />
                  Confirm Save
                </h2>
              </div>

              <div className="px-4 py-4 space-y-2">
                <p className="text-xs text-slate-600">
                  Are you sure you want to save <span className="font-bold text-slate-800">{confirmSave.data.length}</span> records to the database?
                </p>
                <div className="bg-amber-50 border border-amber-200 rounded p-2.5">
                  <p className="text-xs text-amber-800">
                    This action cannot be easily undone. Please review the data carefully before confirming.
                  </p>
                </div>
              </div>

              <div className="px-4 py-3 border-t border-slate-200 bg-slate-50 flex gap-2 justify-end">
                <button
                  onClick={() => setConfirmSave({ isOpen: false, data: [] })}
                  disabled={saving}
                  className="px-3 py-1.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded transition-colors disabled:opacity-50"
                >
                  Cancel
                </button>
                <button
                  onClick={handleConfirmSave}
                  disabled={saving}
                  className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 rounded transition-colors disabled:bg-slate-300"
                >
                  {saving ? <Loader2 className="w-3 h-3 animate-spin" /> : <Check className="w-3 h-3" />}
                  {saving ? 'Saving...' : 'Confirm Save'}
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
