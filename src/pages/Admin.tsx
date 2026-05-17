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
      <div className="min-h-screen bg-gradient-to-br from-rose-50 via-pink-50 to-rose-100 flex items-center justify-center px-4">
        <div className="w-full max-w-sm">
          <div className="bg-white rounded-xl border-2 border-rose-200 p-8 shadow-soft">
            <div className="flex flex-col items-center mb-8">
              <div className="p-3 bg-gradient-to-br from-rose-400 to-pink-500 rounded-xl mb-4 shadow-soft">
                <Lock className="w-6 h-6 text-white" />
              </div>
              <h1 className="text-2xl font-bold bg-gradient-to-r from-rose-600 to-pink-600 bg-clip-text text-transparent">Admin Access</h1>
              <p className="text-sm text-rose-500 mt-2 font-medium">Enter password to continue</p>
            </div>
            <form onSubmit={handleLogin} className="space-y-4">
              <div>
                <input
                  type="password"
                  value={password}
                  onChange={(e) => {
                    setPassword(e.target.value);
                    setPasswordError('');
                  }}
                  placeholder="Password"
                  className="w-full px-4 py-3 border-2 border-rose-200 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-rose-400 focus:border-transparent transition-all placeholder-rose-300"
                  autoFocus
                />
                {passwordError && (
                  <p className="text-sm text-red-600 mt-2 flex items-center gap-1 font-medium">
                    <AlertCircle className="w-4 h-4" />
                    {passwordError}
                  </p>
                )}
              </div>
              <button
                type="submit"
                className="w-full py-3 px-4 bg-gradient-to-r from-rose-500 to-pink-500 hover:from-rose-600 hover:to-pink-600 text-white font-semibold rounded-lg transition-all shadow-soft hover:shadow-glow"
              >
                Unlock
              </button>
            </form>
            <Link
              to="/"
              className="flex items-center justify-center gap-2 text-sm text-rose-600 hover:text-rose-700 font-medium mt-6 transition-colors"
            >
              <ArrowLeft className="w-4 h-4" />
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
    <div className="min-h-screen bg-gradient-to-br from-rose-50 via-pink-50 to-rose-100">
      {/* Toast */}
      {toast && (
        <div
          className={`fixed top-4 right-4 z-50 flex items-center gap-2 px-5 py-3 rounded-lg text-sm font-medium transition-all duration-300 shadow-soft ${
            toast.type === 'success' ? 'bg-emerald-500 text-white' : 'bg-red-500 text-white'
          }`}
        >
          {toast.type === 'success' ? (
            <CheckCircle className="w-5 h-5 shrink-0" />
          ) : (
            <AlertCircle className="w-5 h-5 shrink-0" />
          )}
          {toast.message}
        </div>
      )}

      <div className="max-w-6xl mx-auto px-4 py-8">
        {/* Header */}
        <div className="flex items-center justify-between mb-8">
          <div className="flex items-center gap-3">
            {hasIcon ? (
              <div className="flex items-center">
                <img src="/icon.png" alt="GPO" className="w-12 h-12 object-contain" />
              </div>
            ) : (
              <div className="p-3 bg-gradient-to-br from-rose-400 to-pink-500 rounded-xl shadow-soft">
                <BarChart3 className="w-8 h-8 text-white" />
              </div>
            )}
            <div>
              <p className="text-sm text-rose-500 font-medium">Admin Panel</p>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <button
              onClick={() => setAuthenticated(false)}
              className="flex items-center gap-2 px-4 py-2 bg-red-50 hover:bg-red-100 text-red-600 text-sm font-semibold rounded-lg border-2 border-red-200 transition-all"
            >
              <Lock className="w-4 h-4" />
              Lock
            </button>
            <Link
              to="/"
              className="flex items-center gap-2 px-4 py-2 bg-rose-100 hover:bg-rose-200 text-rose-700 text-sm font-semibold rounded-lg border-2 border-rose-200 transition-all"
            >
              <ArrowLeft className="w-4 h-4" />
              Dashboard
            </Link>
          </div>
        </div>

        {/* Upload Section */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-8">
          <div className="md:col-span-2 bg-white rounded-xl border-2 border-rose-200 p-6 shadow-soft">
            <p className="text-sm font-bold text-rose-600 uppercase tracking-wide mb-3">
              Upload File
            </p>
            <button
              onClick={() => fileInputRef.current?.click()}
              className="w-full flex items-center gap-3 px-4 py-4 border-3 border-dashed border-rose-300 hover:border-rose-500 hover:bg-rose-50 rounded-lg transition-all duration-200 text-left group"
            >
              <div className="p-2 bg-rose-100 group-hover:bg-rose-200 rounded-lg transition-colors flex-shrink-0">
                <FileSpreadsheet className="w-6 h-6 text-rose-500 group-hover:text-rose-600 transition-colors" />
              </div>
              <div className="flex-1">
                {fileName ? (
                  <>
                    <p className="text-sm font-bold text-rose-700">{fileName}</p>
                    <p className="text-xs text-rose-500 font-medium">{uploadedRecords.length} records</p>
                  </>
                ) : (
                  <>
                    <p className="text-sm font-bold text-rose-700">Choose Excel or CSV</p>
                    <p className="text-xs text-rose-500 font-medium">.xlsx, .xls, .csv</p>
                  </>
                )}
              </div>
              <Upload className="w-5 h-5 text-rose-400 group-hover:text-rose-600 transition-colors flex-shrink-0" />
            </button>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={handleFile}
            />
          </div>

          <div className="bg-white rounded-xl border-2 border-rose-200 p-6 shadow-soft flex flex-col justify-between">
            <p className="text-sm font-bold text-rose-600 uppercase tracking-wide mb-3">
              Action
            </p>
            <button
              onClick={openSaveConfirmation}
              disabled={saving || uploadedRecords.length === 0}
              className="flex items-center justify-center gap-2 w-full py-3 px-4 bg-gradient-to-r from-rose-500 to-pink-500 hover:from-rose-600 hover:to-pink-600 disabled:from-rose-300 disabled:to-pink-300 text-white font-bold rounded-lg transition-all duration-200 shadow-soft hover:shadow-glow"
            >
              {saving ? <Loader2 className="w-4 h-4 animate-spin" /> : <Save className="w-4 h-4" />}
              {saving ? 'Saving...' : 'Review & Save'}
            </button>
          </div>
        </div>

        {/* Upload Preview Summary */}
        {uploadedRecords.length > 0 && (
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-8">
            <div className="bg-white border-2 border-rose-200 rounded-xl p-4 shadow-soft">
              <p className="text-xs font-bold text-rose-600 uppercase tracking-wide mb-2">Records to Upload</p>
              <p className="text-3xl font-bold bg-gradient-to-r from-rose-600 to-pink-600 bg-clip-text text-transparent">{uploadedRecords.length}</p>
            </div>
            <div className="bg-white border-2 border-rose-200 rounded-xl p-4 shadow-soft">
              <p className="text-xs font-bold text-rose-600 uppercase tracking-wide mb-2">Upload Profit</p>
              <p className="text-3xl font-bold bg-gradient-to-r from-rose-600 to-pink-600 bg-clip-text text-transparent">{formatNumber(uploadProfit)}</p>
            </div>
            <div className="bg-white border-2 border-rose-200 rounded-xl p-4 shadow-soft">
              <p className="text-xs font-bold text-rose-600 uppercase tracking-wide mb-2">Upload Used Limit</p>
              <p className="text-3xl font-bold bg-gradient-to-r from-rose-600 to-pink-600 bg-clip-text text-transparent">{formatNumber(uploadLimit)}</p>
            </div>
          </div>
        )}

        <div className="mb-6">
          <input
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search GPO by last 4 digits"
            className="w-full max-w-sm px-4 py-3 border-2 border-rose-200 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-rose-400 focus:border-transparent placeholder-rose-300 transition-all"
          />
        </div>

        {/* Saved Records Table */}
        <div className="bg-white rounded-xl border-2 border-rose-200 overflow-hidden mb-8 shadow-soft">
          <div className="px-6 py-4 border-b-2 border-rose-200 bg-gradient-to-r from-rose-50 to-pink-50 flex items-center justify-between">
            <h2 className="text-sm font-bold text-rose-700 uppercase tracking-wide">
              Saved Records ({records.length})
            </h2>
            <button
              onClick={loadRecords}
              disabled={loading}
              className="text-sm text-rose-600 hover:text-rose-700 font-bold transition-colors"
            >
              {loading ? 'Loading...' : 'Refresh'}
            </button>
          </div>

          {loading && records.length === 0 ? (
            <div className="flex items-center justify-center py-16">
              <Loader2 className="w-8 h-8 text-rose-500 animate-spin" />
            </div>
          ) : records.length === 0 ? (
            <div className="flex flex-col items-center justify-center py-16 text-rose-300">
              <FileSpreadsheet className="w-12 h-12 mb-3 text-rose-200" />
              <p className="text-sm font-bold text-rose-400">No saved records</p>
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-gradient-to-r from-rose-100 to-pink-100 border-b-2 border-rose-200">
                    <th className="px-4 py-3 text-left font-bold text-rose-700 w-10">#</th>
                    <th className="px-4 py-3 text-left font-bold text-rose-700">GPO Number</th>
                    <th className="px-4 py-3 text-center font-bold text-rose-700">Interest Rate</th>
                    <th className="px-4 py-3 text-right font-bold text-rose-700">Profit</th>
                    <th className="px-4 py-3 text-right font-bold text-rose-700">Used Limit</th>
                    <th className="px-4 py-3 text-center font-bold text-rose-700 w-16">Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRecords.map((item, i) => {
                    const isEditingThis = editState?.id === item.id;
                    return (
                      <tr key={item.id} className={`border-b border-rose-100 transition-colors ${i % 2 === 0 ? 'bg-white' : 'bg-rose-50'} hover:bg-rose-100 group`}>
                        <td className="px-4 py-3 text-rose-500 text-sm font-medium">{i + 1}</td>
                        <td className="px-4 py-3 bg-gradient-to-r from-rose-100 to-pink-100">
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
                              className="w-full px-2 py-1 text-sm border-2 border-rose-300 rounded focus:outline-none focus:ring-2 focus:ring-rose-500"
                            />
                          ) : (
                            <span
                              className="text-rose-900 font-bold cursor-pointer hover:underline"
                              onClick={() => startEdit(item.id, 'gpo', item.gpo)}
                            >
                              {item.gpo}
                            </span>
                          )}
                        </td>
                        <td className="px-4 py-3 text-center">
                          <span className="text-rose-900 font-bold">
                            {(item.interestRate * 100).toFixed(1)}%
                          </span>
                        </td>
                        <td className="px-4 py-3 text-right">
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
                              className="w-32 px-2 py-1 text-sm text-right border-2 border-rose-300 rounded focus:outline-none focus:ring-2 focus:ring-rose-500"
                            />
                          ) : (
                            <span
                              className="text-rose-900 font-bold cursor-pointer hover:underline"
                              onClick={() => startEdit(item.id, 'profit', item.profit)}
                            >
                              {formatNumber(item.profit)}
                            </span>
                          )}
                        </td>
                        <td className="px-4 py-3 text-right">
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
                              className="w-40 px-2 py-1 text-sm text-right border-2 border-rose-300 rounded focus:outline-none focus:ring-2 focus:ring-rose-500"
                            />
                          ) : (
                            <span
                              className="text-rose-900 font-bold cursor-pointer hover:underline"
                              onClick={() => startEdit(item.id, 'usedLimit', item.usedLimit)}
                            >
                              {formatNumber(item.usedLimit)}
                            </span>
                          )}
                        </td>
                        <td className="px-4 py-3 text-center">
                          {isEditingThis ? (
                            <div className="flex items-center justify-center gap-1">
                              <button
                                onClick={requestCommitEdit}
                                className="p-2 rounded bg-emerald-100 text-emerald-600 hover:bg-emerald-200 transition-colors"
                              >
                                <Check className="w-4 h-4" />
                              </button>
                              <button
                                onClick={cancelEdit}
                                className="p-2 rounded bg-rose-100 text-rose-500 hover:bg-rose-200 transition-colors"
                              >
                                <X className="w-4 h-4" />
                              </button>
                            </div>
                          ) : (
                            <div className="flex items-center justify-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                              <button
                                onClick={() => startEdit(item.id, 'gpo', item.gpo)}
                                className="p-2 rounded bg-rose-100 text-rose-600 hover:bg-rose-200 transition-colors"
                              >
                                <Pencil className="w-4 h-4" />
                              </button>
                              <button
                                onClick={() => setDeleteConfirm(item.id)}
                                className="p-2 rounded bg-red-100 text-red-600 hover:bg-red-200 transition-colors"
                              >
                                <Trash2 className="w-4 h-4" />
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
                    <tr className="bg-gradient-to-r from-rose-100 to-pink-100 border-t-2 border-rose-200">
                      <td colSpan={2} className="px-4 py-3 text-sm font-bold text-rose-700 uppercase tracking-wide">
                        Totals
                      </td>
                      <td className="px-4 py-3" />
                      <td className="px-4 py-3 text-right font-bold text-rose-900">
                        {formatNumber(totalProfit)}
                      </td>
                      <td className="px-4 py-3 text-right font-bold text-rose-900">
                        {formatNumber(totalLimit)}
                      </td>
                      <td className="px-4 py-3" />
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>
          )}
        </div>

        {/* Preview Modal */}
        {preview.isOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
            <div className="bg-white rounded-2xl shadow-2xl max-w-2xl w-full max-h-[80vh] flex flex-col border-2 border-rose-200">
              <div className="px-6 py-4 border-b-2 border-rose-200 bg-gradient-to-r from-rose-50 to-pink-50 flex items-center justify-between">
                <h2 className="text-lg font-bold bg-gradient-to-r from-rose-600 to-pink-600 bg-clip-text text-transparent">Preview Upload Data</h2>
                <button
                  onClick={() => setPreview({ isOpen: false, data: [] })}
                  className="text-rose-400 hover:text-rose-600 transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="flex-1 overflow-auto">
                <table className="w-full text-sm border-collapse">
                  <thead className="sticky top-0 bg-gradient-to-r from-rose-100 to-pink-100 border-b-2 border-rose-200">
                    <tr>
                      <th className="px-4 py-3 text-left font-bold text-rose-700 w-10">#</th>
                      <th className="px-4 py-3 text-left font-bold text-rose-700">GPO Number</th>
                      <th className="px-4 py-3 text-center font-bold text-rose-700">Interest Rate</th>
                      <th className="px-4 py-3 text-right font-bold text-rose-700">Profit</th>
                      <th className="px-4 py-3 text-right font-bold text-rose-700">Used Limit</th>
                    </tr>
                  </thead>
                  <tbody>
                    {preview.data.map((item, i) => (
                      <tr key={i} className={`border-b border-rose-100 ${i % 2 === 0 ? 'bg-white' : 'bg-rose-50'} hover:bg-rose-100`}>
                        <td className="px-4 py-3 text-rose-500 font-medium">{i + 1}</td>
                        <td className="px-4 py-3 bg-gradient-to-r from-rose-100 to-pink-100 font-bold text-rose-900">{item.gpo}</td>
                        <td className="px-4 py-3 text-center font-bold text-rose-900">
                          {(item.interestRate * 100).toFixed(1)}%
                        </td>
                        <td className="px-4 py-3 text-right font-bold text-rose-900">
                          {formatNumber(item.profit)}
                        </td>
                        <td className="px-4 py-3 text-right font-bold text-rose-900">
                          {formatNumber(item.usedLimit)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="px-6 py-4 border-t-2 border-rose-200 bg-gradient-to-r from-rose-50 to-pink-50 flex items-center justify-between gap-3">
                <p className="text-sm text-rose-600 font-bold">
                  <span className="font-bold">{preview.data.length}</span> records to upload
                </p>
                <div className="flex gap-3">
                  <button
                    onClick={() => setPreview({ isOpen: false, data: [] })}
                    className="px-4 py-2 text-sm font-bold text-rose-700 bg-rose-100 hover:bg-rose-200 border-2 border-rose-200 rounded-lg transition-all"
                  >
                    Back
                  </button>
                  <button
                    onClick={() => {
                      setPreview({ isOpen: false, data: [] });
                      openSaveConfirmation();
                    }}
                    className="flex items-center gap-2 px-4 py-2 text-sm font-bold text-white bg-gradient-to-r from-rose-500 to-pink-500 hover:from-rose-600 hover:to-pink-600 rounded-lg transition-all shadow-soft hover:shadow-glow"
                  >
                    <Check className="w-4 h-4" />
                    Proceed to Confirm
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {pendingEditConfirm && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
            <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full border-2 border-rose-200">
              <div className="px-6 py-4 border-b-2 border-rose-200 bg-gradient-to-r from-rose-50 to-pink-50">
                <h2 className="text-lg font-bold text-rose-700 flex items-center gap-2">
                  <AlertCircle className="w-5 h-5 text-amber-500" />
                  Confirm Edit
                </h2>
              </div>
              <div className="px-6 py-5 space-y-3">
                <p className="text-sm text-rose-700">
                  Are you sure you want to update <span className="font-bold text-rose-900">{pendingEditConfirm.field}</span> for this record?
                </p>
                <div className="bg-rose-50 border-2 border-rose-200 rounded-lg p-3">
                  <p className="text-sm font-bold text-rose-900">New value: {pendingEditConfirm.value}</p>
                </div>
              </div>
              <div className="px-6 py-4 border-t-2 border-rose-200 bg-gradient-to-r from-rose-50 to-pink-50 flex gap-3 justify-end">
                <button
                  onClick={() => setPendingEditConfirm(null)}
                  className="px-4 py-2 text-sm font-bold text-rose-700 bg-rose-100 hover:bg-rose-200 border-2 border-rose-200 rounded-lg transition-all"
                >
                  Cancel
                </button>
                <button
                  onClick={confirmEdit}
                  className="flex items-center gap-2 px-4 py-2 text-sm font-bold text-white bg-emerald-500 hover:bg-emerald-600 rounded-lg transition-all shadow-soft"
                >
                  <Check className="w-4 h-4" />
                  Confirm Edit
                </button>
              </div>
            </div>
          </div>
        )}

        {deleteConfirm && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
            <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full border-2 border-red-200">
              <div className="px-6 py-4 border-b-2 border-red-200 bg-red-50">
                <h2 className="text-lg font-bold text-red-700 flex items-center gap-2">
                  <AlertCircle className="w-5 h-5" />
                  Confirm Delete
                </h2>
              </div>
              <div className="px-6 py-5 space-y-3">
                <p className="text-sm text-red-700">
                  Are you sure you want to delete this record? This action cannot be undone.
                </p>
              </div>
              <div className="px-6 py-4 border-t-2 border-red-200 bg-red-50 flex gap-3 justify-end">
                <button
                  onClick={() => setDeleteConfirm(null)}
                  className="px-4 py-2 text-sm font-bold text-red-700 bg-red-100 hover:bg-red-200 border-2 border-red-200 rounded-lg transition-all"
                >
                  Cancel
                </button>
                <button
                  onClick={() => confirmDelete(deleteConfirm)}
                  className="flex items-center gap-2 px-4 py-2 text-sm font-bold text-white bg-red-500 hover:bg-red-600 rounded-lg transition-all shadow-soft"
                >
                  <Trash2 className="w-4 h-4" />
                  Delete
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Confirmation Modal */}
        {confirmSave.isOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4 backdrop-blur-sm">
            <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full border-2 border-rose-200">
              <div className="px-6 py-4 border-b-2 border-rose-200 bg-gradient-to-r from-rose-50 to-pink-50">
                <h2 className="text-lg font-bold text-rose-700 flex items-center gap-2">
                  <AlertCircle className="w-5 h-5 text-amber-500" />
                  Confirm Save
                </h2>
              </div>

              <div className="px-6 py-5 space-y-3">
                <p className="text-sm text-rose-700">
                  Are you sure you want to save <span className="font-bold text-rose-900">{confirmSave.data.length}</span> records to the database?
                </p>
                <div className="bg-amber-50 border-2 border-amber-300 rounded-lg p-4">
                  <p className="text-sm font-bold text-amber-900">
                    This action cannot be easily undone. Please review the data carefully before confirming.
                  </p>
                </div>
              </div>

              <div className="px-6 py-4 border-t-2 border-rose-200 bg-gradient-to-r from-rose-50 to-pink-50 flex gap-3 justify-end">
                <button
                  onClick={() => setConfirmSave({ isOpen: false, data: [] })}
                  disabled={saving}
                  className="px-4 py-2 text-sm font-bold text-rose-700 bg-rose-100 hover:bg-rose-200 border-2 border-rose-200 rounded-lg transition-all disabled:opacity-50"
                >
                  Cancel
                </button>
                <button
                  onClick={handleConfirmSave}
                  disabled={saving}
                  className="flex items-center gap-2 px-4 py-2 text-sm font-bold text-white bg-emerald-500 hover:bg-emerald-600 rounded-lg transition-all shadow-soft disabled:bg-emerald-300"
                >
                  {saving ? <Loader2 className="w-4 h-4 animate-spin" /> : <Check className="w-4 h-4" />}
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
