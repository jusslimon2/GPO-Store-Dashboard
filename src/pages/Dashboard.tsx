import { useState, useEffect } from 'react';
import { Link } from 'react-router-dom';
import {
  BarChart3,
  TrendingUp,
  DollarSign,
  Hash,
  Shield,
  Loader2,
  AlertCircle,
} from 'lucide-react';
import { fetchAllRecords, type GpoDoc } from '../lib/firebase';

function formatNumber(n: number): string {
  return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 3 });
}

export default function Dashboard() {
  const [records, setRecords] = useState<GpoDoc[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [hasIcon, setHasIcon] = useState(false);

  useEffect(() => {
    fetchAllRecords()
      .then((data) => setRecords(data))
      .catch((err) => setError(err.message))
      .finally(() => setLoading(false));
  }, []);

  const totalProfit = records.reduce((s, r) => s + r.profit, 0);
  const totalLimit = records.reduce((s, r) => s + r.usedLimit, 0);
  const uniqueBatches = new Set(records.map((r) => r.batchId)).size;

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

  return (
    <div className="min-h-screen bg-white">
      <div className="max-w-6xl mx-auto px-4 py-6">
        {/* Header */}
        <div className="flex items-center justify-between mb-6">
          <div className="flex items-center gap-3">
            {hasIcon ? (
              <div className="flex items-center">
                <img src="/icon.png" alt="GPO" className="w-6 h-6 object-contain" />
              </div>
            ) : (
              <div className="p-2 bg-slate-800 rounded-lg">
                <BarChart3 className="w-5 h-5 text-white" />
              </div>
            )}
            <div>
              <h1 className="text-xl font-bold text-slate-800">GPO Statistics</h1>
            </div>
          </div>
          <Link
            to="/admin"
            className="flex items-center gap-2 px-3 py-1.5 bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold rounded transition-colors"
          >
            <Shield className="w-3.5 h-3.5" />
            Admin
          </Link>
        </div>
        <div className="mb-6">
          <input
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search GPO by last 4 digits"
            className="w-full max-w-sm px-3 py-2 border border-slate-300 rounded text-xs focus:outline-none focus:ring-1 focus:ring-blue-500"
          />
        </div>

        {loading ? (
          <div className="flex items-center justify-center py-16">
            <Loader2 className="w-6 h-6 text-blue-500 animate-spin" />
          </div>
        ) : error ? (
          <div className="flex items-center gap-2 justify-center py-16 text-red-600">
            <AlertCircle className="w-4 h-4" />
            <p className="text-xs">{error}</p>
          </div>
        ) : records.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-16 text-slate-400">
            <BarChart3 className="w-10 h-10 mb-2 text-slate-200" />
            <p className="text-xs font-medium">No records yet</p>
          </div>
        ) : (
          <>
            {/* Table */}
            <div className="bg-white border border-slate-200 rounded overflow-hidden">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-slate-200 border-b border-slate-300">
                    <th className="px-3 py-2 text-left font-bold text-slate-700">GPO Number</th>
                    <th className="px-3 py-2 text-center font-bold text-slate-700">Interest Rate</th>
                    <th className="px-3 py-2 text-right font-bold text-slate-700">Profit</th>
                    <th className="px-3 py-2 text-right font-bold text-slate-700">Used Limit</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRecords.map((item) => (
                    <tr key={item.id} className="border-b border-slate-100">
                      <td className="px-3 py-2 bg-blue-100 text-slate-800 font-medium">{item.gpo}</td>
                      <td className="px-3 py-2 bg-pink-100 text-center text-slate-800 font-medium">
                        {(item.interestRate * 100).toFixed(1)}%
                      </td>
                      <td className="px-3 py-2 bg-pink-100 text-right text-slate-800 font-medium">
                        {formatNumber(item.profit)}
                      </td>
                      <td className="px-3 py-2 bg-green-100 text-right text-slate-800 font-medium">
                        {formatNumber(item.usedLimit)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
