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
    <div className="min-h-screen bg-gradient-to-br from-rose-50 via-pink-50 to-rose-100">
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
              <h1 className="text-3xl font-bold bg-gradient-to-r from-rose-600 to-pink-600 bg-clip-text text-transparent">GPO Statistics</h1>
              <p className="text-xs text-rose-500 font-medium">Dashboard</p>
            </div>
          </div>
          <Link
            to="/admin"
            className="flex items-center gap-2 px-4 py-2 bg-gradient-to-r from-rose-500 to-pink-500 hover:from-rose-600 hover:to-pink-600 text-white text-xs font-semibold rounded-lg transition-all shadow-soft hover:shadow-glow"
          >
            <Shield className="w-4 h-4" />
            Admin
          </Link>
        </div>
        <div className="mb-8">
          <input
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search GPO by last 4 digits"
            className="w-full max-w-sm px-4 py-3 border-2 border-rose-200 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-rose-400 focus:border-transparent placeholder-rose-300 transition-all"
          />
        </div>

        {loading ? (
          <div className="flex items-center justify-center py-20">
            <Loader2 className="w-8 h-8 text-rose-500 animate-spin" />
          </div>
        ) : error ? (
          <div className="flex items-center gap-2 justify-center py-20 bg-red-50 border-2 border-red-200 rounded-xl p-6">
            <AlertCircle className="w-5 h-5 text-red-500 flex-shrink-0" />
            <p className="text-sm text-red-600 font-medium">{error}</p>
          </div>
        ) : records.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-20 bg-rose-50 rounded-xl border-2 border-rose-200">
            <BarChart3 className="w-12 h-12 mb-3 text-rose-200" />
            <p className="text-sm font-semibold text-rose-400">No records yet</p>
          </div>
        ) : (
          <>
            {/* Table */}
            <div className="bg-white border-2 border-rose-200 rounded-xl overflow-hidden shadow-soft">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-gradient-to-r from-rose-100 to-pink-100 border-b-2 border-rose-200">
                    <th className="px-4 py-3 text-left font-bold text-rose-700">GPO Number</th>
                    <th className="px-4 py-3 text-center font-bold text-rose-700">Interest Rate</th>
                    <th className="px-4 py-3 text-right font-bold text-rose-700">Profit</th>
                    <th className="px-4 py-3 text-right font-bold text-rose-700">Used Limit</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRecords.map((item, idx) => (
                    <tr key={item.id} className={`border-b border-rose-100 transition-colors ${idx % 2 === 0 ? 'bg-white' : 'bg-rose-50'} hover:bg-rose-100`}>
                      <td className="px-4 py-3 bg-gradient-to-r from-rose-100 to-pink-100 text-rose-900 font-semibold">{item.gpo}</td>
                      <td className="px-4 py-3 text-center text-rose-900 font-semibold">
                        {(item.interestRate * 100).toFixed(1)}%
                      </td>
                      <td className="px-4 py-3 text-right text-rose-900 font-semibold">
                        {formatNumber(item.profit)}
                      </td>
                      <td className="px-4 py-3 text-right text-rose-900 font-semibold">
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
