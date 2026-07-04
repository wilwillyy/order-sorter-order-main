import logoUrl from './logo.png';
import React, { useEffect, useState } from 'react';
import type { ChangeEvent, ReactNode } from 'react';
import * as XLSX from 'xlsx';
import {
  AlertTriangle,
  Archive,
  BarChart3,
  CheckCircle,
  History,
  LayoutDashboard,
  ListChecks,
  Menu,
  Package,
  PackageOpen,
  Printer,
  Search,
  Trash2,
  TrendingUp,
  Upload,
  X,
} from 'lucide-react';

type TabId = 'dashboard' | 'import' | 'picking' | 'logs';
type ToastType = 'info' | 'success' | 'error';
type OrderStatus = 'Pending' | 'Ready to Shipment' | 'Completed';
type TimeFilter = 'all' | 'day' | 'week' | 'month';

type ParsedItem = {
  product: string;
  color: string;
  qty: number;
  rawText: string;
};

type Order = {
  id: string;
  marketplace: string;
  shippingMethod: string;
  date: string;
  items: ParsedItem[];
  status: OrderStatus;
};

type ToastState = {
  message: string;
  type: ToastType;
} | null;

type LogEntry = {
  time: string;
  action: string;
};

type ProducedItemsMap = Record<string, boolean>;

type ConfirmDialogState = {
  isOpen: boolean;
  orderId: string | null;
};

type AggregatedProduct = {
  id: string;
  product: string;
  color: string;
  qty: number;
};

type NavItemProps = {
  id: TabId;
  label: string;
  icon: ReactNode;
};

const STORAGE_KEYS = {
  orders: 'wms_orders',
  history: 'wms_history',
  producedItems: 'wms_producedItems',
  logs: 'wms_logs',
} as const;

const isBrowser = typeof window !== 'undefined';

const getStorage = <T,>(key: string, defaultValue: T): T => {
  if (!isBrowser) return defaultValue;

  try {
    const saved = window.localStorage.getItem(key);
    return saved !== null ? (JSON.parse(saved) as T) : defaultValue;
  } catch (error) {
    console.error(`Gagal membaca ${key} dari LocalStorage:`, error);
    return defaultValue;
  }
};

const setStorage = <T,>(key: string, value: T) => {
  if (!isBrowser) return;

  try {
    window.localStorage.setItem(key, JSON.stringify(value));
  } catch (error) {
    console.error(`Gagal menyimpan ${key} ke LocalStorage:`, error);
  }
};

const removeStorage = (key: string) => {
  if (!isBrowser) return;

  try {
    window.localStorage.removeItem(key);
  } catch (error) {
    console.error(`Gagal menghapus ${key} dari LocalStorage:`, error);
  }
};

export default function App() {
  const [currentTab, setCurrentTab] = useState<TabId>('dashboard');
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [orders, setOrders] = useState<Order[]>([]);
  const [historyOrders, setHistoryOrders] = useState<Order[]>([]);
  const [producedItems, setProducedItems] = useState<ProducedItemsMap>({});
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [loading, setLoading] = useState(false);
  const [toast, setToast] = useState<ToastState>(null);
  const [hasHydrated, setHasHydrated] = useState(false);

  useEffect(() => {
    setOrders(getStorage<Order[]>(STORAGE_KEYS.orders, []));
    setHistoryOrders(getStorage<Order[]>(STORAGE_KEYS.history, []));
    setProducedItems(getStorage<ProducedItemsMap>(STORAGE_KEYS.producedItems, {}));
    setLogs(getStorage<LogEntry[]>(STORAGE_KEYS.logs, []));
    setHasHydrated(true);
  }, []);

  useEffect(() => {
    if (!hasHydrated) return;
    setStorage(STORAGE_KEYS.orders, orders);
  }, [hasHydrated, orders]);

  useEffect(() => {
    if (!hasHydrated) return;
    setStorage(STORAGE_KEYS.history, historyOrders);
  }, [hasHydrated, historyOrders]);

  useEffect(() => {
    if (!hasHydrated) return;
    setStorage(STORAGE_KEYS.producedItems, producedItems);
  }, [hasHydrated, producedItems]);

  useEffect(() => {
    if (!hasHydrated) return;
    setStorage(STORAGE_KEYS.logs, logs);
  }, [hasHydrated, logs]);

  const showToast = (message: string, type: ToastType = 'info') => {
    setToast({ message, type });
    window.setTimeout(() => setToast(null), 3000);
  };

  const logActivity = (action: string) => {
    const time = new Date().toLocaleTimeString();
    setLogs(prev => [{ time, action }, ...prev].slice(0, 50));
  };

  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const input = e.target;

    setLoading(true);
    const reader = new FileReader();

    reader.onload = evt => {
      try {
        const buffer = evt.target?.result;
        if (!(buffer instanceof ArrayBuffer)) {
          throw new Error('Format file tidak valid.');
        }

        const data = new Uint8Array(buffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, { defval: '' });

        const orderMap = new Map<string, Order>();

        jsonData.forEach(row => {
          const getVal = (keys: string[]) => {
            const foundKey = Object.keys(row).find(k => keys.includes(k.toLowerCase().trim()));
            return foundKey ? String(row[foundKey]) : '';
          };

          const trackingNumber = getVal(['tracking_number', 'resi']);
          const shopeeOrderSn = getVal(['order_sn']);
          const shopeeShipping = getVal(['shipping_method']);
          const productInfo = getVal(['product_info']);

          const orderId = getVal(['order id']);
          const packageId = getVal(['package id']);
          const productName = getVal(['product name']);
          const variation = getVal(['variation']);
          const tiktokQty = getVal(['quantity']);
          const tiktokShipping = getVal(['shipping provider name']);

          const dateStr = getVal(['waktu pesanan dibuat', 'order creation date', 'created time', 'waktu dibuat']);
          let orderDate = new Date().toISOString();
          if (dateStr) {
            const d = new Date(dateStr);
            if (!Number.isNaN(d.getTime())) orderDate = d.toISOString();
          }

          const finalId = shopeeOrderSn || packageId || orderId;
          if (!finalId) return; 

          let marketplace = 'Lainnya';
          const digitOnlyOrderId = orderId.replace(/\D/g, '');
          
          if (trackingNumber.toUpperCase().includes('SPX') || shopeeOrderSn) {
            marketplace = 'Shopee';
          } else if ((digitOnlyOrderId.length >= 18 || packageId) && !trackingNumber.toUpperCase().includes('SPX')) {
            marketplace = 'TikTok Shop';
          }

          const parsedItems: ParsedItem[] = [];
          if (marketplace === 'Shopee') {
            if (productInfo.includes('[1]')) {
              const itemBlocks = productInfo.split(/\[\d+\]/).filter(Boolean);
              itemBlocks.forEach((block: string) => {
                const nameMatch = block.match(/Nama Produk:\s*(.*?)\s*;/);
                const varMatch = block.match(/Nama Variasi:\s*(.*?)\s*;/);
                const qtyMatch = block.match(/Jumlah:\s*(\d+)\s*;/);
                
                parsedItems.push({
                  product: nameMatch ? nameMatch[1].trim() : 'Produk',
                  color: varMatch ? varMatch[1].trim() : '-',
                  qty: qtyMatch ? parseInt(qtyMatch[1], 10) : 1,
                  rawText: block.trim()
                });
              });
            } else if (productInfo) {
              parsedItems.push({
                product: productInfo || 'Produk Tidak Diketahui',
                color: getVal(['variasi warna', 'warna', 'variasi']) || '-',
                qty: parseInt(getVal(['quantity', 'jumlah', 'qty']) || '1', 10),
                rawText: productInfo
              });
            }
          } else if (marketplace === 'TikTok Shop') {
            parsedItems.push({
              product: productName || 'Produk Tidak Diketahui',
              color: variation || '-',
              qty: parseInt(tiktokQty || '1', 10),
              rawText: `${productName} - ${variation}`
            });
          }

          if (orderMap.has(finalId)) {
            const existingOrder = orderMap.get(finalId);
            if (existingOrder) {
              existingOrder.items.push(...parsedItems);
            }
          } else {
            orderMap.set(finalId, {
              id: finalId,
              marketplace: marketplace,
              shippingMethod: marketplace === 'Shopee' ? shopeeShipping : tiktokShipping,
              date: orderDate,
              items: parsedItems,
              status: 'Pending'
            });
          }
        });

        const newOrders = Array.from(orderMap.values());
        const existingIds = new Set(orders.map(order => order.id));
        const uniqueNewOrders = newOrders.filter(order => !existingIds.has(order.id));
        const duplicateCount = newOrders.length - uniqueNewOrders.length;

        if (uniqueNewOrders.length > 0) {
          setOrders(prev => [...prev, ...uniqueNewOrders]);
          setHistoryOrders(prev => [...prev, ...uniqueNewOrders]);

          const importMessage = duplicateCount > 0
            ? `Berhasil import ${uniqueNewOrders.length} pesanan. ${duplicateCount} duplikat dilewati.`
            : `Berhasil import ${uniqueNewOrders.length} pesanan!`;

          showToast(importMessage, 'success');
          logActivity(`Import Excel dengan ${uniqueNewOrders.length} pesanan baru`);
          setCurrentTab('picking');
        } else if (duplicateCount > 0) {
          showToast(`Semua pesanan di file ini sudah pernah di-import (${duplicateCount} duplikat).`, 'info');
        } else {
          showToast('Tidak ada data valid. Pastikan file memiliki kolom yang benar.', 'error');
        }
      } catch (error) {
        console.error(error);
        showToast('Gagal membaca file. Pastikan format CSV/Excel benar.', 'error');
      } finally {
        setLoading(false);
        input.value = '';
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const DashboardView = () => {
    const [resetDialog, setResetDialog] = useState(false);

    const totalOrders = orders.length;
    const shopeeOrders = orders.filter(o => o.marketplace === 'Shopee').length;
    const tiktokOrders = orders.filter(o => o.marketplace === 'TikTok Shop').length;
    const pendingOrders = orders.filter(o => o.status === 'Pending').length;
    const readyOrders = orders.filter(o => o.status === 'Ready to Shipment').length;
    const completedOrders = orders.filter(o => o.status === 'Completed').length;
    const progress = totalOrders === 0 ? 0 : Math.round((completedOrders / totalOrders) * 100);

    const productMap = new Map<string, AggregatedProduct>();
    orders.forEach(order => {
      // Abaikan pesanan yang sudah Completed agar Master Picking selalu update sesuai antrean
      if (order.status === 'Completed') return; 

      order.items.forEach(item => {
        const prodName = item.product || 'Produk Tidak Diketahui';
        const color = item.color || '-';
        const key = `${prodName.trim().toLowerCase()}|${color.trim().toLowerCase()}`;
        
        if (productMap.has(key)) {
          const existingProduct = productMap.get(key);
          if (existingProduct) {
            existingProduct.qty += item.qty;
          }
        } else {
          productMap.set(key, {
            id: key,
            product: prodName.trim(),
            color: color.trim() !== '' ? color.trim() : '-',
            qty: item.qty
          });
        }
      });
    });

    const aggregatedProducts = Array.from(productMap.values()).sort((a, b) => b.qty - a.qty);

    const toggleProduce = (itemId: string) => {
      setProducedItems(prev => ({ ...prev, [itemId]: !prev[itemId] }));
    };

    const confirmResetData = () => {
      setOrders([]);
      setProducedItems({});
      // HANYA menghapus antrean berjalan, History dan Log tetap aman!
      removeStorage(STORAGE_KEYS.orders);
      removeStorage(STORAGE_KEYS.producedItems);
      setResetDialog(false);
      showToast('Antrean Pesanan berhasil direset!', 'success');
      logActivity('Mereset daftar antrean aktif');
    };

    return (
      <div className="space-y-8 animate-in fade-in zoom-in-95 duration-500">
        
        {resetDialog && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in">
            <div className="bg-white dark:bg-slate-800 rounded-2xl p-6 md:p-8 max-w-md w-full shadow-2xl border border-rose-200 dark:border-rose-900/50 zoom-in-95 animate-in">
              <div className="flex items-center gap-4 text-rose-500 mb-4">
                <AlertTriangle size={32} />
                <h3 className="text-xl font-bold text-slate-800 dark:text-white">Reset Antrean Pesanan?</h3>
              </div>
              <p className="text-slate-600 dark:text-slate-300 mb-8 font-medium">
                Apakah Anda yakin ingin menghapus semua daftar <strong>Antrean Aktif</strong> dan <strong>Master Picking</strong>? <br/><br/>
                <span className="text-sm text-emerald-600 dark:text-emerald-400 bg-emerald-50 dark:bg-emerald-900/30 px-3 py-2 rounded-lg block border border-emerald-200 dark:border-emerald-800/50">
                  Tenang saja, Data Statistik dan Log Aktivitas Anda tidak akan hilang.
                </span>
              </p>
              <div className="flex justify-end gap-3">
                <button
                  onClick={() => setResetDialog(false)}
                  className="px-5 py-2.5 rounded-xl font-bold text-slate-600 bg-slate-100 hover:bg-slate-200 dark:bg-slate-700 dark:text-slate-300 dark:hover:bg-slate-600 transition-colors"
                >
                  Batal
                </button>
                <button
                  onClick={confirmResetData}
                  className="px-5 py-2.5 rounded-xl font-bold text-white bg-rose-500 hover:bg-rose-600 shadow-lg shadow-rose-500/30 transition-all active:scale-95"
                >
                  Ya, Bersihkan Antrean
                </button>
              </div>
            </div>
          </div>
        )}

        <div>
            <h2 className="text-3xl font-extrabold text-slate-800 dark:text-white tracking-tight">Dashboard Utama</h2>
            <p className="text-slate-500 mt-2 text-lg">Pantau ringkasan pesanan dan kebutuhan produksi antrean Anda.</p>
        </div>
        
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 md:gap-8">
          <div className="bg-gradient-to-br from-blue-500 to-indigo-600 p-8 md:p-10 rounded-[2rem] shadow-xl shadow-blue-500/20 text-white flex flex-col justify-between relative overflow-hidden group hover:-translate-y-1 transition-transform">
            <Package className="absolute -right-6 -top-6 w-40 h-40 text-white opacity-10 group-hover:scale-110 transition-transform duration-700 ease-out" />
            <div>
              <div className="flex items-center gap-3 mb-4 opacity-90">
                <Package size={28} /> <h3 className="font-semibold text-xl tracking-wide uppercase">Total Pesanan Aktif</h3>
              </div>
              <p className="text-7xl font-black drop-shadow-md">{totalOrders}</p>
            </div>
            <div className="mt-10 flex flex-wrap gap-3 text-sm font-bold">
              <span className="bg-white/20 backdrop-blur-md px-6 py-3 rounded-xl flex-1 text-center shadow-inner border border-white/10 text-lg">Shopee: {shopeeOrders}</span>
              <span className="bg-white/20 backdrop-blur-md px-6 py-3 rounded-xl flex-1 text-center shadow-inner border border-white/10 text-lg">TikTok: {tiktokOrders}</span>
            </div>
          </div>

          <div className="bg-gradient-to-br from-orange-400 to-rose-500 p-8 md:p-10 rounded-[2rem] shadow-xl shadow-orange-500/20 text-white flex flex-col justify-between relative overflow-hidden group hover:-translate-y-1 transition-transform">
            <PackageOpen className="absolute -right-6 -top-6 w-40 h-40 text-white opacity-10 group-hover:scale-110 transition-transform duration-700 ease-out" />
            <div>
              <div className="flex items-center gap-3 mb-4 opacity-90">
                <PackageOpen size={28} /> <h3 className="font-semibold text-xl tracking-wide uppercase">Belum di Packing</h3>
              </div>
              <p className="text-7xl font-black drop-shadow-md">{pendingOrders}</p>
            </div>
            <div className="mt-10">
              <span className="bg-white/20 backdrop-blur-md px-6 py-3 rounded-xl inline-block w-full text-center shadow-inner border border-white/10 text-base font-semibold">
                Pesanan status PENDING di Antrean
              </span>
            </div>
          </div>

          <div className="bg-gradient-to-br from-emerald-400 to-teal-500 p-8 md:p-10 rounded-[2rem] shadow-xl shadow-emerald-500/20 text-white flex flex-col justify-between relative overflow-hidden group hover:-translate-y-1 transition-transform">
            <CheckCircle className="absolute -right-6 -top-6 w-40 h-40 text-white opacity-10 group-hover:scale-110 transition-transform duration-700 ease-out" />
            <div>
              <div className="flex items-center gap-3 mb-4 opacity-90">
                <CheckCircle size={28} /> <h3 className="font-semibold text-xl tracking-wide uppercase">Siap Dikirim Ke Kurir</h3>
              </div>
              <p className="text-7xl font-black drop-shadow-md">{readyOrders}</p>
            </div>
            <button 
              onClick={() => {
                if(readyOrders === 0) return showToast('Tidak ada pesanan siap kirim', 'error');
                setOrders(prev => prev.map(o => o.status === 'Ready to Shipment' ? { ...o, status: 'Completed' } : o));
                
                // Update history juga agar statusnya valid
                setHistoryOrders(prev => prev.map(o => o.status === 'Ready to Shipment' ? { ...o, status: 'Completed' } : o));

                logActivity(`Menyerahkan ${readyOrders} pesanan ke kurir`);
                showToast('Pesanan diserahkan ke kurir!', 'success');
              }}
              className="mt-10 w-full bg-white hover:bg-slate-50 text-teal-700 py-4 rounded-xl text-lg font-bold transition-all shadow-lg hover:shadow-xl active:scale-95 border-2 border-transparent hover:border-teal-200">
              Tandai Selesai Dikirim
            </button>
          </div>

          <div className="bg-gradient-to-br from-fuchsia-500 to-purple-600 p-8 md:p-10 rounded-[2rem] shadow-xl shadow-purple-500/20 text-white flex flex-col justify-between relative overflow-hidden group hover:-translate-y-1 transition-transform">
            <Archive className="absolute -right-6 -top-6 w-40 h-40 text-white opacity-10 group-hover:scale-110 transition-transform duration-700 ease-out" />
            <div>
              <div className="flex items-center gap-3 mb-4 opacity-90">
                <Archive size={28} /> <h3 className="font-semibold text-xl tracking-wide uppercase">Progress Harian</h3>
              </div>
              <div className="flex items-end gap-2">
                <p className="text-7xl font-black drop-shadow-md">{progress}%</p>
                <p className="text-xl opacity-80 mb-2 font-medium">Selesai</p>
              </div>
            </div>
            <div className="w-full bg-white/20 h-6 mt-10 rounded-full overflow-hidden backdrop-blur-md p-1 shadow-inner">
              <div className="bg-white h-full rounded-full transition-all duration-1000 ease-out relative shadow-sm" style={{ width: `${progress}%` }}>
                {progress > 0 && <div className="absolute inset-0 bg-white/50 animate-pulse rounded-full"></div>}
              </div>
            </div>
          </div>
        </div>

        <div className="mt-8 bg-white dark:bg-[#0f172a] rounded-[2rem] p-8 md:p-10 shadow-xl shadow-slate-200/50 dark:shadow-none border border-slate-200 dark:border-slate-800">
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-8 gap-4">
            <div className="flex items-center gap-4">
              <div className="p-4 bg-indigo-100 dark:bg-indigo-900/50 rounded-2xl text-indigo-600 dark:text-indigo-400">
                <ListChecks size={32} />
              </div>
              <div>
                <h3 className="text-2xl font-bold text-slate-800 dark:text-white tracking-tight">Rekap Kebutuhan Barang (Master Picking)</h3>
                <p className="text-slate-500 mt-1">Total produk tergabung berdasarkan Nama & Warna (Abaikan ukuran) dari pesanan yang belum selesai.</p>
              </div>
            </div>
          </div>

          <div className="overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-700">
            <table className="w-full text-left text-sm">
              <thead className="bg-slate-50 dark:bg-slate-800/80 text-slate-700 dark:text-slate-300">
                <tr>
                  <th className="p-4 font-bold border-b border-slate-200 dark:border-slate-700">Nama Produk</th>
                  <th className="p-4 font-bold border-b border-slate-200 dark:border-slate-700">Variasi / Warna</th>
                  <th className="p-4 font-bold border-b border-slate-200 dark:border-slate-700 text-center">Total Dibutuhkan</th>
                  <th className="p-4 font-bold border-b border-slate-200 dark:border-slate-700 text-center">Status Produksi</th>
                </tr>
              </thead>
              <tbody>
                {aggregatedProducts.length === 0 ? (
                  <tr>
                    <td colSpan={4} className="p-12 text-center text-slate-500 text-lg">Semua pesanan sudah selesai diproduksi atau antrean kosong!</td>
                  </tr>
                ) : (
                  aggregatedProducts.map((item, idx) => {
                    const isDone = producedItems[item.id];
                    return (
                      <tr key={idx} className={`border-b border-slate-100 dark:border-slate-800/50 last:border-0 hover:bg-indigo-50/50 dark:hover:bg-slate-800 transition-colors ${isDone ? 'bg-emerald-50/30 dark:bg-emerald-900/10' : ''}`}>
                        <td className={`p-4 font-medium transition-all ${isDone ? 'text-slate-400 line-through dark:text-slate-500' : 'text-slate-800 dark:text-slate-200'}`}>
                          {item.product}
                        </td>
                        <td className={`p-4 transition-all ${isDone ? 'text-slate-400 line-through dark:text-slate-500' : 'text-slate-600 dark:text-slate-400'}`}>
                          {item.color}
                        </td>
                        <td className="p-4 text-center">
                          <span className={`inline-flex items-center justify-center font-black px-4 py-1.5 rounded-lg text-lg min-w-[3.5rem] border transition-all ${
                            isDone 
                            ? 'bg-slate-100 text-slate-400 border-slate-200 dark:bg-slate-800 dark:text-slate-500 dark:border-slate-700' 
                            : 'bg-indigo-100 text-indigo-700 border-indigo-200 dark:bg-indigo-900/60 dark:text-indigo-300 dark:border-indigo-800'
                          }`}>
                            {item.qty}
                          </span>
                        </td>
                        <td className="p-4 text-center">
                          {isDone ? (
                            <button 
                              onClick={() => toggleProduce(item.id)}
                              className="flex items-center justify-center gap-2 mx-auto bg-emerald-100 hover:bg-emerald-200 text-emerald-700 px-4 py-2 rounded-xl font-bold transition-all active:scale-95"
                            >
                              <CheckCircle size={18} /> Selesai
                            </button>
                          ) : (
                            <button 
                              onClick={() => toggleProduce(item.id)}
                              className="mx-auto bg-slate-100 hover:bg-indigo-600 text-slate-600 hover:text-white px-4 py-2 rounded-xl font-bold transition-all border border-slate-200 hover:border-indigo-600 active:scale-95 shadow-sm"
                            >
                              Tandai Selesai
                            </button>
                          )}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>

        <div className="mt-12 flex justify-center pb-8">
          <button 
            onClick={() => setResetDialog(true)}
            className="flex items-center gap-2 px-6 py-3 bg-white dark:bg-slate-800 border-2 border-rose-100 dark:border-rose-900/30 text-rose-500 hover:bg-rose-50 hover:border-rose-300 dark:hover:bg-rose-900/20 rounded-full font-bold text-sm transition-all shadow-sm"
          >
            <Trash2 size={16} /> Reset Semua Data Pesanan
          </button>
        </div>

      </div>
    );
  };

  const ImportView = () => (
    <div className="max-w-3xl mx-auto space-y-6">
      <div className="bg-white dark:bg-[#0f172a] rounded-xl p-8 border border-slate-200 dark:border-slate-800 shadow-sm text-center">
        <div className="w-20 h-20 bg-blue-50 dark:bg-blue-900/20 rounded-full flex items-center justify-center mx-auto mb-6">
          <Upload className="text-blue-500" size={36} />
        </div>
        <h2 className="text-2xl font-bold text-slate-800 dark:text-white mb-2">Upload Data Pesanan</h2>
        <p className="text-slate-500 dark:text-slate-400 mb-8 max-w-lg mx-auto">
          Sistem otomatis mengelompokkan pesanan berdasarkan TikTok Shop atau Shopee. Format CSV sangat disarankan untuk TikTok Shop.
        </p>
        
        <div className="relative inline-block">
          <input 
            type="file" 
            accept=".csv, .xls, .xlsx" 
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
            onChange={handleFileUpload}
            disabled={loading}
          />
          <button className={`bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3 px-8 rounded-full shadow-lg shadow-blue-500/30 transition-all ${loading ? 'opacity-50 cursor-wait' : ''}`}>
            {loading ? 'Memproses File...' : 'Pilih File Excel / CSV'}
          </button>
        </div>
      </div>
    </div>
  );

  const PickingListView = () => {
    const [searchQuery, setSearchQuery] = useState('');
    const [confirmDialog, setConfirmDialog] = useState<ConfirmDialogState>({ isOpen: false, orderId: null });
    
    const filteredOrders = orders.filter(order => {
      // Sembunyikan pesanan yang sudah completed dari antrean ini
      if (order.status === 'Completed') return false; 
      const q = searchQuery.toLowerCase();
      const matchOrder = order.id.toLowerCase().includes(q);
      const matchProduct = order.items.some(item => item.product.toLowerCase().includes(q));
      return matchOrder || matchProduct;
    });

    const executeUpdate = (orderId: string, newStatus: OrderStatus) => {
      setOrders(prev => prev.map(o => o.id === orderId ? { ...o, status: newStatus } : o));
      // Update history juga agar sinkron
      setHistoryOrders(prev => prev.map(o => o.id === orderId ? { ...o, status: newStatus } : o));
      
      logActivity(`Status pesanan ${orderId} diubah menjadi ${newStatus}`);
      showToast(`Pesanan dipindahkan ke status ${newStatus}`, 'success');
      setConfirmDialog({ isOpen: false, orderId: null });
    };

    const handleUpdateShippingStatus = (orderId: string, currentStatus: OrderStatus) => {
      if (currentStatus === 'Pending') {
        const order = orders.find(o => o.id === orderId);
        if (!order) return;
        
        const isFullyProduced = order.items.every(item => {
          const prodName = item.product || 'Produk Tidak Diketahui';
          const color = item.color || '-';
          const key = `${prodName.trim().toLowerCase()}|${color.trim().toLowerCase()}`;
          return producedItems[key] === true; 
        });

        if (!isFullyProduced) {
          setConfirmDialog({ isOpen: true, orderId: orderId });
        } else {
          executeUpdate(orderId, 'Ready to Shipment');
        }
      } else if (currentStatus === 'Ready to Shipment') {
        executeUpdate(orderId, 'Pending');
      }
    };

    const getStatusBadge = (status: OrderStatus) => {
      switch(status) {
        case 'Pending': return <span className="px-2 py-1 bg-amber-100 text-amber-700 dark:bg-amber-900/50 dark:text-amber-400 rounded text-xs font-bold uppercase">PENDING</span>;
        case 'Ready to Shipment': return <span className="px-2 py-1 bg-blue-100 text-blue-700 dark:bg-blue-900/50 dark:text-blue-400 rounded text-xs font-bold uppercase">READY TO SHIP</span>;
        default: return <span className="px-2 py-1 bg-slate-100 text-slate-700 rounded text-xs font-bold uppercase">{status}</span>;
      }
    };

    return (
      <div className="space-y-6 print:space-y-0">
        {confirmDialog.isOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in">
            <div className="bg-white dark:bg-slate-800 rounded-2xl p-6 md:p-8 max-w-md w-full shadow-2xl border border-amber-200 dark:border-amber-900/50 zoom-in-95 animate-in">
              <div className="flex items-center gap-4 text-amber-500 mb-4">
                <AlertTriangle size={32} />
                <h3 className="text-xl font-bold text-slate-800 dark:text-white">Peringatan!</h3>
              </div>
              <p className="text-slate-600 dark:text-slate-300 mb-8 font-medium">
                Pesanan Belum Selesai dibuat (Belum dichecklist di Master Picking), Apakah Anda Yakin untuk Melanjutkannya ke Kurir?
              </p>
              <div className="flex justify-end gap-3">
                <button
                  onClick={() => setConfirmDialog({ isOpen: false, orderId: null })}
                  className="px-5 py-2.5 rounded-xl font-bold text-slate-600 bg-slate-100 hover:bg-slate-200 dark:bg-slate-700 dark:text-slate-300 dark:hover:bg-slate-600 transition-colors"
                >
                  No
                </button>
                <button
                  onClick={() => confirmDialog.orderId && executeUpdate(confirmDialog.orderId, 'Ready to Shipment')}
                  className="px-5 py-2.5 rounded-xl font-bold text-white bg-amber-500 hover:bg-amber-600 shadow-lg shadow-amber-500/30 transition-all active:scale-95"
                >
                  Yes, Lanjutkan
                </button>
              </div>
            </div>
          </div>
        )}

        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 print:hidden">
          <div>
            <h2 className="text-2xl font-bold text-slate-800 dark:text-white uppercase tracking-tight">Daftar Antrean Pesanan Aktif</h2>
            <p className="text-slate-500 text-sm mt-1">Menampilkan daftar pesanan utuh yang masuk ke sistem WMS (Tidak menampilkan yang sudah selesai).</p>
          </div>
          <div className="flex gap-3 w-full md:w-auto">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
              <input 
                type="text" placeholder="Cari No Pesanan / Produk..." 
                className="pl-9 pr-4 py-2 rounded-lg border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-white"
                value={searchQuery} onChange={e => setSearchQuery(e.target.value)}
              />
            </div>
            <button onClick={() => isBrowser && window.print()} className="bg-slate-800 hover:bg-slate-900 text-white px-4 py-2 rounded-lg flex items-center gap-2">
              <Printer size={18} /> Print
            </button>
          </div>
        </div>

        <div className="bg-white dark:bg-[#0f172a] rounded-xl shadow-sm border border-slate-200 dark:border-slate-800 overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left text-sm text-slate-600 dark:text-slate-300">
              <thead className="bg-slate-50 dark:bg-[#1e293b] text-slate-700 dark:text-slate-200 font-semibold border-b border-slate-200 dark:border-slate-700">
                <tr>
                  <th className="p-4 whitespace-nowrap">Marketplace</th>
                  <th className="p-4 whitespace-nowrap">No. Pesanan</th>
                  <th className="p-4 min-w-[300px]">Nama Produk</th>
                  <th className="p-4 text-center">Qty Total</th>
                  <th className="p-4 whitespace-nowrap">Kurir</th>
                  <th className="p-4 text-center">Status</th>
                  <th className="p-4 text-center print:hidden">Aksi (Pengiriman)</th>
                </tr>
              </thead>
              <tbody>
                {filteredOrders.length === 0 ? (
                  <tr><td colSpan={7} className="p-8 text-center text-slate-500">Tidak ada pesanan tertunda di Antrean.</td></tr>
                ) : filteredOrders.map((order, idx) => {
                  const totalQty = order.items.reduce((sum, item) => sum + item.qty, 0);
                  const isShopee = order.marketplace.toLowerCase().includes('shopee');
                  
                  return (
                    <tr key={idx} className="border-b border-slate-100 dark:border-slate-800/80 hover:bg-slate-50 dark:hover:bg-slate-800/30">
                      <td className="p-4 font-bold">
                        <span className={isShopee ? 'text-orange-500' : 'text-slate-800 dark:text-white'}>
                          {order.marketplace.toUpperCase()}
                        </span>
                      </td>
                      <td className="p-4 font-mono text-xs">{order.id}</td>
                      <td className="p-4 text-xs leading-relaxed">
                        <div className="space-y-2">
                          {order.items.map((item, i) => (
                            <div key={i} className="bg-slate-50 dark:bg-slate-800/50 p-2 rounded border border-slate-100 dark:border-slate-700">
                              <span className="font-semibold text-slate-800 dark:text-slate-200">[{i+1}] {item.product}</span> <br/>
                              <span className="text-slate-500 dark:text-slate-400">Variasi: {item.color} | Qty: {item.qty}</span>
                            </div>
                          ))}
                        </div>
                      </td>
                      <td className="p-4 text-center text-lg font-bold text-blue-600 dark:text-blue-400">{totalQty}</td>
                      <td className="p-4">{order.shippingMethod}</td>
                      <td className="p-4 text-center">
                        {getStatusBadge(order.status)}
                      </td>
                      <td className="p-4 text-center print:hidden">
                        <button 
                          onClick={() => handleUpdateShippingStatus(order.id, order.status)}
                          className={`px-3 py-2 rounded-md text-xs font-bold w-full transition-colors border shadow-sm active:scale-95 ${
                            order.status === 'Ready to Shipment' 
                            ? 'bg-slate-100 text-slate-600 border-slate-200 hover:bg-slate-200 dark:bg-slate-800 dark:border-slate-700' 
                            : 'bg-blue-600 text-white border-blue-600 hover:bg-blue-700'
                          }`}
                        >
                          {order.status === 'Ready to Shipment' ? 'Batalkan' : 'Siap untuk Kirim'}
                        </button>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  };

  const OrderLogsView = () => {
    const [timeFilter, setTimeFilter] = useState<TimeFilter>('all');
    const [resetStatsDialog, setResetStatsDialog] = useState(false);

    // Menggunakan Data History yang anti-reset dashboard
    const filterOrders = (ordersToFilter: Order[], filterType: TimeFilter) => {
      const now = new Date();
      return ordersToFilter.filter(order => {
        const orderDate = new Date(order.date);
        if (Number.isNaN(orderDate.getTime())) return true;

        if (filterType === 'day') {
          return orderDate.getDate() === now.getDate() && 
                 orderDate.getMonth() === now.getMonth() && 
                 orderDate.getFullYear() === now.getFullYear();
        } else if (filterType === 'week') {
          const diffTime = Math.abs(now.getTime() - orderDate.getTime());
          const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
          return diffDays <= 7;
        } else if (filterType === 'month') {
          return orderDate.getMonth() === now.getMonth() && 
                 orderDate.getFullYear() === now.getFullYear();
        }
        return true; 
      });
    };

    // HISTORY ORDERS digunakan di sini
    const displayOrders = filterOrders(historyOrders, timeFilter);
    const totalDisplayOrders = displayOrders.length;
    
    const productStatsMap = new Map<string, { name: string; color: string; qty: number }>();
    let totalItemsPurchased = 0;

    displayOrders.forEach(order => {
      order.items.forEach(item => {
        totalItemsPurchased += item.qty;
        const prodName = item.product || 'Produk Tidak Diketahui';
        const color = item.color || '-';
        const key = `${prodName.trim()} - ${color.trim()}`;
        
        if (productStatsMap.has(key)) {
          const existingProduct = productStatsMap.get(key);
          if (existingProduct) {
            existingProduct.qty += item.qty;
          }
        } else {
          productStatsMap.set(key, { name: prodName.trim(), color: color.trim(), qty: item.qty });
        }
      });
    });

    const topProducts = Array.from(productStatsMap.values()).sort((a, b) => b.qty - a.qty);

    const confirmResetStats = () => {
      setHistoryOrders([]);
      removeStorage(STORAGE_KEYS.history);
      setResetStatsDialog(false);
      showToast('Data Statistik & Log berhasil dihapus permanen!', 'success');
      logActivity('Mereset seluruh data riwayat penjualan (Statistik)');
    };

    return (
      <div className="space-y-6 animate-in fade-in zoom-in-95">
        
        {resetStatsDialog && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in">
            <div className="bg-white dark:bg-slate-800 rounded-2xl p-6 md:p-8 max-w-md w-full shadow-2xl border border-rose-200 dark:border-rose-900/50 zoom-in-95 animate-in">
              <div className="flex items-center gap-4 text-rose-500 mb-4">
                <AlertTriangle size={32} />
                <h3 className="text-xl font-bold text-slate-800 dark:text-white">Reset Data Statistik?</h3>
              </div>
              <p className="text-slate-600 dark:text-slate-300 mb-8 font-medium">
                Tindakan ini akan menghapus semua riwayat penjualan Anda dari awal. Jika Anda ingin memulai perhitungan di bulan baru, silakan lanjutkan.
              </p>
              <div className="flex justify-end gap-3">
                <button
                  onClick={() => setResetStatsDialog(false)}
                  className="px-5 py-2.5 rounded-xl font-bold text-slate-600 bg-slate-100 hover:bg-slate-200 dark:bg-slate-700 dark:text-slate-300 dark:hover:bg-slate-600 transition-colors"
                >
                  Batal
                </button>
                <button
                  onClick={confirmResetStats}
                  className="px-5 py-2.5 rounded-xl font-bold text-white bg-rose-500 hover:bg-rose-600 shadow-lg shadow-rose-500/30 transition-all active:scale-95"
                >
                  Hapus Permanen
                </button>
              </div>
            </div>
          </div>
        )}

        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
          <div>
            <h2 className="text-2xl font-bold text-slate-800 dark:text-white uppercase tracking-tight flex items-center gap-3">
              <History className="text-blue-500" /> Log & Statistik Penjualan
            </h2>
            <p className="text-slate-500 text-sm mt-1">Pantau akumulasi seluruh riwayat pesanan (Tidak akan terhapus meski antrean gudang dibersihkan).</p>
          </div>
          
          <div className="flex bg-white dark:bg-slate-800 rounded-lg p-1 border border-slate-200 dark:border-slate-700 shadow-sm">
            {([
              { id: 'day', label: 'Hari Ini' },
              { id: 'week', label: 'Minggu Ini' },
              { id: 'month', label: 'Bulan Ini' },
              { id: 'all', label: 'Semua Waktu' }
            ] as const).map(f => (
              <button
                key={f.id}
                onClick={() => setTimeFilter(f.id)}
                className={`px-4 py-2 rounded-md text-sm font-semibold transition-all ${
                  timeFilter === f.id 
                  ? 'bg-blue-100 text-blue-700 dark:bg-blue-900/50 dark:text-blue-400' 
                  : 'text-slate-500 hover:text-slate-800 dark:text-slate-400 dark:hover:text-slate-200'
                }`}
              >
                {f.label}
              </button>
            ))}
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center gap-6">
            <div className="p-4 bg-indigo-50 dark:bg-indigo-900/30 text-indigo-500 rounded-full">
              <BarChart3 size={32} />
            </div>
            <div>
              <p className="text-sm font-bold text-slate-500 uppercase tracking-wide">Total Order Filtered</p>
              <p className="text-4xl font-black text-slate-800 dark:text-white">{totalDisplayOrders} <span className="text-lg font-medium text-slate-400">Pesanan</span></p>
            </div>
          </div>
          
          <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center gap-6">
            <div className="p-4 bg-teal-50 dark:bg-teal-900/30 text-teal-500 rounded-full">
              <TrendingUp size={32} />
            </div>
            <div>
              <p className="text-sm font-bold text-slate-500 uppercase tracking-wide">Total Produk Terjual</p>
              <p className="text-4xl font-black text-slate-800 dark:text-white">{totalItemsPurchased} <span className="text-lg font-medium text-slate-400">Pcs</span></p>
            </div>
          </div>
        </div>

        <div className="bg-white dark:bg-[#0f172a] rounded-2xl shadow-sm border border-slate-200 dark:border-slate-800 overflow-hidden mt-8">
          <div className="p-6 border-b border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-800/50">
            <h3 className="text-lg font-bold text-slate-800 dark:text-white flex items-center gap-2">
              <TrendingUp size={20} className="text-rose-500" /> Urutan Produk Terfavorit
            </h3>
          </div>
          <div className="overflow-x-auto p-4">
            <table className="w-full text-left text-sm">
              <thead>
                <tr className="text-slate-400 border-b border-slate-200 dark:border-slate-700">
                  <th className="p-4 font-bold">Peringkat</th>
                  <th className="p-4 font-bold">Nama Produk</th>
                  <th className="p-4 font-bold">Warna / Variasi</th>
                  <th className="p-4 font-bold text-center">Total Dibeli</th>
                </tr>
              </thead>
              <tbody>
                {topProducts.length === 0 ? (
                  <tr><td colSpan={4} className="p-8 text-center text-slate-500">Belum ada data penjualan di periode ini.</td></tr>
                ) : topProducts.map((item, idx) => (
                  <tr key={idx} className="border-b border-slate-100 dark:border-slate-800/50 hover:bg-slate-50 dark:hover:bg-slate-800/80">
                    <td className="p-4 font-black text-lg text-slate-400">
                      #{idx + 1}
                      {idx === 0 && <span className="ml-2 text-xs bg-amber-100 text-amber-700 px-2 py-1 rounded-full uppercase">Paling Laris</span>}
                    </td>
                    <td className="p-4 font-medium text-slate-800 dark:text-slate-200">{item.name}</td>
                    <td className="p-4 text-slate-500">{item.color}</td>
                    <td className="p-4 text-center">
                      <span className="font-bold text-blue-600 dark:text-blue-400 text-lg">{item.qty}</span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Tombol Reset Mandiri untuk Log & Statistik */}
        <div className="mt-8 flex justify-center pb-8 border-t border-slate-200 dark:border-slate-800 pt-8">
          <button 
            onClick={() => setResetStatsDialog(true)}
            className="flex items-center gap-2 px-6 py-3 bg-white dark:bg-slate-800 border-2 border-rose-100 dark:border-rose-900/30 text-rose-500 hover:bg-rose-50 hover:border-rose-300 dark:hover:bg-rose-900/20 rounded-full font-bold text-sm transition-all shadow-sm"
          >
            <Trash2 size={16} /> Reset Semua Data Statistik
          </button>
        </div>

      </div>
    );
  };

  const NavItem = ({ id, label, icon }: NavItemProps) => (
    <button
      onClick={() => {
        setCurrentTab(id);
        setIsSidebarOpen(false); // Otomatis tutup sidebar di HP saat menu di klik
      }}
      className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-colors text-sm font-medium ${
        currentTab === id 
        ? 'bg-blue-50 text-blue-700 dark:bg-blue-900/40 dark:text-blue-400' 
        : 'text-slate-600 hover:bg-slate-100 dark:text-slate-400 dark:hover:bg-slate-800/50'
      }`}
    >
      {icon} {label}
    </button>
  );

  return (
    <div className="flex flex-col h-screen bg-slate-50 dark:bg-[#0b1120] font-sans selection:bg-blue-200 selection:text-blue-900 relative">
      
      {toast && (
        <div className={`fixed top-4 right-4 z-[60] px-6 py-3 rounded-lg shadow-xl border flex items-center gap-3 animate-in fade-in slide-in-from-top-4 ${
          toast.type === 'success' ? 'bg-emerald-50 border-emerald-200 text-emerald-800 dark:bg-emerald-950 dark:border-emerald-800 dark:text-emerald-300' :
          toast.type === 'error' ? 'bg-rose-50 border-rose-200 text-rose-800 dark:bg-rose-950 dark:border-rose-800 dark:text-rose-300' :
          'bg-white border-slate-200 text-slate-800 dark:bg-slate-800 dark:border-slate-700 dark:text-slate-200'
        }`}>
          {toast.type === 'success' ? <CheckCircle size={20} /> : <AlertTriangle size={20} />}
          <p className="font-medium text-sm">{toast.message}</p>
        </div>
      )}

      {/* Header Mobile (Hanya tampil di HP) */}
      <div className="md:hidden bg-white dark:bg-[#0f172a] border-b border-slate-200 dark:border-slate-800 p-4 flex items-center justify-between z-20 shadow-sm">
        <div className="flex items-center gap-3">
          <img src={logoUrl} alt="Ghaniya Stuff Order" className="h-8 w-8 rounded-lg object-cover shadow-md" />
          <h1 className="font-bold text-lg text-slate-800 dark:text-white tracking-tight">Ghaniya Stuff Order</h1>
        </div>
        <button 
          onClick={() => setIsSidebarOpen(true)} 
          className="text-slate-600 dark:text-slate-300 p-2 bg-slate-100 dark:bg-slate-800 rounded-lg active:scale-95 transition-all"
        >
          <Menu size={24} />
        </button>
      </div>

      <div className="flex flex-1 overflow-hidden">
        {/* Overlay Blur untuk Mobile Sidebar */}
        {isSidebarOpen && (
          <div 
            className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-40 md:hidden animate-in fade-in"
            onClick={() => setIsSidebarOpen(false)}
          />
        )}

        {/* Sidebar Navigation */}
        <aside className={`
          fixed md:static inset-y-0 left-0 z-50 w-64 bg-white dark:bg-[#0f172a] border-r border-slate-200 dark:border-slate-800 flex flex-col print:hidden shadow-2xl md:shadow-none transition-transform duration-300 ease-in-out
          ${isSidebarOpen ? 'translate-x-0' : '-translate-x-full md:translate-x-0'}
        `}>
          <div className="p-6 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <img src={logoUrl} alt="Ghaniya Stuff Order" className="h-8 w-8 rounded-lg object-cover shadow-md" />
              <h1 className="font-bold text-lg text-slate-800 dark:text-white tracking-tight">Ghaniya Stuff Order</h1>
            </div>
            {/* Tombol Tutup Sidebar untuk Mobile */}
            <button onClick={() => setIsSidebarOpen(false)} className="md:hidden text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800 p-1.5 rounded-lg active:scale-95">
              <X size={20} />
            </button>
          </div>
          
          <div className="p-4">
            <p className="text-xs font-bold text-slate-400 dark:text-slate-500 uppercase tracking-wider mb-3 px-4">Menu Utama</p>
          </div>
          
          <nav className="flex-1 px-4 space-y-1 overflow-y-auto">
            <NavItem id="dashboard" label="Dashboard" icon={<LayoutDashboard size={20} />} />
            <NavItem id="import" label="Import Excel" icon={<Upload size={20} />} />
            <NavItem id="picking" label="Antrean Pesanan" icon={<ListChecks size={20} />} />
            <NavItem id="logs" label="Log & Statistik" icon={<BarChart3 size={20} />} />
          </nav>

          <div className="p-4 border-t border-slate-200 dark:border-slate-800 bg-slate-50/50 dark:bg-slate-900/50">
            <div className="rounded-lg p-2">
              <p className="text-xs font-semibold text-slate-500 dark:text-slate-400 mb-3 flex items-center justify-between px-2">
                Activity Log
              </p>
              <div className="space-y-2.5 max-h-32 overflow-y-auto text-[11px] text-slate-600 dark:text-slate-300 px-2">
                {logs.length === 0 ? <span className="text-slate-400 italic">Belum ada aktivitas</span> : 
                  logs.slice(0, 4).map((log, i) => (
                    <div key={i} className="flex flex-col gap-0.5 border-b border-slate-200/50 dark:border-slate-700/50 pb-2 last:border-0">
                      <span className="opacity-50 font-mono">{log.time}</span>
                      <span className="leading-tight">{log.action}</span>
                    </div>
                  ))
                }
              </div>
            </div>
          </div>
        </aside>

        {/* Main Workspace Area */}
        <main className="flex-1 overflow-y-auto relative w-full">
          <div className="p-4 md:p-10 max-w-7xl mx-auto">
            {currentTab === 'dashboard' && <DashboardView />}
            {currentTab === 'import' && <ImportView />}
            {currentTab === 'picking' && <PickingListView />}
            {currentTab === 'logs' && <OrderLogsView />}
          </div>
        </main>
      </div>
    </div>
  );
}
