/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef, useMemo } from 'react';
import {
  Calendar,
  Layers,
  Search,
  Sun,
  Moon,
  LogOut,
  LogIn,
  Plus,
  Trash2,
  FileSpreadsheet,
  Download,
  AlertTriangle,
  Clock,
  CheckCircle2,
  Package,
  Wrench,
  ChevronDown,
  X,
  Upload,
  ExternalLink,
  Minimize2,
  Maximize2
} from 'lucide-react';

const ADMIN_USER = 'admin';
const ADMIN_PASS = 'admin123';

const READ_URL =
  'https://defaulteaa689b48f8740e09c6f7228de4d75.4a.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/33cda42ea60b487ea12bc6d6b2fb094d/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=-W97iRQOOSQ0lRc9j2WefBnqQegAXIPRvQHX1I2bNOU';
const WRITE_URL =
  'https://defaulteaa689b48f8740e09c6f7228de4d75.4a.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/60eeb9c1879249f79dea375e12bfb527/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=aFIZCX9EK8H4QTmltyJtonb_mLDhEi27kQHK_EoglAw';
const DELETE_URL =
  'https://defaulteaa689b48f8740e09c6f7228de4d75.4a.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/b4e3ca791183425e99995266ab810738/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Vy2Jtiuem6SIKGtMW2bSujFxI-9HbF18hq2rQ8v8KoU';

const EXCEL_SHAREPOINT_URL =
  'https://analog-my.sharepoint.com/:x:/r/personal/ariel_mentawan_analog_com/_layouts/15/Doc.aspx?sourcedoc=%7B3febd084-a91a-4d93-814e-e86736350784%7D&action=edit';

interface ColumnDef {
  key: string;
  label: string;
  type: 'text' | 'image' | 'badge';
}

const COLUMNS: ColumnDef[] = [
  { key: 'WW', label: 'WW', type: 'text' },
  { key: 'DATE', label: 'DATE', type: 'text' },
  { key: 'TESTER ID', label: 'TESTER ID', type: 'text' },
  { key: 'HANDLER ID', label: 'HANDLER ID', type: 'text' },
  { key: 'DETAILS', label: 'DETAILS (Image as Evidence)', type: 'image' },
  { key: 'DEVICE', label: 'DEVICE', type: 'text' },
  { key: 'PACKAGE', label: 'PACKAGE', type: 'text' },
  { key: 'Handler Major Part', label: 'Handler Major Part', type: 'text' },
  { key: 'PROBLEM', label: 'PROBLEM', type: 'text' },
  { key: 'ROOT CAUSE', label: 'ROOT CAUSE', type: 'text' },
  { key: 'MATERIAL CODE', label: 'MATERIAL CODE', type: 'text' },
  { key: 'STOCK ITEMS', label: 'Stock Items', type: 'text' },
  { key: 'PENDING ACTION', label: 'PENDING ACTION', type: 'text' },
  { key: 'STATUS', label: 'STATUS', type: 'badge' }
];

interface SparePart {
  id: string;
  name: string;
  code: string;
  handler: string;
  quantity: number;
  image: string;
}

const DEFAULT_SPARES: SparePart[] = [
  {
    id: 'sp-1',
    name: 'Test Socket Pin Block Assembly',
    code: 'SP-MT99-042',
    handler: 'MT99',
    quantity: 18,
    image: 'https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=300&auto=format&fit=crop&q=80'
  },
  {
    id: 'sp-2',
    name: 'High-Precision Pick & Place Rubber Nozzle',
    code: 'SP-MT93-108',
    handler: 'MT93',
    quantity: 5,
    image: 'https://images.unsplash.com/photo-1581092335397-9583fe92d232?w=300&auto=format&fit=crop&q=80'
  },
  {
    id: 'sp-3',
    name: 'High Vacuum Suction Cup Set (Pack of 10)',
    code: 'SP-MT-COM-011',
    handler: 'MT99, MT93',
    quantity: 42,
    image: 'https://images.unsplash.com/photo-1581092580497-e0d23cbdf1dc?w=300&auto=format&fit=crop&q=80'
  },
  {
    id: 'sp-4',
    name: 'Optical Position Sensor Module',
    code: 'SP-MT99-SEN-08',
    handler: 'MT99',
    quantity: 2,
    image: 'https://images.unsplash.com/photo-1581091226825-a6a2a5aee158?w=300&auto=format&fit=crop&q=80'
  },
  {
    id: 'sp-5',
    name: 'Thermal Chuck Heater Cartridge 24V',
    code: 'SP-MT93-TH-99',
    handler: 'MT93',
    quantity: 0,
    image: 'https://images.unsplash.com/photo-1581092162384-8987c1d64718?w=300&auto=format&fit=crop&q=80'
  }
];

function sanitizeValue(v: unknown): string {
  if (v === undefined || v === null) return '';
  const str = String(v).trim();
  if (
    str.toUpperCase().includes('#BLOCKED') ||
    str.startsWith('#REF!') ||
    str.startsWith('#VALUE!') ||
    str.startsWith('#N/A') ||
    str.startsWith('#NAME?')
  ) {
    return '';
  }
  return str;
}

function getRowId(row: Record<string, any>): string {
  if (!row) return '';
  if (row['ID'] !== undefined && row['ID'] !== null && String(row['ID']).trim() !== '') return String(row['ID']).trim();
  if (row['id'] !== undefined && row['id'] !== null && String(row['id']).trim() !== '') return String(row['id']).trim();
  for (const k of Object.keys(row)) {
    if (k.toLowerCase().replace(/[^a-z0-9]/g, '') === 'id' && row[k]) return String(row[k]).trim();
  }
  if (row['__PowerAppsId__']) return String(row['__PowerAppsId__']);
  if (row['itemInternalId']) return String(row['itemInternalId']);
  return '';
}

function getVal(row: Record<string, any>, key: string): string {
  if (!row) return '';
  if (row[key] !== undefined && row[key] !== null) {
    return sanitizeValue(row[key]);
  }

  const aliases: Record<string, string[]> = {
    'WW': ['WW', 'ww'],
    'DATE': ['DATE', 'Date'],
    'TESTER ID': ['TESTER ID', 'TESTER_ID', 'TESTER_x0020_ID', 'Tester ID'],
    'HANDLER ID': ['HANDLER ID', 'HANDLER_ID', 'HANDLER_x0020_ID', 'Handler ID', 'EventName'],
    'DETAILS': [
      'DETAILS (Image as Evidence)',
      'DETAILS\n(Image as Evidence)',
      'Details (Image as Evidence)',
      'DETAILS',
      'Details'
    ],
    'DEVICE': ['DEVICE', 'Device', 'device'],
    'PACKAGE': ['PACKAGE', 'Package', 'package'],
    'Handler Major Part': ['Handler Major Part', 'HANDLER MAJOR PART', 'Handler_x0020_Major_x0020_Part', 'MajorPart'],
    'PROBLEM': ['PROBLEM', 'Problem', 'problem'],
    'ROOT CAUSE': ['ROOT CAUSE', 'ROOT CASUE', 'ROOT_x0020_CAUSE', 'Root Cause'],
    'MATERIAL CODE': ['MATERIAL CODE', 'Material Code', 'MATERIAL_CODE'],
    'STOCK ITEMS': ['Stock Items\nNon Stock Items', 'Stock Items / Non Stock Items', 'Stock Items', 'StockItems'],
    'PENDING ACTION': ['PENDING ACTION', 'Pending Action', 'PENDING_ACTION'],
    'STATUS': ['STATUS', 'Status', 'status']
  };

  const checks = aliases[key] || [key];
  for (const k of checks) {
    if (row[k] !== undefined && row[k] !== null) {
      const v = sanitizeValue(row[k]);
      if (v !== '') return v;
    }
  }

  const cleanTarget = key.toLowerCase().replace(/[^a-z0-9]/g, '');
  for (const actualKey of Object.keys(row)) {
    const cleanActual = actualKey.toLowerCase().replace(/_x[0-9a-f]{4}_/gi, '').replace(/[^a-z0-9]/g, '');
    if (cleanActual === cleanTarget && row[actualKey] !== undefined && row[actualKey] !== null) {
      return sanitizeValue(row[actualKey]);
    }
  }
  return '';
}

function isImageSource(str: string): boolean {
  if (!str || typeof str !== 'string') return false;
  const trimmed = str.trim();
  if (trimmed.toUpperCase().includes('#BLOCKED')) return false;
  return (
    trimmed.startsWith('data:image') ||
    trimmed.startsWith('blob:') ||
    /\.(jpeg|jpg|gif|png|webp|svg)($|\?)/i.test(trimmed)
  );
}

function compressImage(file: File, maxWidth = 600, quality = 0.6): Promise<string> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = (event) => {
      const img = new Image();
      img.src = event.target?.result as string;
      img.onload = () => {
        let width = img.width;
        let height = img.height;
        if (width > maxWidth) {
          height = Math.round((height * maxWidth) / width);
          width = maxWidth;
        }
        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        if (ctx) {
          ctx.drawImage(img, 0, 0, width, height);
          resolve(canvas.toDataURL('image/jpeg', quality));
        } else {
          resolve(event.target?.result as string);
        }
      };
    };
  });
}

export default function App() {
  // Theme & Density
  const [theme, setTheme] = useState<'dark' | 'light'>(() => {
    return (localStorage.getItem('user_theme_pref') as 'dark' | 'light') || 'dark';
  });
  const [density, setDensity] = useState<'normal' | 'compact'>(() => {
    return (localStorage.getItem('user_density_pref') as 'normal' | 'compact') || 'normal';
  });
  const [sidebarCollapsed, setSidebarCollapsed] = useState<boolean>(() => {
    return localStorage.getItem('sidebar_collapsed_pref') === 'true';
  });

  // Auth State
  const [isLoggedIn, setIsLoggedIn] = useState<boolean>(() => {
    return sessionStorage.getItem('app_user_logged_in') === 'true';
  });
  const [showLoginModal, setShowLoginModal] = useState(false);
  const [loginUsername, setLoginUsername] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [loginError, setLoginError] = useState('');

  // Active Category & Filters
  const [activeTab, setActiveTab] = useState<'repair' | 'spareparts'>('repair');
  const [statusFilter, setStatusFilter] = useState<string>('');
  const [spareFilter, setSpareFilter] = useState<string>('');
  const [searchQuery, setSearchQuery] = useState<string>('');

  // Data
  const [repairData, setRepairData] = useState<Record<string, any>[]>([]);
  const [isLoadingRepair, setIsLoadingRepair] = useState<boolean>(true);
  const [selectedRowIndexes, setSelectedRowIndexes] = useState<Set<number>>(new Set());

  // Spare Parts Data
  const [spareData, setSpareData] = useState<SparePart[]>(() => {
    try {
      const saved = localStorage.getItem('spare_parts_inventory_data');
      return saved ? JSON.parse(saved) : DEFAULT_SPARES;
    } catch {
      return DEFAULT_SPARES;
    }
  });

  // Drawers & Modals
  const [showRepairDrawer, setShowRepairDrawer] = useState(false);
  const [editingRepairIndex, setEditingRepairIndex] = useState<number | null>(null);
  const [repairForm, setRepairForm] = useState({
    id: '',
    ww: '',
    date: '',
    tester: '',
    handler: '',
    device: '',
    pkg: '',
    majorPart: '',
    problem: '',
    cause: '',
    material: '',
    stock: 'Stock Items',
    pending: '',
    status: 'Active',
    photo: ''
  });
  const [repairPhotoFileName, setRepairPhotoFileName] = useState('');
  const [isSavingRepair, setIsSavingRepair] = useState(false);

  // Spare Drawer
  const [showSpareDrawer, setShowSpareDrawer] = useState(false);
  const [editingSpareId, setEditingSpareId] = useState<string | null>(null);
  const [spareForm, setSpareForm] = useState({
    name: '',
    code: '',
    handler: 'MT99, MT93',
    quantity: 10,
    image: ''
  });
  const [sparePhotoFileName, setSparePhotoFileName] = useState('');

  // Stock Adjustment Modal
  const [showStockModal, setShowStockModal] = useState(false);
  const [stockModalPart, setStockModalPart] = useState<SparePart | null>(null);
  const [stockModalAction, setStockModalAction] = useState<'add' | 'withdraw'>('add');
  const [stockModalQty, setStockModalQty] = useState(1);
  const [stockModalReason, setStockModalReason] = useState('');

  // Lightbox
  const [lightboxUrl, setLightboxUrl] = useState<string | null>(null);

  // Toast
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const toastTimeoutRef = useRef<NodeJS.Timeout | null>(null);

  // Custom Confirm Modal (safe replacement for window.confirm in iframe)
  const [confirmModal, setConfirmModal] = useState<{
    isOpen: boolean;
    title: string;
    message: string;
    onConfirm: () => void;
  }>({
    isOpen: false,
    title: '',
    message: '',
    onConfirm: () => {}
  });

  // Table Drag-to-Scroll
  const tableWrapRef = useRef<HTMLDivElement>(null);
  const isDraggingRef = useRef(false);
  const startPosRef = useRef({ x: 0, y: 0, scrollLeft: 0, scrollTop: 0 });
  const hasMovedRef = useRef(false);

  // Persist Themes & Preferences
  useEffect(() => {
    localStorage.setItem('user_theme_pref', theme);
  }, [theme]);

  useEffect(() => {
    localStorage.setItem('user_density_pref', density);
  }, [density]);

  useEffect(() => {
    localStorage.setItem('sidebar_collapsed_pref', String(sidebarCollapsed));
  }, [sidebarCollapsed]);

  useEffect(() => {
    try {
      localStorage.setItem('spare_parts_inventory_data', JSON.stringify(spareData));
    } catch (e) {
      console.error(e);
    }
  }, [spareData]);

  const showToast = (msg: string) => {
    if (toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current);
    setToastMessage(msg);
    toastTimeoutRef.current = setTimeout(() => {
      setToastMessage(null);
    }, 3200);
  };

  // Fetch Power Automate Records
  const fetchRepairData = async () => {
    try {
      setIsLoadingRepair(true);
      const res = await fetch(READ_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({})
      });
      const text = await res.text();
      let data: any;
      try {
        data = JSON.parse(text);
      } catch {
        data = [];
      }
      const list = Array.isArray(data) ? data : data.value || (data.body && data.body.value) || [];
      const validRows = list.filter((r: Record<string, any>) => {
        const str = Object.values(r).join('').trim();
        return str.length > 0 && str !== 'null';
      });
      setRepairData(validRows);
      showToast(`Loaded ${validRows.length} active records`);
    } catch (err) {
      console.error(err);
      showToast('Error loading live Excel records');
    } finally {
      setIsLoadingRepair(false);
    }
  };

  useEffect(() => {
    fetchRepairData();
  }, []);

  // Auth Handling
  const handleLoginSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (loginUsername === ADMIN_USER && loginPassword === ADMIN_PASS) {
      setIsLoggedIn(true);
      sessionStorage.setItem('app_user_logged_in', 'true');
      setShowLoginModal(false);
      setLoginError('');
      showToast('Logged in successfully!');
    } else {
      setLoginError('Invalid username or password (default: admin / admin123)');
    }
  };

  const handleLogout = () => {
    setIsLoggedIn(false);
    sessionStorage.removeItem('app_user_logged_in');
    setSelectedRowIndexes(new Set());
    showToast('Logged out');
  };

  // Drag-to-scroll implementation
  const handleMouseDown = (e: React.MouseEvent) => {
    if ((e.target as HTMLElement).closest('button, input, select, a, textarea')) return;
    if (!tableWrapRef.current) return;
    isDraggingRef.current = true;
    hasMovedRef.current = false;
    startPosRef.current = {
      x: e.pageX - tableWrapRef.current.offsetLeft,
      y: e.pageY - tableWrapRef.current.offsetTop,
      scrollLeft: tableWrapRef.current.scrollLeft,
      scrollTop: tableWrapRef.current.scrollTop
    };
  };

  const handleMouseMove = (e: React.MouseEvent) => {
    if (!isDraggingRef.current || !tableWrapRef.current) return;
    e.preventDefault();
    const x = e.pageX - tableWrapRef.current.offsetLeft;
    const y = e.pageY - tableWrapRef.current.offsetTop;
    const walkX = (x - startPosRef.current.x) * 1.5;
    const walkY = (y - startPosRef.current.y) * 1.5;
    if (Math.abs(walkX) > 4 || Math.abs(walkY) > 4) {
      hasMovedRef.current = true;
    }
    tableWrapRef.current.scrollLeft = startPosRef.current.scrollLeft - walkX;
    tableWrapRef.current.scrollTop = startPosRef.current.scrollTop - walkY;
  };

  const handleMouseUp = () => {
    isDraggingRef.current = false;
  };

  // Filtered repair rows
  const filteredRepairData = useMemo(() => {
    const q = searchQuery.toLowerCase().trim();
    return repairData.filter((row) => {
      const matchStatus = !statusFilter || getVal(row, 'STATUS') === statusFilter;
      const matchQuery =
        !q ||
        COLUMNS.some((col) =>
          String(getVal(row, col.key)).toLowerCase().includes(q)
        );
      return matchStatus && matchQuery;
    });
  }, [repairData, statusFilter, searchQuery]);

  // Counts for Repair Sidebar
  const repairCounts = useMemo(() => {
    const total = repairData.length;
    const active = repairData.filter((r) => getVal(r, 'STATUS') === 'Active').length;
    const pending = repairData.filter((r) => getVal(r, 'STATUS') === 'Pending').length;
    const resolved = repairData.filter((r) => getVal(r, 'STATUS') === 'Resolved').length;
    return { total, active, pending, resolved };
  }, [repairData]);

  // Filtered Spares
  const filteredSpareData = useMemo(() => {
    const q = searchQuery.toLowerCase().trim();
    return spareData.filter((part) => {
      if (spareFilter === 'MT99' && !(part.handler || '').toUpperCase().includes('MT99')) return false;
      if (spareFilter === 'MT93' && !(part.handler || '').toUpperCase().includes('MT93')) return false;
      if (spareFilter === 'low' && Number(part.quantity || 0) > 5) return false;
      if (q) {
        const text = `${part.name} ${part.code} ${part.handler} ${part.quantity}`.toLowerCase();
        if (!text.includes(q)) return false;
      }
      return true;
    });
  }, [spareData, spareFilter, searchQuery]);

  // Counts for Spares Sidebar
  const spareCounts = useMemo(() => {
    const total = spareData.length;
    const mt99 = spareData.filter((s) => (s.handler || '').toUpperCase().includes('MT99')).length;
    const mt93 = spareData.filter((s) => (s.handler || '').toUpperCase().includes('MT93')).length;
    const low = spareData.filter((s) => Number(s.quantity || 0) <= 5).length;
    return { total, mt99, mt93, low };
  }, [spareData]);

  // Open Repair Drawer for Add
  const openAddRepair = () => {
    if (!isLoggedIn) {
      showToast('🔒 Log in to add records');
      setShowLoginModal(true);
      return;
    }
    const currentYear = new Date().getFullYear().toString().slice(-2);
    setRepairForm({
      id: '',
      ww: `WW${currentYear}01`,
      date: new Date().toISOString().split('T')[0],
      tester: '',
      handler: '',
      device: '',
      pkg: '',
      majorPart: '',
      problem: '',
      cause: '',
      material: '',
      stock: 'Stock Items',
      pending: '',
      status: 'Active',
      photo: ''
    });
    setRepairPhotoFileName('');
    setEditingRepairIndex(null);
    setShowRepairDrawer(true);
  };

  // Open Repair Drawer for Edit
  const openEditRepair = (idx: number) => {
    if (!isLoggedIn) {
      showToast('🔒 Log in to edit records');
      setShowLoginModal(true);
      return;
    }
    const row = repairData[idx];
    const rId = getRowId(row);
    const existingImg = getVal(row, 'DETAILS');

    setRepairForm({
      id: rId,
      ww: getVal(row, 'WW'),
      date: getVal(row, 'DATE') || new Date().toISOString().split('T')[0],
      tester: getVal(row, 'TESTER ID'),
      handler: getVal(row, 'HANDLER ID'),
      device: getVal(row, 'DEVICE'),
      pkg: getVal(row, 'PACKAGE'),
      majorPart: getVal(row, 'Handler Major Part'),
      problem: getVal(row, 'PROBLEM'),
      cause: getVal(row, 'ROOT CAUSE'),
      material: getVal(row, 'MATERIAL CODE'),
      stock: getVal(row, 'STOCK ITEMS') || 'Stock Items',
      pending: getVal(row, 'PENDING ACTION'),
      status: getVal(row, 'STATUS') || 'Active',
      photo: existingImg
    });
    setRepairPhotoFileName(existingImg ? 'Existing photo kept' : '');
    setEditingRepairIndex(idx);
    setShowRepairDrawer(true);
  };

  // Submit Repair Record
  const handleSaveRepair = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsSavingRepair(true);
    showToast('Syncing with Excel Online via Power Automate...');

    const uniqueId = repairForm.id || `REC-${Date.now()}`;
    const payload = {
      ID: uniqueId,
      WW: repairForm.ww.trim(),
      DATE: repairForm.date,
      TESTER_ID: repairForm.tester.trim(),
      HANDLER_ID: repairForm.handler.trim(),
      DETAILS: repairForm.photo,
      DEVICE: repairForm.device.trim(),
      PACKAGE: repairForm.pkg.trim(),
      MajorPart: repairForm.majorPart.trim(),
      PROBLEM: repairForm.problem.trim(),
      ROOT_CAUSE: repairForm.cause.trim(),
      MATERIAL_CODE: repairForm.material.trim(),
      StockItems: repairForm.stock,
      PENDING_ACTION: repairForm.pending.trim(),
      STATUS: repairForm.status
    };

    try {
      const res = await fetch(WRITE_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      if (res.ok) {
        showToast(repairForm.id ? 'Record updated!' : 'Record added!');
        setShowRepairDrawer(false);
        setTimeout(fetchRepairData, 1500);
      } else {
        const errorText = await res.text();
        showToast(`Power Automate Error: ${errorText.slice(0, 100)}`);
      }
    } catch (err) {
      console.error(err);
      showToast('Network error while communicating with Power Automate.');
    } finally {
      setIsSavingRepair(false);
    }
  };

  // Delete Single Repair Row
  const requestDeleteRepair = (idx: number) => {
    if (!isLoggedIn) {
      showToast('🔒 Log in to delete records');
      setShowLoginModal(true);
      return;
    }
    const row = repairData[idx];
    const rId = getRowId(row);
    const handler = getVal(row, 'HANDLER ID');

    setConfirmModal({
      isOpen: true,
      title: 'Delete Record',
      message: `Are you sure you want to permanently delete record for Handler "${handler || rId}" from Excel Online?`,
      onConfirm: async () => {
        setConfirmModal((prev) => ({ ...prev, isOpen: false }));
        showToast('Deleting row from Excel Online...');
        try {
          const res = await fetch(DELETE_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ ROW_ID: String(rId) })
          });
          if (res.ok) {
            showToast('Row deleted from Excel successfully!');
            setRepairData((prev) => prev.filter((_, i) => i !== idx));
            setTimeout(fetchRepairData, 2000);
          } else {
            const err = await res.text();
            showToast(`Error deleting: ${err.slice(0, 80)}`);
          }
        } catch (err) {
          console.error(err);
          showToast('Network error while deleting.');
        }
      }
    });
  };

  // Bulk Delete
  const requestBulkDelete = () => {
    if (!isLoggedIn || selectedRowIndexes.size === 0) return;
    const count = selectedRowIndexes.size;
    setConfirmModal({
      isOpen: true,
      title: 'Delete Selected Records',
      message: `Are you sure you want to permanently delete ${count} row(s) directly from Excel Online?`,
      onConfirm: async () => {
        setConfirmModal((prev) => ({ ...prev, isOpen: false }));
        showToast(`Deleting ${count} row(s)...`);
        const idsToDelete: string[] = [];
        selectedRowIndexes.forEach((i) => {
          const rId = getRowId(repairData[i]);
          if (rId) idsToDelete.push(rId);
        });

        try {
          for (const rId of idsToDelete) {
            await fetch(DELETE_URL, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ ROW_ID: String(rId) })
            });
          }
          showToast('Deleted successfully! Updating list...');
          setRepairData((prev) => prev.filter((_, idx) => !selectedRowIndexes.has(idx)));
          setSelectedRowIndexes(new Set());
          setTimeout(fetchRepairData, 2000);
        } catch (err) {
          console.error(err);
          showToast('Failed to complete all deletions.');
        }
      }
    });
  };

  // CSV Export for Repair
  const exportRepairCSV = () => {
    let csv = COLUMNS.map((c) => `"${c.label}"`).join(',') + '\n';
    repairData.forEach((row) => {
      csv +=
        COLUMNS.map((c) => `"${getVal(row, c.key).toString().replace(/"/g, '""')}"`).join(',') + '\n';
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `handler_monitoring_export_${new Date().toISOString().split('T')[0]}.csv`;
    link.click();
    URL.revokeObjectURL(url);
    showToast('Exported repair records to CSV');
  };

  // Spare Parts: Add/Edit
  const openAddSpare = () => {
    if (!isLoggedIn) {
      showToast('🔒 Log in to manage spare parts');
      setShowLoginModal(true);
      return;
    }
    setSpareForm({
      name: '',
      code: '',
      handler: 'MT99, MT93',
      quantity: 10,
      image: ''
    });
    setSparePhotoFileName('');
    setEditingSpareId(null);
    setShowSpareDrawer(true);
  };

  const openEditSpare = (part: SparePart) => {
    if (!isLoggedIn) {
      showToast('🔒 Log in to manage spare parts');
      setShowLoginModal(true);
      return;
    }
    setSpareForm({
      name: part.name,
      code: part.code,
      handler: part.handler,
      quantity: part.quantity,
      image: part.image
    });
    setSparePhotoFileName(part.image ? 'Current image loaded' : '');
    setEditingSpareId(part.id);
    setShowSpareDrawer(true);
  };

  const handleSaveSpare = (e: React.FormEvent) => {
    e.preventDefault();
    if (editingSpareId) {
      setSpareData((prev) =>
        prev.map((item) =>
          item.id === editingSpareId
            ? {
                ...item,
                name: spareForm.name.trim(),
                code: spareForm.code.trim(),
                handler: spareForm.handler.trim(),
                quantity: Number(spareForm.quantity) || 0,
                image: spareForm.image || item.image
              }
            : item
        )
      );
      showToast(`Updated "${spareForm.name}"`);
    } else {
      const newPart: SparePart = {
        id: `sp-${Date.now()}`,
        name: spareForm.name.trim(),
        code: spareForm.code.trim(),
        handler: spareForm.handler.trim(),
        quantity: Number(spareForm.quantity) || 0,
        image: spareForm.image
      };
      setSpareData((prev) => [newPart, ...prev]);
      showToast(`Added "${spareForm.name}" to inventory`);
    }
    setShowSpareDrawer(false);
  };

  const requestDeleteSpare = (part: SparePart) => {
    if (!isLoggedIn) return;
    setConfirmModal({
      isOpen: true,
      title: 'Delete Spare Part',
      message: `Are you sure you want to remove "${part.name}" (${part.code}) from inventory?`,
      onConfirm: () => {
        setSpareData((prev) => prev.filter((p) => p.id !== part.id));
        setConfirmModal((prev) => ({ ...prev, isOpen: false }));
        showToast('Spare part deleted');
      }
    });
  };

  // Stock Adjust Modal Open
  const openStockModal = (part: SparePart, action: 'add' | 'withdraw') => {
    setStockModalPart(part);
    setStockModalAction(action);
    setStockModalQty(1);
    setStockModalReason('');
    setShowStockModal(true);
  };

  const handleStockAdjustSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!stockModalPart) return;

    const currentQty = stockModalPart.quantity;
    if (stockModalAction === 'withdraw' && stockModalQty > currentQty) {
      showToast(`Cannot withdraw ${stockModalQty}. Only ${currentQty} in stock.`);
      return;
    }

    const newQty =
      stockModalAction === 'add'
        ? currentQty + stockModalQty
        : currentQty - stockModalQty;

    setSpareData((prev) =>
      prev.map((item) =>
        item.id === stockModalPart.id ? { ...item, quantity: newQty } : item
      )
    );

    showToast(
      `${stockModalAction === 'add' ? 'Added' : 'Withdrew'} ${stockModalQty} unit(s) of ${stockModalPart.code}${
        stockModalReason ? ` (${stockModalReason})` : ''
      }`
    );
    setShowStockModal(false);
  };

  // CSV Export for Spares
  const exportSpareCSV = () => {
    let csv = '"Name","Product Code","Handler Type","Stock Quantity"\n';
    spareData.forEach((p) => {
      csv += `"${(p.name || '').replace(/"/g, '""')}","${(p.code || '').replace(/"/g, '""')}","${(p.handler || '').replace(/"/g, '""')}","${p.quantity || 0}"\n`;
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `spare_parts_inventory_${new Date().toISOString().split('T')[0]}.csv`;
    link.click();
    URL.revokeObjectURL(url);
    showToast('Exported spare parts inventory to CSV');
  };

  return (
    <div className={`flex flex-row h-screen w-screen overflow-hidden ${theme === 'dark' ? 'theme-dark' : 'theme-light'} ${density === 'compact' ? 'density-compact' : 'density-normal'}`}>
      {/* Ambient Glow Shapes */}
      <div className="ambient-shape shape-1" />
      <div className="ambient-shape shape-2" />
      <div className="ambient-shape shape-3" />
      <div className="ambient-shape shape-4" />

      {/* ==================================================== */}
      {/* SIDEBAR                                              */}
      {/* ==================================================== */}
      <aside
        className={`relative z-20 flex flex-col justify-between p-4 flex-shrink-0 transition-all duration-300 backdrop-blur-xl ${
          sidebarCollapsed ? 'w-18 min-w-18 items-center px-2' : 'w-64 min-w-64'
        }`}
        style={{ background: 'var(--glass-bg)' }}
      >
        <div className="flex flex-col gap-4 w-full">
          {/* Brand Row */}
          <div className="flex items-center justify-between px-1">
            <div
              className="flex items-center gap-3 cursor-pointer select-none"
              onClick={() => setSidebarCollapsed(!sidebarCollapsed)}
              title="Toggle Sidebar (Collapse / Expand)"
            >
              <div className="w-10 h-10 rounded-xl flex items-center justify-center font-extrabold text-white shadow-lg bg-gradient-to-br from-orange-500 to-pink-500 flex-shrink-0 transition-transform hover:scale-105">
                <Layers className="w-5 h-5" />
              </div>
              {!sidebarCollapsed && (
                <div className="flex flex-col">
                  <h2 className="text-sm font-extrabold tracking-tight leading-tight" style={{ color: 'var(--text-main)' }}>
                    Handler Monitoring
                  </h2>
                  <span className="text-[11px] font-semibold" style={{ color: 'var(--text-muted)' }}>
                    Intest 2 Dashboard
                  </span>
                </div>
              )}
            </div>
          </div>

          {/* Quick Sidebar Search */}
          {!sidebarCollapsed && (
            <div className="relative w-full">
              <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" style={{ color: 'var(--text-muted)' }} />
              <input
                type="text"
                placeholder={activeTab === 'repair' ? 'Quick search records...' : 'Search spare parts...'}
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-9 pr-3 py-2 rounded-xl text-xs outline-none transition-all"
                style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
              />
            </div>
          )}

          {/* Category Switcher Tabs */}
          {!sidebarCollapsed && (
            <div className="flex items-center p-1 rounded-xl gap-1" style={{ background: 'var(--control-bg)' }}>
              <button
                type="button"
                className={`flex-1 py-1.5 px-2 rounded-lg text-xs font-bold transition-all ${
                  activeTab === 'repair'
                    ? 'shadow-sm text-pink-500'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
                style={{
                  background: activeTab === 'repair' ? 'var(--control-active)' : 'transparent',
                  color: activeTab === 'repair' ? 'var(--text-main)' : 'var(--text-muted)'
                }}
                onClick={() => {
                  setActiveTab('repair');
                  setSelectedRowIndexes(new Set());
                }}
              >
                Handler Repair
              </button>
              <button
                type="button"
                className={`flex-1 py-1.5 px-2 rounded-lg text-xs font-bold transition-all ${
                  activeTab === 'spareparts'
                    ? 'shadow-sm text-pink-500'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
                style={{
                  background: activeTab === 'spareparts' ? 'var(--control-active)' : 'transparent',
                  color: activeTab === 'spareparts' ? 'var(--text-main)' : 'var(--text-muted)'
                }}
                onClick={() => {
                  setActiveTab('spareparts');
                  setSelectedRowIndexes(new Set());
                }}
              >
                Spare Parts
              </button>
            </div>
          )}

          {/* Navigation Links for Repair */}
          {activeTab === 'repair' && (
            <nav className="flex flex-col gap-1 mt-1">
              {!sidebarCollapsed && (
                <div className="text-[10.5px] font-extrabold uppercase tracking-wider px-2 py-1" style={{ color: 'var(--text-muted)' }}>
                  Main Menu
                </div>
              )}

              <button
                type="button"
                className={`flex items-center justify-between p-2.5 rounded-xl text-xs font-semibold transition-all ${
                  statusFilter === '' ? 'text-pink-500 font-bold' : ''
                }`}
                style={{
                  background: statusFilter === '' ? 'var(--sidebar-active-bg)' : 'transparent',
                  color: statusFilter === '' ? 'var(--sidebar-active-text)' : 'var(--text-secondary)'
                }}
                onClick={() => setStatusFilter('')}
                title="All Records"
              >
                <div className="flex items-center gap-2.5">
                  <div
                    className={`w-7 h-7 rounded-lg flex items-center justify-center transition-all ${
                      statusFilter === '' ? 'bg-gradient-to-br from-orange-500 to-pink-500 text-white shadow-md' : ''
                    }`}
                    style={{ background: statusFilter === '' ? undefined : 'var(--control-bg)', color: statusFilter === '' ? '#fff' : 'var(--text-muted)' }}
                  >
                    <Layers className="w-3.5 h-3.5" />
                  </div>
                  {!sidebarCollapsed && <span>All Records</span>}
                </div>
                {!sidebarCollapsed && (
                  <span className="text-xs px-2 py-0.5 rounded-full font-bold bg-pink-500/20 text-pink-400">
                    {repairCounts.total}
                  </span>
                )}
              </button>

              <button
                type="button"
                className={`flex items-center justify-between p-2.5 rounded-xl text-xs font-semibold transition-all ${
                  statusFilter === 'Active' ? 'text-pink-500 font-bold' : ''
                }`}
                style={{
                  background: statusFilter === 'Active' ? 'var(--sidebar-active-bg)' : 'transparent',
                  color: statusFilter === 'Active' ? 'var(--sidebar-active-text)' : 'var(--text-secondary)'
                }}
                onClick={() => setStatusFilter('Active')}
                title="Active Issues"
              >
                <div className="flex items-center gap-2.5">
                  <div
                    className={`w-7 h-7 rounded-lg flex items-center justify-center transition-all ${
                      statusFilter === 'Active' ? 'bg-gradient-to-br from-orange-500 to-pink-500 text-white shadow-md' : ''
                    }`}
                    style={{ background: statusFilter === 'Active' ? undefined : 'var(--control-bg)', color: statusFilter === 'Active' ? '#fff' : 'var(--text-muted)' }}
                  >
                    <AlertTriangle className="w-3.5 h-3.5" />
                  </div>
                  {!sidebarCollapsed && <span>Active Issues</span>}
                </div>
                {!sidebarCollapsed && (
                  <span className="text-xs px-2 py-0.5 rounded-full font-bold bg-red-500/20 text-red-400">
                    {repairCounts.active}
                  </span>
                )}
              </button>

              <button
                type="button"
                className={`flex items-center justify-between p-2.5 rounded-xl text-xs font-semibold transition-all ${
                  statusFilter === 'Pending' ? 'text-pink-500 font-bold' : ''
                }`}
                style={{
                  background: statusFilter === 'Pending' ? 'var(--sidebar-active-bg)' : 'transparent',
                  color: statusFilter === 'Pending' ? 'var(--sidebar-active-text)' : 'var(--text-secondary)'
                }}
                onClick={() => setStatusFilter('Pending')}
                title="Pending Review"
              >
                <div className="flex items-center gap-2.5">
                  <div
                    className={`w-7 h-7 rounded-lg flex items-center justify-center transition-all ${
                      statusFilter === 'Pending' ? 'bg-gradient-to-br from-orange-500 to-pink-500 text-white shadow-md' : ''
                    }`}
                    style={{ background: statusFilter === 'Pending' ? undefined : 'var(--control-bg)', color: statusFilter === 'Pending' ? '#fff' : 'var(--text-muted)' }}
                  >
                    <Clock className="w-3.5 h-3.5" />
                  </div>
                  {!sidebarCollapsed && <span>Pending Review</span>}
                </div>
                {!sidebarCollapsed && (
                  <span className="text-xs px-2 py-0.5 rounded-full font-bold bg-amber-500/20 text-amber-300">
                    {repairCounts.pending}
                  </span>
                )}
              </button>

              <button
                type="button"
                className={`flex items-center justify-between p-2.5 rounded-xl text-xs font-semibold transition-all ${
                  statusFilter === 'Resolved' ? 'text-pink-500 font-bold' : ''
                }`}
                style={{
                  background: statusFilter === 'Resolved' ? 'var(--sidebar-active-bg)' : 'transparent',
                  color: statusFilter === 'Resolved' ? 'var(--sidebar-active-text)' : 'var(--text-secondary)'
                }}
                onClick={() => setStatusFilter('Resolved')}
                title="Resolved"
              >
                <div className="flex items-center gap-2.5">
                  <div
                    className={`w-7 h-7 rounded-lg flex items-center justify-center transition-all ${
                      statusFilter === 'Resolved' ? 'bg-gradient-to-br from-orange-500 to-pink-500 text-white shadow-md' : ''
                    }`}
                    style={{ background: statusFilter === 'Resolved' ? undefined : 'var(--control-bg)', color: statusFilter === 'Resolved' ? '#fff' : 'var(--text-muted)' }}
                  >
                    <CheckCircle2 className="w-3.5 h-3.5" />
                  </div>
                  {!sidebarCollapsed && <span>Resolved</span>}
                </div>
                {!sidebarCollapsed && (
                  <span className="text-xs px-2 py-0.5 rounded-full font-bold bg-emerald-500/20 text-emerald-400">
                    {repairCounts.resolved}
                  </span>
                )}
              </button>

              {isLoggedIn && !sidebarCollapsed && (
                <>
                  <div className="text-[10.5px] font-extrabold uppercase tracking-wider px-2 py-1 mt-3" style={{ color: 'var(--text-muted)' }}>
                    Quick Actions
                  </div>
                  <button
                    type="button"
                    className="flex items-center gap-2.5 p-2.5 rounded-xl text-xs font-semibold transition-all hover:opacity-80 text-left"
                    style={{ color: 'var(--text-secondary)' }}
                    onClick={exportRepairCSV}
                  >
                    <div className="w-7 h-7 rounded-lg flex items-center justify-center" style={{ background: 'var(--control-bg)', color: 'var(--text-muted)' }}>
                      <Download className="w-3.5 h-3.5" />
                    </div>
                    <span>Export CSV</span>
                  </button>
                </>
              )}
            </nav>
          )}

          {/* Navigation Links for Spare Parts */}
          {activeTab === 'spareparts' && (
            <nav className="flex flex-col gap-1 mt-1">
              {!sidebarCollapsed && (
                <div className="text-[10.5px] font-extrabold uppercase tracking-wider px-2 py-1" style={{ color: 'var(--text-muted)' }}>
                  Inventory Filter
                </div>
              )}

              <button
                type="button"
                className={`flex items-center justify-between p-2.5 rounded-xl text-xs font-semibold transition-all ${
                  spareFilter === '' ? 'text-pink-500 font-bold' : ''
                }`}
                style={{
                  background: spareFilter === '' ? 'var(--sidebar-active-bg)' : 'transparent',
                  color: spareFilter === '' ? 'var(--sidebar-active-text)' : 'var(--text-secondary)'
                }}
                onClick={() => setSpareFilter('')}
                title="All Inventory"
              >
                <div className="flex items-center gap-2.5">
                  <div
                    className={`w-7 h-7 rounded-lg flex items-center justify-center transition-all ${
                      spareFilter === '' ? 'bg-gradient-to-br from-orange-500 to-pink-500 text-white shadow-md' : ''
                    }`}
                    style={{ background: spareFilter === '' ? undefined : 'var(--control-bg)', color: spareFilter === '' ? '#fff' : 'var(--text-muted)' }}
                  >
                    <Package className="w-3.5 h-3.5" />
                  </div>
                  {!sidebarCollapsed && <span>All Inventory</span>}
                </div>
                {!sidebarCollapsed && (
                  <span className="text-xs px-2 py-0.5 rounded-full font-bold bg-pink-500/20 text-pink-400">
                    {spareCounts.total}
                  </span>
                )}
              </button>

              <button
                type="button"
                className={`flex items-center justify-between p-2.5 rounded-xl text-xs font-semibold transition-all ${
                  spareFilter === 'MT99' ? 'text-pink-500 font-bold' : ''
                }`}
                style={{
                  background: spareFilter === 'MT99' ? 'var(--sidebar-active-bg)' : 'transparent',
                  color: spareFilter === 'MT99' ? 'var(--sidebar-active-text)' : 'var(--text-secondary)'
                }}
                onClick={() => setSpareFilter('MT99')}
                title="MT99 Spares"
              >
                <div className="flex items-center gap-2.5">
                  <div
                    className={`w-7 h-7 rounded-lg flex items-center justify-center transition-all ${
                      spareFilter === 'MT99' ? 'bg-gradient-to-br from-orange-500 to-pink-500 text-white shadow-md' : ''
                    }`}
                    style={{ background: spareFilter === 'MT99' ? undefined : 'var(--control-bg)', color: spareFilter === 'MT99' ? '#fff' : 'var(--text-muted)' }}
                  >
                    <Wrench className="w-3.5 h-3.5" />
                  </div>
                  {!sidebarCollapsed && <span>MT99 Spares</span>}
                </div>
                {!sidebarCollapsed && (
                  <span className="text-xs px-2 py-0.5 rounded-full font-bold" style={{ background: 'var(--control-bg)', color: 'var(--text-muted)' }}>
                    {spareCounts.mt99}
                  </span>
                )}
              </button>

              <button
                type="button"
                className={`flex items-center justify-between p-2.5 rounded-xl text-xs font-semibold transition-all ${
                  spareFilter === 'MT93' ? 'text-pink-500 font-bold' : ''
                }`}
                style={{
                  background: spareFilter === 'MT93' ? 'var(--sidebar-active-bg)' : 'transparent',
                  color: spareFilter === 'MT93' ? 'var(--sidebar-active-text)' : 'var(--text-secondary)'
                }}
                onClick={() => setSpareFilter('MT93')}
                title="MT93 Spares"
              >
                <div className="flex items-center gap-2.5">
                  <div
                    className={`w-7 h-7 rounded-lg flex items-center justify-center transition-all ${
                      spareFilter === 'MT93' ? 'bg-gradient-to-br from-orange-500 to-pink-500 text-white shadow-md' : ''
                    }`}
                    style={{ background: spareFilter === 'MT93' ? undefined : 'var(--control-bg)', color: spareFilter === 'MT93' ? '#fff' : 'var(--text-muted)' }}
                  >
                    <Wrench className="w-3.5 h-3.5" />
                  </div>
                  {!sidebarCollapsed && <span>MT93 Spares</span>}
                </div>
                {!sidebarCollapsed && (
                  <span className="text-xs px-2 py-0.5 rounded-full font-bold" style={{ background: 'var(--control-bg)', color: 'var(--text-muted)' }}>
                    {spareCounts.mt93}
                  </span>
                )}
              </button>

              <button
                type="button"
                className={`flex items-center justify-between p-2.5 rounded-xl text-xs font-semibold transition-all ${
                  spareFilter === 'low' ? 'text-pink-500 font-bold' : ''
                }`}
                style={{
                  background: spareFilter === 'low' ? 'var(--sidebar-active-bg)' : 'transparent',
                  color: spareFilter === 'low' ? 'var(--sidebar-active-text)' : 'var(--text-secondary)'
                }}
                onClick={() => setSpareFilter('low')}
                title="Low / Out of Stock"
              >
                <div className="flex items-center gap-2.5">
                  <div
                    className={`w-7 h-7 rounded-lg flex items-center justify-center transition-all ${
                      spareFilter === 'low' ? 'bg-gradient-to-br from-orange-500 to-pink-500 text-white shadow-md' : ''
                    }`}
                    style={{ background: spareFilter === 'low' ? undefined : 'var(--control-bg)', color: spareFilter === 'low' ? '#fff' : 'var(--text-muted)' }}
                  >
                    <AlertTriangle className="w-3.5 h-3.5" />
                  </div>
                  {!sidebarCollapsed && <span>Low Stock (≤5)</span>}
                </div>
                {!sidebarCollapsed && (
                  <span className="text-xs px-2 py-0.5 rounded-full font-bold bg-amber-500/20 text-amber-300">
                    {spareCounts.low}
                  </span>
                )}
              </button>

              {isLoggedIn && !sidebarCollapsed && (
                <>
                  <div className="text-[10.5px] font-extrabold uppercase tracking-wider px-2 py-1 mt-3" style={{ color: 'var(--text-muted)' }}>
                    Quick Actions
                  </div>
                  <button
                    type="button"
                    className="flex items-center gap-2.5 p-2.5 rounded-xl text-xs font-semibold transition-all hover:opacity-80 text-left"
                    style={{ color: 'var(--text-secondary)' }}
                    onClick={exportSpareCSV}
                  >
                    <div className="w-7 h-7 rounded-lg flex items-center justify-center" style={{ background: 'var(--control-bg)', color: 'var(--text-muted)' }}>
                      <Download className="w-3.5 h-3.5" />
                    </div>
                    <span>Export Inventory</span>
                  </button>
                </>
              )}
            </nav>
          )}
        </div>

        {/* Sidebar Bottom: Theme & Auth */}
        <div className="w-full pt-3">
          <div
            className="flex items-center justify-between p-2 rounded-2xl gap-2 backdrop-blur-md"
            style={{ background: 'var(--control-bg)' }}
          >
            {/* Theme Toggle */}
            <div
              className="flex items-center p-1 rounded-full cursor-pointer gap-1 transition-all"
              style={{ background: 'rgba(255,255,255,0.08)' }}
              onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}
              title={`Switch to ${theme === 'dark' ? 'Light' : 'Dark'} Mode`}
            >
              <div
                className={`w-6 h-6 rounded-full flex items-center justify-center transition-all ${
                  theme === 'light' ? 'bg-white text-amber-500 shadow-sm' : 'text-slate-400'
                }`}
              >
                <Sun className="w-3.5 h-3.5" />
              </div>
              <div
                className={`w-6 h-6 rounded-full flex items-center justify-center transition-all ${
                  theme === 'dark' ? 'bg-white text-slate-800 shadow-sm' : 'text-slate-400'
                }`}
              >
                <Moon className="w-3.5 h-3.5" />
              </div>
            </div>

            {!sidebarCollapsed && (
              <span className="text-xs font-bold whitespace-nowrap select-none" style={{ color: 'var(--text-main)' }}>
                {theme === 'dark' ? 'Dark Mode' : 'Light Mode'}
              </span>
            )}

            {/* Auth Action Button */}
            {isLoggedIn ? (
              <button
                type="button"
                className="w-8 h-8 rounded-xl flex items-center justify-center transition-all bg-red-500/20 text-red-400 hover:bg-red-500 hover:text-white"
                onClick={handleLogout}
                title="Logout"
              >
                <LogOut className="w-4 h-4" />
              </button>
            ) : (
              <button
                type="button"
                className="w-8 h-8 rounded-xl flex items-center justify-center transition-all hover:scale-105"
                style={{ background: 'var(--control-active)', color: 'var(--text-main)' }}
                onClick={() => setShowLoginModal(true)}
                title="Login to Edit"
              >
                <LogIn className="w-4 h-4" />
              </button>
            )}
          </div>
        </div>
      </aside>

      {/* ==================================================== */}
      {/* MAIN CONTENT                                         */}
      {/* ==================================================== */}
      <main className="flex-1 flex flex-col h-screen overflow-hidden min-w-0 relative">
        {/* Header Bar */}
        <header
          className="relative z-20 flex items-center justify-between px-6 py-3.5 gap-4 backdrop-blur-xl"
          style={{ background: 'var(--glass-bg)' }}
        >
          <div className="flex items-center gap-3">
            <button
              type="button"
              className="w-9 h-9 rounded-xl flex items-center justify-center transition-all hover:scale-105"
              style={{ background: 'var(--control-bg)', color: 'var(--text-muted)' }}
              onClick={() => setSidebarCollapsed(!sidebarCollapsed)}
              title="Toggle Sidebar"
            >
              <Minimize2 className="w-4 h-4" />
            </button>

            <div className="relative flex items-center w-80 max-w-sm">
              <Search className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none" style={{ color: 'var(--text-muted)' }} />
              <input
                type="text"
                placeholder={activeTab === 'repair' ? 'Search across repair fields...' : 'Search spare parts & stock...'}
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-10 pr-4 py-2 rounded-xl text-xs outline-none transition-all"
                style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
              />
            </div>
          </div>

          <div className="flex items-center gap-2.5">
            {/* Density switch */}
            <div className="flex items-center p-1 rounded-xl gap-1" style={{ background: 'var(--control-bg)' }}>
              <button
                type="button"
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${
                  density === 'compact'
                    ? 'shadow-sm text-pink-500 font-bold'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
                style={{
                  background: density === 'compact' ? 'var(--control-active)' : 'transparent',
                  color: density === 'compact' ? 'var(--text-main)' : 'var(--text-muted)'
                }}
                onClick={() => setDensity('compact')}
              >
                Compact
              </button>
              <button
                type="button"
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${
                  density === 'normal'
                    ? 'shadow-sm text-pink-500 font-bold'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
                style={{
                  background: density === 'normal' ? 'var(--control-active)' : 'transparent',
                  color: density === 'normal' ? 'var(--text-main)' : 'var(--text-muted)'
                }}
                onClick={() => setDensity('normal')}
              >
                Normal
              </button>
            </div>

            {/* Bulk delete button */}
            {isLoggedIn && activeTab === 'repair' && selectedRowIndexes.size > 0 && (
              <button
                type="button"
                className="flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-bold transition-all bg-red-500/20 text-red-400 hover:bg-red-500 hover:text-white shadow-sm"
                onClick={requestBulkDelete}
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Delete Selected ({selectedRowIndexes.size})</span>
              </button>
            )}

            {/* Open Excel Link */}
            <a
              href={EXCEL_SHAREPOINT_URL}
              target="_blank"
              rel="noopener noreferrer"
              className="flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-semibold transition-all hover:scale-105 shadow-sm"
              style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
              title="Open Original Excel Sheet in SharePoint"
            >
              <FileSpreadsheet className="w-4 h-4 text-emerald-500" />
              <span>Open Excel</span>
              <ExternalLink className="w-3 h-3 opacity-60 ml-0.5" />
            </a>

            {/* Add Row / Add Spare button */}
            {isLoggedIn && activeTab === 'repair' && (
              <button
                type="button"
                className="flex items-center gap-1.5 px-4 py-2 rounded-xl text-xs font-bold text-white shadow-lg bg-gradient-to-br from-orange-500 to-pink-500 hover:opacity-95 transition-all"
                onClick={openAddRepair}
              >
                <Plus className="w-4 h-4" />
                <span>Add Row</span>
              </button>
            )}

            {isLoggedIn && activeTab === 'spareparts' && (
              <button
                type="button"
                className="flex items-center gap-1.5 px-4 py-2 rounded-xl text-xs font-bold text-white shadow-lg bg-gradient-to-br from-orange-500 to-pink-500 hover:opacity-95 transition-all"
                onClick={openAddSpare}
              >
                <Plus className="w-4 h-4" />
                <span>Add Spare Part</span>
              </button>
            )}
          </div>
        </header>

        {/* ==================================================== */}
        {/* TABLE CANVAS (HANDLER REPAIR)                       */}
        {/* ==================================================== */}
        {activeTab === 'repair' && (
          <div
            ref={tableWrapRef}
            onMouseDown={handleMouseDown}
            onMouseMove={handleMouseMove}
            onMouseUp={handleMouseUp}
            className="flex-1 m-4 mt-3 rounded-2xl overflow-auto select-none cursor-grab active:cursor-grabbing backdrop-blur-2xl"
            style={{
              background: 'var(--glass-card)',
              boxShadow: 'var(--glass-shadow)'
            }}
          >
            <table className="w-max min-w-full border-collapse border-spacing-0 text-xs" style={{ fontSize: 'var(--font-size)' }}>
              <thead>
                <tr>
                  {isLoggedIn && (
                    <th
                      className="sticky top-0 z-10 w-12 text-center p-3"
                      style={{ background: 'var(--glass-th)', color: 'var(--text-muted)' }}
                    >
                      <input
                        type="checkbox"
                        checked={selectedRowIndexes.size > 0 && selectedRowIndexes.size === filteredRepairData.length}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setSelectedRowIndexes(new Set(filteredRepairData.map((_, i) => i)));
                          } else {
                            setSelectedRowIndexes(new Set());
                          }
                        }}
                        className="w-4 h-4 rounded cursor-pointer"
                      />
                    </th>
                  )}
                  <th
                    className="sticky top-0 z-10 w-12 text-center p-3 font-bold text-[11px] uppercase tracking-wider"
                    style={{ background: 'var(--glass-th)', color: 'var(--text-muted)' }}
                  >
                    #
                  </th>
                  {COLUMNS.map((col) => (
                    <th
                      key={col.key}
                      className="sticky top-0 z-10 text-left p-3 font-bold text-[11px] uppercase tracking-wider whitespace-nowrap"
                      style={{ background: 'var(--glass-th)', color: 'var(--text-muted)' }}
                    >
                      {col.label}
                    </th>
                  ))}
                  {isLoggedIn && (
                    <th
                      className="sticky top-0 z-10 text-center p-3 font-bold text-[11px] uppercase tracking-wider w-36"
                      style={{ background: 'var(--glass-th)', color: 'var(--text-muted)' }}
                    >
                      ACTIONS
                    </th>
                  )}
                </tr>
              </thead>
              <tbody>
                {isLoadingRepair ? (
                  <tr>
                    <td
                      colSpan={COLUMNS.length + 3}
                      className="text-center py-16 text-sm font-semibold"
                      style={{ color: 'var(--text-muted)' }}
                    >
                      Loading live data from Excel Online...
                    </td>
                  </tr>
                ) : filteredRepairData.length === 0 ? (
                  <tr>
                    <td
                      colSpan={COLUMNS.length + 3}
                      className="text-center py-16 text-sm font-semibold"
                      style={{ color: 'var(--text-muted)' }}
                    >
                      No matching records found.
                    </td>
                  </tr>
                ) : (
                  filteredRepairData.map((row, idx) => {
                    const isSelected = selectedRowIndexes.has(idx);
                    return (
                      <tr
                        key={idx}
                        className={`transition-colors border-b border-white/5 ${
                          isSelected ? 'bg-orange-500/15' : 'hover:bg-white/5'
                        }`}
                        style={{
                          height: 'var(--row-height)',
                          backgroundColor: idx % 2 === 1 ? 'var(--row-stripe-bg)' : undefined
                        }}
                      >
                        {isLoggedIn && (
                          <td className="text-center px-3 py-2">
                            <input
                              type="checkbox"
                              checked={isSelected}
                              onChange={(e) => {
                                e.stopPropagation();
                                const next = new Set(selectedRowIndexes);
                                if (next.has(idx)) next.delete(idx);
                                else next.add(idx);
                                setSelectedRowIndexes(next);
                              }}
                              className="w-4 h-4 rounded cursor-pointer"
                            />
                          </td>
                        )}
                        <td className="text-center px-3 py-2 font-semibold text-xs" style={{ color: 'var(--text-muted)' }}>
                          {idx + 1}
                        </td>
                        {COLUMNS.map((col) => {
                          const val = getVal(row, col.key);
                          const isImg = isImageSource(val) || (col.type === 'image' && val.startsWith('http'));

                          if (col.type === 'image') {
                            return (
                              <td key={col.key} className="px-3 py-2">
                                {isImg ? (
                                  <img
                                    src={val}
                                    alt="Evidence"
                                    onClick={() => !hasMovedRef.current && setLightboxUrl(val)}
                                    className="w-8 h-8 rounded-lg object-cover cursor-pointer hover:scale-125 transition-transform shadow-md"
                                  />
                                ) : (
                                  <span style={{ color: 'var(--text-muted)' }}>&mdash;</span>
                                )}
                              </td>
                            );
                          }

                          if (col.type === 'badge') {
                            const st = val || 'Active';
                            let badgeStyle = 'bg-red-500/20 text-red-400';
                            if (st === 'Pending') badgeStyle = 'bg-amber-500/20 text-amber-300';
                            if (st === 'Resolved') badgeStyle = 'bg-emerald-500/20 text-emerald-400';

                            return (
                              <td key={col.key} className="px-3 py-2">
                                <span className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] font-bold ${badgeStyle}`}>
                                  <span className="w-1.5 h-1.5 rounded-full bg-current" />
                                  {st}
                                </span>
                              </td>
                            );
                          }

                          return (
                            <td key={col.key} className="px-3 py-2 whitespace-nowrap" style={{ color: 'var(--text-main)' }}>
                              {val !== '' ? val : <span style={{ color: 'var(--text-muted)' }}>&mdash;</span>}
                            </td>
                          );
                        })}

                        {isLoggedIn && (
                          <td className="px-3 py-2 text-center whitespace-nowrap">
                            <button
                              type="button"
                              className="px-2.5 py-1 rounded-lg text-xs font-semibold mr-1.5 transition-all text-sky-400 hover:bg-sky-400/20"
                              onClick={() => openEditRepair(idx)}
                              title="Edit Record"
                            >
                              Edit
                            </button>
                            <button
                              type="button"
                              className="px-2.5 py-1 rounded-lg text-xs font-semibold transition-all text-red-400 hover:bg-red-400/20"
                              onClick={() => requestDeleteRepair(idx)}
                              title="Delete Record"
                            >
                              Delete
                            </button>
                          </td>
                        )}
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        )}

        {/* ==================================================== */}
        {/* TABLE CANVAS (SPARE PARTS INVENTORY)                 */}
        {/* ==================================================== */}
        {activeTab === 'spareparts' && (
          <div
            ref={tableWrapRef}
            onMouseDown={handleMouseDown}
            onMouseMove={handleMouseMove}
            onMouseUp={handleMouseUp}
            className="flex-1 m-4 mt-3 rounded-2xl overflow-auto select-none cursor-grab active:cursor-grabbing backdrop-blur-2xl"
            style={{
              background: 'var(--glass-card)',
              boxShadow: 'var(--glass-shadow)'
            }}
          >
            <table className="w-max min-w-full border-collapse border-spacing-0 text-xs" style={{ fontSize: 'var(--font-size)' }}>
              <thead>
                <tr>
                  <th
                    className="sticky top-0 z-10 w-12 text-center p-3 font-bold text-[11px] uppercase tracking-wider"
                    style={{ background: 'var(--glass-th)', color: 'var(--text-muted)' }}
                  >
                    #
                  </th>
                  <th
                    className="sticky top-0 z-10 text-left p-3 font-bold text-[11px] uppercase tracking-wider whitespace-nowrap"
                    style={{ background: 'var(--glass-th)', color: 'var(--text-muted)' }}
                  >
                    NAME
                  </th>
                  <th
                    className="sticky top-0 z-10 text-left p-3 font-bold text-[11px] uppercase tracking-wider whitespace-nowrap"
                    style={{ background: 'var(--glass-th)', color: 'var(--text-muted)' }}
                  >
                    PRODUCT CODE
                  </th>
                  <th
                    className="sticky top-0 z-10 text-left p-3 font-bold text-[11px] uppercase tracking-wider whitespace-nowrap"
                    style={{ background: 'var(--glass-th)', color: 'var(--text-muted)' }}
                  >
                    HANDLER TYPE
                  </th>
                  <th
                    className="sticky top-0 z-10 text-left p-3 font-bold text-[11px] uppercase tracking-wider whitespace-nowrap"
                    style={{ background: 'var(--glass-th)', color: 'var(--text-muted)' }}
                  >
                    STOCKS QUANTITY
                  </th>
                  <th
                    className="sticky top-0 z-10 text-left p-3 font-bold text-[11px] uppercase tracking-wider whitespace-nowrap"
                    style={{ background: 'var(--glass-th)', color: 'var(--text-muted)' }}
                  >
                    IMAGE
                  </th>
                  <th
                    className="sticky top-0 z-10 text-center p-3 font-bold text-[11px] uppercase tracking-wider whitespace-nowrap w-40"
                    style={{ background: 'var(--glass-th)', color: 'var(--text-muted)' }}
                  >
                    STOCK ADJUST
                  </th>
                  {isLoggedIn && (
                    <th
                      className="sticky top-0 z-10 text-center p-3 font-bold text-[11px] uppercase tracking-wider whitespace-nowrap w-36"
                      style={{ background: 'var(--glass-th)', color: 'var(--text-muted)' }}
                    >
                      ACTIONS
                    </th>
                  )}
                </tr>
              </thead>
              <tbody>
                {filteredSpareData.length === 0 ? (
                  <tr>
                    <td
                      colSpan={isLoggedIn ? 8 : 7}
                      className="text-center py-16 text-sm font-semibold"
                      style={{ color: 'var(--text-muted)' }}
                    >
                      No spare parts match the filter.
                    </td>
                  </tr>
                ) : (
                  filteredSpareData.map((part, idx) => {
                    const qty = Number(part.quantity || 0);
                    let badgeClass = 'bg-emerald-500/20 text-emerald-400';
                    let badgeLabel = `${qty} in stock`;
                    if (qty === 0) {
                      badgeClass = 'bg-red-500/20 text-red-400';
                      badgeLabel = 'Out of Stock';
                    } else if (qty <= 5) {
                      badgeClass = 'bg-amber-500/20 text-amber-300';
                      badgeLabel = `${qty} (Low Stock)`;
                    }

                    return (
                      <tr
                        key={part.id}
                        className="transition-colors border-b border-white/5 hover:bg-white/5"
                        style={{
                          height: 'var(--row-height)',
                          backgroundColor: idx % 2 === 1 ? 'var(--row-stripe-bg)' : undefined
                        }}
                      >
                        <td className="text-center px-3 py-2 font-semibold text-xs" style={{ color: 'var(--text-muted)' }}>
                          {idx + 1}
                        </td>
                        <td className="px-3 py-2 font-bold whitespace-nowrap" style={{ color: 'var(--text-main)' }}>
                          {part.name}
                        </td>
                        <td className="px-3 py-2 whitespace-nowrap">
                          <code className="text-xs px-2 py-0.5 rounded font-mono" style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}>
                            {part.code}
                          </code>
                        </td>
                        <td className="px-3 py-2 whitespace-nowrap">
                          {part.handler.split(',').map((h, i) => (
                            <span
                              key={i}
                              className="text-[11px] font-bold px-2 py-0.5 rounded-md mr-1"
                              style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                            >
                              {h.trim()}
                            </span>
                          ))}
                        </td>
                        <td className="px-3 py-2 whitespace-nowrap">
                          <span className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-extrabold ${badgeClass}`}>
                            {badgeLabel}
                          </span>
                        </td>
                        <td className="px-3 py-2 whitespace-nowrap">
                          {part.image ? (
                            <img
                              src={part.image}
                              alt={part.name}
                              onClick={() => !hasMovedRef.current && setLightboxUrl(part.image)}
                              className="w-8 h-8 rounded-lg object-cover cursor-pointer hover:scale-125 transition-transform shadow-md"
                            />
                          ) : (
                            <span style={{ color: 'var(--text-muted)' }}>&mdash;</span>
                          )}
                        </td>
                        <td className="px-3 py-2 text-center whitespace-nowrap">
                          <div className="inline-flex items-center gap-1 p-1 rounded-xl" style={{ background: 'var(--control-bg)' }}>
                            <button
                              type="button"
                              className="w-6 h-6 rounded-lg flex items-center justify-center font-extrabold transition-all hover:bg-red-500 hover:text-white"
                              onClick={() => openStockModal(part, 'withdraw')}
                              title="Withdraw Stocks (-)"
                            >
                              &minus;
                            </button>
                            <span className="min-w-7 text-center font-extrabold text-xs" style={{ color: 'var(--text-main)' }}>
                              {qty}
                            </span>
                            <button
                              type="button"
                              className="w-6 h-6 rounded-lg flex items-center justify-center font-extrabold transition-all hover:bg-emerald-500 hover:text-white"
                              onClick={() => openStockModal(part, 'add')}
                              title="Add / Replenish Stocks (+)"
                            >
                              +
                            </button>
                          </div>
                        </td>

                        {isLoggedIn && (
                          <td className="px-3 py-2 text-center whitespace-nowrap">
                            <button
                              type="button"
                              className="px-2.5 py-1 rounded-lg text-xs font-semibold mr-1.5 transition-all text-sky-400 hover:bg-sky-400/20"
                              onClick={() => openEditSpare(part)}
                              title="Edit Part"
                            >
                              Edit
                            </button>
                            <button
                              type="button"
                              className="px-2.5 py-1 rounded-lg text-xs font-semibold transition-all text-red-400 hover:bg-red-400/20"
                              onClick={() => requestDeleteSpare(part)}
                              title="Delete Part"
                            >
                              Delete
                            </button>
                          </td>
                        )}
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        )}

        {/* Footer */}
        <footer
          className="relative z-10 flex items-center justify-between px-6 py-2.5 text-xs backdrop-blur-xl"
          style={{ background: 'var(--glass-bg)', color: 'var(--text-muted)' }}
        >
          <span>
            {activeTab === 'repair'
              ? `Showing ${filteredRepairData.length} of ${repairData.length} repair records`
              : `Showing ${filteredSpareData.length} of ${spareData.length} spare parts`}
          </span>
          <span className="text-[11px] font-semibold opacity-75">
            Intest 2 • SharePoint Excel Online Synced
          </span>
        </footer>
      </main>

      {/* ==================================================== */}
      {/* REPAIR ADD / EDIT DRAWER                            */}
      {/* ==================================================== */}
      {showRepairDrawer && (
        <>
          <div
            className="fixed inset-0 bg-black/60 backdrop-blur-sm z-40"
            onClick={() => setShowRepairDrawer(false)}
          />
          <div
            className="fixed top-0 right-0 w-[560px] max-w-[95vw] h-full z-50 flex flex-col backdrop-blur-2xl shadow-2xl transition-all"
            style={{ background: 'var(--modal-bg)' }}
          >
            <div className="flex items-center justify-between p-6 border-b border-white/5">
              <h2 className="text-base font-extrabold flex items-center gap-2" style={{ color: 'var(--text-main)' }}>
                {editingRepairIndex !== null ? (
                  <>
                    <Wrench className="w-5 h-5 text-sky-400" />
                    <span>Edit Record</span>
                  </>
                ) : (
                  <>
                    <Plus className="w-5 h-5 text-orange-500" />
                    <span>Add New Row / Record</span>
                  </>
                )}
              </h2>
              <button
                type="button"
                className="w-8 h-8 rounded-lg flex items-center justify-center transition-all hover:bg-white/10"
                style={{ color: 'var(--text-muted)' }}
                onClick={() => setShowRepairDrawer(false)}
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <form onSubmit={handleSaveRepair} className="flex-1 overflow-y-auto p-6 flex flex-col gap-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>WW *</label>
                  <input
                    type="text"
                    required
                    value={repairForm.ww}
                    onChange={(e) => setRepairForm({ ...repairForm, ww: e.target.value })}
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  />
                </div>

                <div>
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Date *</label>
                  <input
                    type="date"
                    required
                    value={repairForm.date}
                    onChange={(e) => setRepairForm({ ...repairForm, date: e.target.value })}
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  />
                </div>

                <div>
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Tester ID</label>
                  <input
                    type="text"
                    value={repairForm.tester}
                    onChange={(e) => setRepairForm({ ...repairForm, tester: e.target.value })}
                    placeholder="e.g. TST-01"
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  />
                </div>

                <div>
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Handler ID *</label>
                  <input
                    type="text"
                    required
                    value={repairForm.handler}
                    onChange={(e) => setRepairForm({ ...repairForm, handler: e.target.value })}
                    placeholder="e.g. MT99-04"
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  />
                </div>

                <div className="col-span-2">
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>
                    DETAILS (Image as Evidence)
                  </label>
                  <div className="flex items-center gap-3">
                    <div
                      className="w-14 h-14 rounded-xl flex items-center justify-center overflow-hidden flex-shrink-0"
                      style={{ background: 'var(--control-bg)', color: 'var(--text-muted)' }}
                    >
                      {repairForm.photo ? (
                        <img src={repairForm.photo} alt="Preview" className="w-full h-full object-cover" />
                      ) : (
                        <Upload className="w-5 h-5" />
                      )}
                    </div>
                    <label className="cursor-pointer px-4 py-2 rounded-xl text-xs font-semibold flex items-center gap-2 transition-all hover:opacity-90 shadow-sm" style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}>
                      <Upload className="w-3.5 h-3.5" />
                      <span>Choose Photo</span>
                      <input
                        type="file"
                        accept="image/*"
                        className="hidden"
                        onChange={async (e) => {
                          const file = e.target.files?.[0];
                          if (file) {
                            setRepairPhotoFileName(`${file.name} (Compressing...)`);
                            const compressed = await compressImage(file, 600, 0.6);
                            setRepairForm((prev) => ({ ...prev, photo: compressed }));
                            setRepairPhotoFileName(`${file.name} (Ready)`);
                          }
                        }}
                      />
                    </label>
                    <span className="text-xs truncate max-w-xs" style={{ color: 'var(--text-muted)' }}>
                      {repairPhotoFileName || 'No photo chosen'}
                    </span>
                  </div>
                </div>

                <div>
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Device</label>
                  <input
                    type="text"
                    value={repairForm.device}
                    onChange={(e) => setRepairForm({ ...repairForm, device: e.target.value })}
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  />
                </div>

                <div>
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Package</label>
                  <input
                    type="text"
                    value={repairForm.pkg}
                    onChange={(e) => setRepairForm({ ...repairForm, pkg: e.target.value })}
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  />
                </div>

                <div className="col-span-2">
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Handler Major Part</label>
                  <input
                    type="text"
                    value={repairForm.majorPart}
                    onChange={(e) => setRepairForm({ ...repairForm, majorPart: e.target.value })}
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  />
                </div>

                <div className="col-span-2">
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Problem *</label>
                  <input
                    type="text"
                    required
                    value={repairForm.problem}
                    onChange={(e) => setRepairForm({ ...repairForm, problem: e.target.value })}
                    placeholder="Describe issue..."
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  />
                </div>

                <div className="col-span-2">
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Root Cause</label>
                  <input
                    type="text"
                    value={repairForm.cause}
                    onChange={(e) => setRepairForm({ ...repairForm, cause: e.target.value })}
                    placeholder="Root cause identified..."
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  />
                </div>

                <div>
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Material Code</label>
                  <input
                    type="text"
                    value={repairForm.material}
                    onChange={(e) => setRepairForm({ ...repairForm, material: e.target.value })}
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  />
                </div>

                <div>
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Stock Items</label>
                  <select
                    value={repairForm.stock}
                    onChange={(e) => setRepairForm({ ...repairForm, stock: e.target.value })}
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none cursor-pointer"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  >
                    <option value="Stock Items">Stock Items</option>
                    <option value="Non Stock Items">Non Stock Items</option>
                  </select>
                </div>

                <div>
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Pending Action</label>
                  <input
                    type="text"
                    value={repairForm.pending}
                    onChange={(e) => setRepairForm({ ...repairForm, pending: e.target.value })}
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  />
                </div>

                <div>
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Status</label>
                  <select
                    value={repairForm.status}
                    onChange={(e) => setRepairForm({ ...repairForm, status: e.target.value })}
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none cursor-pointer font-bold"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  >
                    <option value="Active">🔴 Active</option>
                    <option value="Pending">🟡 Pending</option>
                    <option value="Resolved">🟢 Resolved</option>
                  </select>
                </div>
              </div>

              <div className="flex items-center justify-end gap-2 mt-auto pt-6 border-t border-white/5">
                <button
                  type="button"
                  className="px-4 py-2 rounded-xl text-xs font-semibold transition-all hover:bg-white/10"
                  style={{ color: 'var(--text-main)' }}
                  onClick={() => setShowRepairDrawer(false)}
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={isSavingRepair}
                  className="px-5 py-2 rounded-xl text-xs font-bold text-white shadow-lg bg-gradient-to-br from-orange-500 to-pink-500 hover:opacity-95 transition-all disabled:opacity-50"
                >
                  {isSavingRepair ? 'Saving to Excel...' : editingRepairIndex !== null ? 'Update Record' : 'Save Record'}
                </button>
              </div>
            </form>
          </div>
        </>
      )}

      {/* ==================================================== */}
      {/* SPARE ADD / EDIT DRAWER                              */}
      {/* ==================================================== */}
      {showSpareDrawer && (
        <>
          <div
            className="fixed inset-0 bg-black/60 backdrop-blur-sm z-40"
            onClick={() => setShowSpareDrawer(false)}
          />
          <div
            className="fixed top-0 right-0 w-[540px] max-w-[95vw] h-full z-50 flex flex-col backdrop-blur-2xl shadow-2xl transition-all"
            style={{ background: 'var(--modal-bg)' }}
          >
            <div className="flex items-center justify-between p-6 border-b border-white/5">
              <div>
                <h2 className="text-base font-extrabold" style={{ color: 'var(--text-main)' }}>
                  {editingSpareId ? 'Edit Spare Part' : 'Add Spare Part'}
                </h2>
                <p className="text-xs" style={{ color: 'var(--text-muted)' }}>
                  Register stock equipment or spare part into inventory.
                </p>
              </div>
              <button
                type="button"
                className="w-8 h-8 rounded-lg flex items-center justify-center transition-all hover:bg-white/10"
                style={{ color: 'var(--text-muted)' }}
                onClick={() => setShowSpareDrawer(false)}
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <form onSubmit={handleSaveSpare} className="flex-1 overflow-y-auto p-6 flex flex-col gap-4">
              <div>
                <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Part Name *</label>
                <input
                  type="text"
                  required
                  value={spareForm.name}
                  onChange={(e) => setSpareForm({ ...spareForm, name: e.target.value })}
                  placeholder="e.g. Test Socket Pin Block"
                  className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                  style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Product Code *</label>
                  <input
                    type="text"
                    required
                    value={spareForm.code}
                    onChange={(e) => setSpareForm({ ...spareForm, code: e.target.value })}
                    placeholder="e.g. SP-MT99-042"
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  />
                </div>

                <div>
                  <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Handler Compatibility *</label>
                  <input
                    type="text"
                    required
                    value={spareForm.handler}
                    onChange={(e) => setSpareForm({ ...spareForm, handler: e.target.value })}
                    placeholder="e.g. MT99, MT93"
                    className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                  />
                </div>
              </div>

              <div>
                <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Stock Quantity *</label>
                <input
                  type="number"
                  required
                  min="0"
                  value={spareForm.quantity}
                  onChange={(e) => setSpareForm({ ...spareForm, quantity: parseInt(e.target.value, 10) || 0 })}
                  className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                  style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                />
              </div>

              <div>
                <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>Product Image</label>
                <div className="flex items-center gap-3">
                  <div
                    className="w-14 h-14 rounded-xl flex items-center justify-center overflow-hidden flex-shrink-0"
                    style={{ background: 'var(--control-bg)', color: 'var(--text-muted)' }}
                  >
                    {spareForm.image ? (
                      <img src={spareForm.image} alt="Preview" className="w-full h-full object-cover" />
                    ) : (
                      <Upload className="w-5 h-5" />
                    )}
                  </div>
                  <label className="cursor-pointer px-4 py-2 rounded-xl text-xs font-semibold flex items-center gap-2 transition-all hover:opacity-90 shadow-sm" style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}>
                    <Upload className="w-3.5 h-3.5" />
                    <span>Upload Image</span>
                    <input
                      type="file"
                      accept="image/*"
                      className="hidden"
                      onChange={async (e) => {
                        const file = e.target.files?.[0];
                        if (file) {
                          const compressed = await compressImage(file, 400, 0.7);
                          setSpareForm((prev) => ({ ...prev, image: compressed }));
                          setSparePhotoFileName(file.name);
                        }
                      }}
                    />
                  </label>
                  <span className="text-xs truncate max-w-xs" style={{ color: 'var(--text-muted)' }}>
                    {sparePhotoFileName || 'No image chosen'}
                  </span>
                </div>
              </div>

              <div className="flex items-center justify-end gap-2 mt-auto pt-6 border-t border-white/5">
                <button
                  type="button"
                  className="px-4 py-2 rounded-xl text-xs font-semibold transition-all hover:bg-white/10"
                  style={{ color: 'var(--text-main)' }}
                  onClick={() => setShowSpareDrawer(false)}
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 rounded-xl text-xs font-bold text-white shadow-lg bg-gradient-to-br from-orange-500 to-pink-500 hover:opacity-95 transition-all"
                >
                  {editingSpareId ? 'Update Part' : 'Save Spare Part'}
                </button>
              </div>
            </form>
          </div>
        </>
      )}

      {/* ==================================================== */}
      {/* STOCK ADJUST MODAL                                   */}
      {/* ==================================================== */}
      {showStockModal && stockModalPart && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div
            className="w-full max-w-sm rounded-2xl p-6 backdrop-blur-2xl shadow-2xl flex flex-col gap-4 animate-in fade-in zoom-in duration-200"
            style={{ background: 'var(--modal-bg)' }}
          >
            <div>
              <h3 className="text-base font-extrabold" style={{ color: 'var(--text-main)' }}>
                {stockModalAction === 'add' ? 'Add / Replenish Stocks' : 'Withdraw Stocks'}
              </h3>
              <p className="text-xs" style={{ color: 'var(--text-muted)' }}>
                {stockModalPart.name} ({stockModalPart.code}) &bull; Current: {stockModalPart.quantity} units
              </p>
            </div>

            <form onSubmit={handleStockAdjustSubmit} className="flex flex-col gap-4">
              <div>
                <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>
                  Quantity to {stockModalAction === 'add' ? 'Add' : 'Withdraw'}
                </label>
                <input
                  type="number"
                  min="1"
                  required
                  value={stockModalQty}
                  onChange={(e) => setStockModalQty(Math.max(1, parseInt(e.target.value, 10) || 1))}
                  className="w-full px-4 py-2.5 rounded-xl text-center text-lg font-extrabold outline-none"
                  style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                />
              </div>

              <div>
                <label className="text-xs font-bold block mb-1.5" style={{ color: 'var(--text-secondary)' }}>
                  Reason / Note (Optional)
                </label>
                <input
                  type="text"
                  value={stockModalReason}
                  onChange={(e) => setStockModalReason(e.target.value)}
                  placeholder="e.g. PM Maintenance MT99"
                  className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                  style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                />
              </div>

              <div className="flex items-center justify-end gap-2 pt-2">
                <button
                  type="button"
                  className="px-4 py-2 rounded-xl text-xs font-semibold transition-all hover:bg-white/10"
                  style={{ color: 'var(--text-main)' }}
                  onClick={() => setShowStockModal(false)}
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className={`px-5 py-2 rounded-xl text-xs font-bold text-white shadow-lg transition-all ${
                    stockModalAction === 'add'
                      ? 'bg-gradient-to-br from-emerald-500 to-teal-600 hover:opacity-95'
                      : 'bg-gradient-to-br from-red-500 to-rose-600 hover:opacity-95'
                  }`}
                >
                  Confirm {stockModalAction === 'add' ? 'Add' : 'Withdraw'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ==================================================== */}
      {/* LOGIN MODAL                                          */}
      {/* ==================================================== */}
      {showLoginModal && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div
            className="w-full max-w-sm rounded-2xl p-6 backdrop-blur-2xl shadow-2xl flex flex-col gap-4 animate-in fade-in zoom-in duration-200"
            style={{ background: 'var(--modal-bg)' }}
          >
            <div>
              <h3 className="text-base font-extrabold" style={{ color: 'var(--text-main)' }}>
                Login to Edit
              </h3>
              <p className="text-xs" style={{ color: 'var(--text-muted)' }}>
                Sign in with credentials to modify records & inventory.
              </p>
            </div>

            {loginError && (
              <div className="p-2.5 rounded-xl text-xs font-semibold bg-red-500/20 text-red-300">
                {loginError}
              </div>
            )}

            <form onSubmit={handleLoginSubmit} className="flex flex-col gap-3">
              <div>
                <label className="text-xs font-bold block mb-1" style={{ color: 'var(--text-secondary)' }}>Username</label>
                <input
                  type="text"
                  required
                  placeholder="admin"
                  value={loginUsername}
                  onChange={(e) => setLoginUsername(e.target.value)}
                  className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                  style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                />
              </div>

              <div>
                <label className="text-xs font-bold block mb-1" style={{ color: 'var(--text-secondary)' }}>Password</label>
                <input
                  type="password"
                  required
                  placeholder="admin123"
                  value={loginPassword}
                  onChange={(e) => setLoginPassword(e.target.value)}
                  className="w-full px-3.5 py-2 rounded-xl text-xs outline-none"
                  style={{ background: 'var(--control-bg)', color: 'var(--text-main)' }}
                />
              </div>

              <div className="flex items-center justify-end gap-2 pt-3">
                <button
                  type="button"
                  className="px-4 py-2 rounded-xl text-xs font-semibold transition-all hover:bg-white/10"
                  style={{ color: 'var(--text-main)' }}
                  onClick={() => {
                    setShowLoginModal(false);
                    setLoginError('');
                  }}
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 rounded-xl text-xs font-bold text-white shadow-lg bg-gradient-to-br from-orange-500 to-pink-500 hover:opacity-95 transition-all"
                >
                  Login
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ==================================================== */}
      {/* CONFIRM MODAL (SAFE REPLACEMENT FOR WINDOW.CONFIRM)  */}
      {/* ==================================================== */}
      {confirmModal.isOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div
            className="w-full max-w-sm rounded-2xl p-6 backdrop-blur-2xl shadow-2xl flex flex-col gap-4 animate-in fade-in zoom-in duration-200"
            style={{ background: 'var(--modal-bg)' }}
          >
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl flex items-center justify-center bg-red-500/20 text-red-400 flex-shrink-0">
                <AlertTriangle className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-base font-extrabold" style={{ color: 'var(--text-main)' }}>
                  {confirmModal.title}
                </h3>
                <p className="text-xs mt-1" style={{ color: 'var(--text-muted)' }}>
                  {confirmModal.message}
                </p>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                type="button"
                className="px-4 py-2 rounded-xl text-xs font-semibold transition-all hover:bg-white/10"
                style={{ color: 'var(--text-main)' }}
                onClick={() => setConfirmModal((prev) => ({ ...prev, isOpen: false }))}
              >
                Cancel
              </button>
              <button
                type="button"
                className="px-5 py-2 rounded-xl text-xs font-bold text-white shadow-lg bg-red-500 hover:bg-red-600 transition-all"
                onClick={confirmModal.onConfirm}
              >
                Confirm Delete
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================== */}
      {/* IMAGE LIGHTBOX MODAL                                 */}
      {/* ==================================================== */}
      {lightboxUrl && (
        <div
          className="fixed inset-0 bg-black/90 backdrop-blur-md z-50 flex items-center justify-center p-6 cursor-pointer"
          onClick={() => setLightboxUrl(null)}
        >
          <img
            src={lightboxUrl}
            alt="Preview Enlarge"
            className="max-w-[90vw] max-h-[85vh] rounded-2xl object-contain shadow-2xl"
          />
        </div>
      )}

      {/* Toast Notification */}
      {toastMessage && (
        <div
          className="fixed bottom-6 right-6 px-5 py-3 rounded-xl text-xs font-bold shadow-2xl z-50 backdrop-blur-xl border border-white/10 animate-in fade-in slide-in-from-bottom-4 duration-300"
          style={{ background: 'var(--modal-bg)', color: 'var(--text-main)' }}
        >
          {toastMessage}
        </div>
      )}
    </div>
  );
}
