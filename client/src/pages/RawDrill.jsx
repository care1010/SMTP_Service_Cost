import React, { useState, useEffect, useRef, useCallback } from 'react';
import axios from 'axios';
import FilterBar from '../components/FilterBar';
import { HiOutlineFilter, HiOutlineUpload, HiDatabase, HiChevronRight } from "react-icons/hi";
import $ from 'jquery';
import 'datatables.net-dt';
import 'datatables.net-dt/css/dataTables.dataTables.css';

// ── CJ74 columns ──
const CJ74_COLS = [
    { data: 'sap_wbs',                  title: 'SAP WBS' },
    { data: 'loa_id',                   title: 'LOA ID' },
    { data: 'year',                     title: 'Year' },
    { data: 'per',                      title: 'Period No' },
    { data: 'period',                   title: 'Period' },
    { data: 'cost_element',             title: 'Cost Element' },
    { data: 'cost_element_name',        title: 'CE Name' },
    { data: 'ptd_val',                  title: 'PTD (K€)', className: 'text-right' },
    { data: 'cost_element_descr',       title: 'CE Description' },
    { data: 'wbs_type',                 title: 'WBS Type' },
    { data: 'categories',               title: 'Category' },
    { data: 'cost_revenue',             title: 'Cost/Revenue' },
    { data: 'refdocno',                 title: 'Ref Doc No' },
    { data: 'document_no',              title: 'Document No' },
    { data: 'doc_date',                 title: 'Doc Date' },
    { data: 'postg_date',               title: 'Posting Date' },
    { data: 'created_on',               title: 'Created On' },
    { data: 'user_name',                title: 'User Name' },
    { data: 'pur_doc',                  title: 'Pur Doc' },
    { data: 'purchase_order_text',      title: 'PO Text' },
    { data: 'material',                 title: 'Material' },
    { data: 'material_description',     title: 'Material Desc' },
    { data: 'cocd',                     title: 'CoCd' },
    { data: 'proj_def',                 title: 'Project Def' },
    { data: 'profit_ctr',               title: 'Profit Ctr' },
    { data: 'tcurr',                    title: 'Trans Curr' },
    { data: 'val_in_rc',                title: 'Val in RC' },
    { data: 'offst_acct',              title: 'Offset Acct' },
    { data: 'name_of_offsetting_account', title: 'Offset Acct Name' },
];

// ── CJI5 columns ──
const CJI5_COLS = [
    { data: 'sap_wbs',          title: 'SAP WBS' },
    { data: 'loa_id',           title: 'LOA ID' },
    { data: 'year',             title: 'Year' },
    { data: 'per',              title: 'Period No' },
    { data: 'cost_element',     title: 'Cost Element' },
    { data: 'cost_element_descr', title: 'CE Description' },
    { data: 'oc_val',           title: 'OC (K€)', className: 'text-right' },
    { data: 'wbs_type',         title: 'WBS Type' },
    { data: 'categories',       title: 'Category' },
    { data: 'project_def',      title: 'Project Def' },
    { data: 'refdocno',         title: 'Ref Doc No' },
    { data: 'supplier',         title: 'Supplier' },
    { data: 'name',             title: 'Name' },
    { data: 'co_object_name',   title: 'CO Object' },
    { data: 'item',             title: 'Item' },
    { data: 'matl_group',       title: 'Material Group' },
    { data: 'material',         title: 'Material' },
    { data: 'description',      title: 'Description' },
    { data: 'quantity',         title: 'Quantity' },
    { data: 'debit_date',       title: 'Debit Date' },
    { data: 'doc_date',         title: 'Doc Date' },
    { data: 'user_name',        title: 'User Name' },
    { data: 'docc',             title: 'Doc C' },
    { data: 'cocode',           title: 'Co Code' },
    { data: 'report_currency',  title: 'Report Currency' },
    { data: 'val_in_rep_cur',   title: 'Val in Rep Cur' },
    { data: 'tcurr',            title: 'Trans Curr' },
    { data: 'exch_rate',        title: 'Exch Rate' },
];

const fmt = (val) => {
    if (val === null || val === undefined || val === '') return '-';
    const num = Number(val);
    if (isNaN(num)) return val;
    return num.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};

const fmtDate = (val) => {
    if (!val) return '-';
    try { return new Date(val).toLocaleDateString('en-GB'); } catch { return val; }
};

// ════════════════════════════════════════════════════
// MAIN COMPONENT
// ════════════════════════════════════════════════════
const RawDrill = ({ user, filters, onFilterChange, onResetFilters }) => {
    const [activeTab, setActiveTab]     = useState('cj74');
    const [isSidebarOpen, setIsSidebarOpen] = useState(false);
    const [options, setOptions]         = useState({});

    const tableRef          = useRef(null);
    const lastDrawRef       = useRef(0);
    const dtInstance        = useRef(null);
    const filtersRef        = useRef(filters);
    const prevFiltersStr    = useRef('');
    const isMounted         = useRef(false);
    const [totalValue, setTotalValue]   = useState(0);

    // Keep filtersRef in sync
    useEffect(() => { filtersRef.current = filters; }, [filters]);

    // ── Build query params from filters ──
    const buildParams = useCallback((extraParams = {}) => {
        const p = {};
        Object.keys(filters).forEach(key => {
            const val = filters[key];
            if (Array.isArray(val)) {
                const cleaned = val.filter(v => v && v !== 'All');
                if (cleaned.length > 0) p[key] = cleaned.join(',');
            } else if (val && val !== 'All') {
                p[key] = val;
            }
        });
        return { ...p, ...extraParams };
    }, [filters]);

    // ── Fetch filter options (synced with shared filters) ──
    useEffect(() => {
        if (!user) return;
        const controller = new AbortController();
        const timer = setTimeout(async () => {
            try {
                const params = new URLSearchParams(buildParams({
                    type: user?.type,
                    allowedCustomers: (user?.allowedCustomers || []).join(',')
                }));
                const res = await axios.get(
                    `${process.env.REACT_APP_API_URL}/api/data/filter-options?${params.toString()}`,
                    { signal: controller.signal }
                );
                setOptions(res.data);
            } catch (err) {
                if (axios.isCancel(err) || err.name === 'CanceledError' || err.name === 'AbortError') return;
                console.error('Filter fetch error:', err.message);
            }
        }, 400);
        return () => { clearTimeout(timer); controller.abort(); };
    }, [user, JSON.stringify(filters)]); // eslint-disable-line

    // ── Initialize / Reinitialize DataTable ──
    useEffect(() => {
        if (!tableRef.current) return;

        // Destroy old instance
        if (dtInstance.current) {
            dtInstance.current.destroy();
            dtInstance.current = null;
        }

        // Clear table
        $(tableRef.current).empty();

        const columns = activeTab === 'cj74' ? CJ74_COLS : CJI5_COLS;

        dtInstance.current = $(tableRef.current).DataTable({
            serverSide:  true,
            processing:  true,
            scrollX:     true,
            scrollY:     'calc(100vh - 320px)',
            scrollCollapse: true,
            autoWidth:   false,
            pageLength:  50,
            lengthMenu:  [25, 50, 100, 200],
            dom: '<"flex justify-between items-center mb-4"lf>rt<"flex justify-between items-center mt-4"ip>',
            ajax: {
                url:  `${process.env.REACT_APP_API_URL}/api/data/raw-get-data`,
                type: 'GET',
                data: (d) => {
                    const filterParams = {};
                    const currentFilters = filtersRef.current;
                    Object.keys(currentFilters).forEach(key => {
                        const val = currentFilters[key];
                        if (Array.isArray(val)) {
                            const cleaned = val.filter(v => v && v !== 'All');
                            if (cleaned.length > 0) filterParams[key] = cleaned.join(',');
                        } else if (val && val !== 'All') {
                            filterParams[key] = val;
                        }
                    });
                    return {
                        ...d,
                        ...filterParams,
                        tableType: activeTab,
                        type:      user?.type,
                        allowedCustomers: (user?.allowedCustomers || []).join(',')
                    };
                },
                // 🔥 NEW: Intercept backend response to set Total Value Card
                dataSrc: function (json) {
                    const currentDraw = parseInt(json.draw) || 0;
                    // Sirf tabhi card update karo jab request latest ho!
                    if (currentDraw >= lastDrawRef.current) {
                        lastDrawRef.current = currentDraw;
                        setTotalValue(json.totalValue || 0);
                    }
                    return json.data || [];
                },
                error: (xhr) => {
                    console.error('DataTable Ajax Error:', xhr.responseJSON?.error || 'Unknown error');
                }
            },
            columns: columns.map(col => ({
                title:          col.title,
                data:           col.data,
                defaultContent: '-',
                className:      col.className || 'text-left',
                width:          col.data === 'sap_wbs' || col.data === 'purchase_order_text' || col.data === 'name_of_offsetting_account' ? '200px' : '120px',
                render: (data, type) => {
                    if (type !== 'display') return data;
                    // Numeric columns
                    if (col.data === 'ptd_val' || col.data === 'oc_val') {
                        return `<span class="font-bold text-blue-700">${fmt(data)}</span>`;
                    }
                    // Date columns
                    if (col.data.includes('date') || col.data === 'created_on' || col.data === 'postg_date') {
                        return fmtDate(data);
                    }
                    if (data === null || data === undefined || data === '') return '-';
                    // Long text truncate
                    const str = String(data);
                    if (str.length > 40) {
                        return `<span title="${str.replace(/"/g, '&quot;')}">${str.substring(0, 38)}…</span>`;
                    }
                    return str;
                }
            })),
            language: {
                processing:  '<div class="text-blue-600 font-bold py-4">Loading data...</div>',
                zeroRecords: '<div class="text-slate-400 font-bold py-8 text-center">No records found for selected filters.</div>',
                emptyTable:  '<div class="text-slate-400 font-bold py-8 text-center">No data available.</div>',
            }
        });

        prevFiltersStr.current = JSON.stringify(filters);
        isMounted.current = true;

        return () => {
            if (dtInstance.current) {
                dtInstance.current.destroy();
                dtInstance.current = null;
            }
        };
    }, [activeTab]);

    // ── Reload on filter change (no reinit) ──
    useEffect(() => {
        if (!isMounted.current) return;
        const currentStr = JSON.stringify(filters);
        if (prevFiltersStr.current === currentStr) return;
        prevFiltersStr.current = currentStr;

        if (dtInstance.current) {
            dtInstance.current.ajax.reload(null, false);
        }
    }, [filters]);

    // ── Export handler ──
    const handleExport = () => {
        const params = new URLSearchParams(buildParams({
            tableType:        activeTab,
            type:             user?.type || '',
            allowedCustomers: (user?.allowedCustomers || []).join(',')
        }));
        window.location.href = `${process.env.REACT_APP_API_URL}/api/data/raw-export?${params.toString()}`;
    };

    return (
        <div className="flex bg-[#f8fafc] min-h-screen relative">

            {/* ── MAIN CONTENT ── */}
            <div className={`flex-1 p-6 transition-all duration-300 ${isSidebarOpen ? 'mr-[380px]' : 'mr-[40px]'} overflow-hidden`}>

                {/* Header */}
                <div className="flex justify-between items-center mb-5">
                    <div className="flex items-center gap-3">
                        <div className="bg-slate-800 p-2.5 rounded-xl text-white shadow-lg">
                            <HiDatabase size={22} />
                        </div>
                        <div>
                            <h1 className="text-2xl font-black text-slate-800 uppercase tracking-tight">Line Item Explorer</h1>
                            <p className="text-xs text-slate-400 font-semibold mt-0.5">Raw data from SAP CJ74 & CJI5 tables</p>
                        </div>
                    </div>
                    <button
                        onClick={handleExport}
                        className="bg-white border-b-4 border-blue-500 shadow-md px-6 py-2.5 rounded-xl font-bold flex items-center gap-2 hover:bg-slate-50 transition-all active:scale-95 text-sm"
                    >
                        <HiOutlineUpload className="text-blue-600 text-lg" />
                        <span className="text-blue-700">Export to Excel</span>
                    </button>
                </div>

                {/* Tabs */}
                <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-5 gap-4">
                    {/* Left Side: Tabs */}
                    <div className="flex gap-2 mb-5 bg-slate-100 p-1.5 rounded-2xl w-fit border border-slate-200 shadow-sm">
                        <button
                            onClick={() => { setActiveTab('cj74'); setTotalValue(0); lastDrawRef.current = 0; }}
                            className={`px-8 py-2.5 rounded-xl font-black text-sm transition-all duration-200
                                ${activeTab === 'cj74'
                                    ? 'bg-white shadow-md text-blue-600 border border-blue-100'
                                    : 'text-slate-500 hover:text-slate-700'
                                }`}
                        >
                            📊 CJ74 — Actuals (PTD)
                        </button>
                        <button
                            onClick={() => { setActiveTab('cji5'); setTotalValue(0); lastDrawRef.current = 0; }}
                            className={`px-8 py-2.5 rounded-xl font-black text-sm transition-all duration-200
                                ${activeTab === 'cji5'
                                    ? 'bg-white shadow-md text-blue-600 border border-blue-100'
                                    : 'text-slate-500 hover:text-slate-700'
                                }`}
                        >
                            📋 CJI5 — Commitments (OC)
                        </button>
                    </div>

                    {/* 🔥 Right Side: Dynamic Power BI Style Card */}
                    <div className="bg-white border-2 border-slate-200 shadow-sm px-6 py-2.5 rounded-xl flex items-center gap-4 min-w-[250px]">
                        <div>
                            <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                                {activeTab === 'cj74' ? 'Total Actuals (PTD)' : 'Total Commitments (OC)'}
                            </p>
                            <p className={`text-xl font-black ${activeTab === 'cj74' ? 'text-emerald-600' : 'text-blue-600'}`}>
                                {fmt(totalValue)} <span className="text-xs text-slate-400 font-bold ml-1">K€</span>
                            </p>
                        </div>
                    </div>
                </div>

                {/* Table Card */}
                <div className="bg-white rounded-[1.5rem] shadow-xl border border-slate-100 overflow-hidden">
                    {/* Table wrapper — scrollX DataTables handles karta hai */}
                    <div className="p-4">
                        <table
                            ref={tableRef}
                            className="display nowrap cell-border pbi-table"
                            style={{ width: '100%' }}
                        />
                    </div>
                </div>
            </div>

            {/* ── SIDEBAR FILTER PANEL ── */}
            <div className={`fixed right-0 top-0 h-screen bg-white border-l border-slate-200 transition-all duration-300 z-[2001] shadow-2xl flex flex-col
                ${isSidebarOpen ? 'w-[380px]' : 'w-[40px]'}`}>

                {/* Toggle handle */}
                <div
                    onClick={() => setIsSidebarOpen(prev => !prev)}
                    className={`h-full flex flex-col items-center pt-8 cursor-pointer hover:bg-slate-50 transition-colors
                        ${isSidebarOpen ? 'w-[40px] border-r border-slate-100' : 'w-full'}`}
                >
                    <HiOutlineFilter className="text-xl mb-4 text-blue-600" />
                    {!isSidebarOpen && (
                        <span
                            className="font-black text-[13px] tracking-[0.2em] text-slate-700 uppercase"
                            style={{ writingMode: 'vertical-lr', textOrientation: 'mixed' }}
                        >
                            Filters
                        </span>
                    )}
                    {isSidebarOpen && (
                        <HiChevronRight className="text-slate-300 mt-auto mb-10 text-xl" />
                    )}
                </div>

                {/* Sidebar content */}
                {isSidebarOpen && (
                    <div className="flex-1 flex flex-col animate-in fade-in duration-200" style={{ width: '340px', position: 'absolute', left: '40px', top: 0, height: '100%', background: 'white' }}>
                        <div className="p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                            <span className="font-black text-lg text-slate-800 tracking-tight">Filters Pane</span>
                            <button onClick={onResetFilters} className="text-[11px] font-black uppercase text-red-500 hover:underline">
                                Reset All
                            </button>
                        </div>
                        <div className="flex-1 overflow-y-auto p-6 custom-scrollbar">
                            <FilterBar
                                filters={filters}
                                options={options}
                                onFilterChange={onFilterChange}
                                onReset={onResetFilters}
                            />
                        </div>
                        <div className="p-5 border-t border-slate-100 bg-white">
                            <button
                                onClick={() => setIsSidebarOpen(false)}
                                className="w-full py-3.5 bg-blue-600 text-white rounded-2xl font-black text-xs uppercase shadow-lg hover:bg-blue-700 transition-all active:scale-95"
                            >
                                Apply Filters
                            </button>
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
};

export default RawDrill;