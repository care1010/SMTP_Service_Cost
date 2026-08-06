import React, { useState, useEffect, useCallback, useRef } from 'react';
import axios from 'axios';
import DataTable from '../components/DataTable';
import FilterBar from '../components/FilterBar';
import KpiCards from '../components/KpiCards';
import AsblModal from '../components/AsblModal';
import ReviewChanges from './ReviewChanges';
import Swal from 'sweetalert2';
import $ from 'jquery';
import { 
    HiOutlineFilter, 
    HiOutlineRefresh, 
    HiChevronRight, 
    HiOutlineSave, 
    HiOutlineViewGrid,
    HiOutlineUpload,
    HiChevronDown,
    HiOutlineDatabase,
} from "react-icons/hi";

// ─────────────────────────────────────────────────────────────
// Helper: array params properly encode karna
// allowedCustomers ke duplicates bhi yahan dedupe honge
// ─────────────────────────────────────────────────────────────
const buildQueryParams = (filters, extra = {}) => {
    const params = new URLSearchParams();
    Object.keys(filters).forEach(key => {
        const val = filters[key];
        if (Array.isArray(val)) {
            const cleaned = [...new Set(val)].filter(v => v && v !== 'All');
            if (cleaned.length > 0) {
                params.append(key, cleaned.join(','));
                // params.append(key, cleaned.join('|||'));
            }
        } else if (val && val !== 'All') {
            params.append(key, val);
        }
    });
    Object.keys(extra).forEach(k => {
        if (extra[k] !== undefined && extra[k] !== null && extra[k] !== '') {
            params.append(k, extra[k]);
        }
        // if (extra[k] !== undefined && extra[k] !== null) {
        //     // Hum yahan empty string ('') ko bhi set karenge taaki global filter reset ho sake!
        //     params.set(k, extra[k]); 
        // }
    });
    return params;
};

// Dedupe array — allowedCustomers mein duplicates remove karo
const dedupeArray = (arr) => [...new Set(arr || [])];

// ─────────────────────────────────────────────────────────────
// Inline WBS Type Dropdown
// Banner ke andar rahega — select karne ke baad bhi visible
// ─────────────────────────────────────────────────────────────
const WbsTypeInlineDropdown = ({ options = [], selected = [], onChange }) => {
    const [isOpen, setIsOpen] = useState(false);
    const containerRef = useRef(null);

    useEffect(() => {
        const handleOutside = (e) => {
            if (containerRef.current && !containerRef.current.contains(e.target)) {
                setIsOpen(false);
            }
        };
        document.addEventListener('mousedown', handleOutside);
        return () => document.removeEventListener('mousedown', handleOutside);
    }, []);

    const toggleOption = (val) => {
        if (selected.includes(val)) {
            onChange('wbs_type', selected.filter(v => v !== val));
        } else {
            onChange('wbs_type', [...selected, val]);
        }
    };

    const displayText = selected.length > 0 ? selected.join(', ') : 'Select WBS Type ▾';

    return (
        <div ref={containerRef} className="relative inline-block ml-2" style={{ minWidth: '190px' }}>
            <button
                onClick={(e) => { e.stopPropagation(); setIsOpen(prev => !prev); }}
                className={`flex items-center gap-2 px-3 py-1.5 border-2 rounded-lg text-sm font-bold shadow-sm transition-all
                    ${selected.length > 0
                        ? 'bg-blue-600 border-blue-600 text-white hover:bg-blue-700'
                        : 'bg-white border-orange-400 text-orange-700 hover:bg-orange-50'
                    }`}
            >
                <span className="truncate max-w-[150px]">{displayText}</span>
                <HiChevronDown className={`flex-shrink-0 transition-transform duration-200 ${isOpen ? 'rotate-180' : ''}`} />
            </button>

            {isOpen && (
                <div className="absolute top-full left-0 mt-1 bg-white border border-slate-200 rounded-xl shadow-2xl z-[9999] min-w-[200px] py-1 overflow-hidden">
                    {options.length === 0 ? (
                        <div className="px-4 py-3 text-sm text-slate-400 italic">No options available</div>
                    ) : (
                        options.map(opt => {
                            const isSelected = selected.includes(opt);
                            return (
                                <div
                                    key={opt}
                                    onClick={(e) => { e.stopPropagation(); toggleOption(opt); }}
                                    className={`flex items-center gap-3 px-4 py-2.5 text-sm cursor-pointer transition-colors
                                        ${isSelected ? 'bg-blue-50 text-blue-700 font-bold' : 'text-slate-700 hover:bg-slate-50'}`}
                                >
                                    <span className={`w-4 h-4 rounded border-2 flex items-center justify-center flex-shrink-0 transition-all
                                        ${isSelected ? 'bg-blue-600 border-blue-600' : 'border-slate-300'}`}>
                                        {isSelected && (
                                            <svg className="w-2.5 h-2.5 text-white" viewBox="0 0 10 10" fill="none">
                                                <path d="M2 5l2.5 2.5L8 3" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                                            </svg>
                                        )}
                                    </span>
                                    <span>{opt}</span>
                                </div>
                            );
                        })
                    )}
                    {selected.length > 0 && (
                        <div className="border-t border-slate-100 mt-0.5">
                            <button
                                onClick={(e) => { e.stopPropagation(); onChange('wbs_type', []); setIsOpen(false); }}
                                className="w-full text-left px-4 py-2 text-xs text-red-500 font-bold hover:bg-red-50 transition-colors"
                            >
                                ✕ Clear Selection
                            </button>
                        </div>
                    )}
                </div>
            )}
        </div>
    );
};


// ═══════════════════════════════════════════════════
// MAIN COMPONENT
// Props: filters, onFilterChange, onResetFilters — App.js se aate hain
// ═══════════════════════════════════════════════════
const SummaryView = ({ user, filters, onFilterChange, onResetFilters, setActiveTab }) => {
    const [options, setOptions] = useState({});
    const [kpiData, setKpiData] = useState(null);
    const [isModalOpen, setIsModalOpen] = useState(false);
    const [showAll, setShowAll] = useState(false); 
    const [loading, setLoading] = useState(false);
    const [isReviewMode, setIsReviewMode] = useState(false);
    const [collapseView, setCollapseView] = useState(false);
    const [isSidebarOpen, setIsSidebarOpen] = useState(false);

    // Dedupe allowedCustomers — login ke time duplicate values aa sakti hain
    const allowedCustomers = dedupeArray(user?.allowedCustomers);

    // DEBOUNCED + ABORTABLE FILTER OPTIONS FETCH
    // 500ms debounce + AbortController:
    // User 3 filters rapidly select kare -> sirf LAST request backend tak pahunche
    // Purani in-flight requests cancel ho jaati hain -> no pool exhaustion
    const debounceTimer = useRef(null);
    const abortControllerRef = useRef(null);

    useEffect(() => {
        if (!user) return;

        // Clear pending debounce
        if (debounceTimer.current) clearTimeout(debounceTimer.current);

        // Cancel any in-flight request immediately
        if (abortControllerRef.current) {
            abortControllerRef.current.abort();
        }

        debounceTimer.current = setTimeout(async () => {
            // New AbortController for this request
            abortControllerRef.current = new AbortController();
            try {
                const params = buildQueryParams(filters, {
                    type: user?.type,
                    allowedCustomers: allowedCustomers.join(',')
                });
                const res = await axios.get(
                    `${process.env.REACT_APP_API_URL}/api/data/filter-options?${params.toString()}`,
                    { signal: abortControllerRef.current.signal }
                );
                setOptions(res.data);
            } catch (err) {
                if (axios.isCancel(err) || err.name === 'CanceledError' || err.name === 'AbortError') {
                    return;
                }
                console.error('Filter options fetch error:', err.message);
            }
        }, 500);

        return () => {
            if (debounceTimer.current) clearTimeout(debounceTimer.current);
            if (abortControllerRef.current) abortControllerRef.current.abort();
        };
    }, [user, JSON.stringify(filters)]); // eslint-disable-line

    // ─── GLOBAL SAVE HANDLER ───
    const handleSave = async () => {
        const updates = [];
        $('.nc-input.is-changed').each(function () {
            updates.push({
                loa_name: $(this).data('loa'),
                categories: $(this).data('cat'),
                wbs_type: $(this).data('wbstype'),
                value: $(this).val()
            });
        });

        if (updates.length === 0) return Swal.fire("Info", "No changes to save.", "info");

        const result = await Swal.fire({
            title: "Save Changes?",
            text: `Confirm saving ${updates.length} modified records to draft.`,
            icon: "question",
            showCancelButton: true,
            confirmButtonText: "Save",
            confirmButtonColor: "#2563eb"
        });

        if (!result.isConfirmed) return;
        
        Swal.fire({ title: "Saving...", allowOutsideClick: false, didOpen: () => Swal.showLoading() });

        try {
            await axios.post(`${process.env.REACT_APP_API_URL}/api/data/update-non-committed`, {
                updates, 
                createdBy: user?.name || user?.email || 'System'
            });
            await Swal.fire("Success", "Draft updated successfully.", "success");
            window.location.reload(); 
        } catch (err) {
            Swal.fire("Error", "Save failed. Please try again.", "error");
        }
    };

    // ─── ACTION HANDLERS ───
    const handleReviewClick = async () => {
        try {
            const res = await axios.get(`${process.env.REACT_APP_API_URL}/api/data/check-pending-changes`);
            if (res.data.count > 0) setIsReviewMode(true);
            else Swal.fire({ icon: 'info', title: 'No Changes', text: 'No pending changes to review.', confirmButtonColor: '#3b82f6' });
        } catch (err) { console.error(err); }
    };

    const handleFullExport = async () => {
        const result = await Swal.fire({
            title: 'Select Export View',
            icon: 'question',
            showDenyButton: true,
            showCancelButton: true,
            confirmButtonText: 'Cost Level',
            denyButtonText: 'Element Level',
            confirmButtonColor: '#16a34a',
            denyButtonColor: '#4169e1'
        });
        if (result.isDismissed) return;

        const exportUrl = new URL(`${process.env.REACT_APP_API_URL}/api/data/export-excel`);
        exportUrl.searchParams.append('type', user?.type || 'user');
        exportUrl.searchParams.append('allowedCustomers', allowedCustomers.join(','));
        exportUrl.searchParams.append('showAll', showAll);
        exportUrl.searchParams.append('collapseView', result.isConfirmed);

        // 🔥 FIX: Multi-select arrays properly encode karo
        Object.keys(filters).forEach(key => {
            const val = filters[key];
            if (Array.isArray(val)) {
                const cleaned = val.filter(v => v && v !== 'All');
                if (cleaned.length > 0) exportUrl.searchParams.append(key, cleaned.join(','));
            } else if (val && val !== 'All') {
                exportUrl.searchParams.append(key, val);
            }
        });

        window.location.href = exportUrl.toString();
    };

    const handleFullRefresh = async () => {
        const result = await Swal.fire({ title: "Sync Database?", text: "Takes 1-2 minutes.", icon: "warning", showCancelButton: true, confirmButtonText: "Yes, Sync" });
        if (!result.isConfirmed) return;
        Swal.fire({ title: "Syncing...", allowOutsideClick: false, didOpen: () => Swal.showLoading() });
        try {
            await axios.post(`${process.env.REACT_APP_API_URL}/api/data/full-refresh`);
            await Swal.fire("Success", "Database Synchronized!", "success");
            window.location.reload();
        } catch (err) { Swal.fire("Error", "Sync Failed", "error"); }
    };

    const handleKpiUpdate = useCallback((data) => setKpiData(data), []);

    // ─── BUILD API URL FOR DATATABLE ───
    // 🔥 FIX: Multi-select arrays ko comma join karke bhejo — blank table bug fix
    const queryParams = buildQueryParams(filters, {
        showAll,
        type: user?.type,
        allowedCustomers: allowedCustomers.join(',')
    });

    const dynamicApiUrl = collapseView
        ? `${process.env.REACT_APP_API_URL}/api/data/wbs-summary-collapse?${queryParams.toString()}`
        : `${process.env.REACT_APP_API_URL}/api/data/wbs-summary?${queryParams.toString()}`;

    // ─── TABLE COLUMNS ───
    const tableColumns = [
        { header: 'BU', field: 'bu' }, 
        { header: 'Customer', field: 'customer' }, 
        { header: 'LOA Name', field: 'loa_name' }, 
        { header: 'LOA ID', field: 'loa_id' }, 
        { header: 'Cost / Revenue', field: 'cost_revenue' }, 
        { header: 'Category', field: 'categories' }, 
        { header: 'ASBL', field: 'asbl' }, 
        // { header: 'ASBL LOA', field: 'asbl_loa' }, 
        { header: 'PTD', field: 'ptd', clickable: true }, 
        { header: 'Open Commitment', field: 'open_commitment_KEUR', clickable: true }, 
        { header: 'Non Committed', field: 'non_committed_editable' }, 
        { header: 'EAC', field: 'eac' }, 
        { header: 'EAC vs ASBL', field: 'eac_vs_asbl' }
    ];

    // ─── WBS TYPE warning banner check ───
    // Banner hamesha show hoga — sirf warning text tab hide hoga
    // jab valid WBS type select ho (Project ya AMC — Warranty/Other exclude)
    const wbsTypeSelected = filters.wbs_type &&
        filters.wbs_type.length > 0 &&
        !filters.wbs_type.includes('All') &&
        filters.wbs_type.some(v => !v.toLowerCase().includes('warranty/other'));

    const wbsTypeOptions = options?.wbs_type || [];

    if (isReviewMode) {
    return (
        <ReviewChanges 
            user={user}
            filters={filters} 
            options={options} 
            onFilterChange={onFilterChange} 
            onResetFilters={onResetFilters}
            onBack={() => setIsReviewMode(false)} 
        />
    );
}

    return (
        <div className="flex bg-[#f8fafc] min-h-screen relative overflow-hidden">
            
            {/* 🟢 MAIN CONTENT AREA */}
            <div className={`flex-1 p-5 transition-all duration-300 ${isSidebarOpen ? 'mr-[380px]' : 'mr-[40px]'}`}>
                
                {loading && (
                    <div className="fixed inset-0 z-[3000] bg-slate-900/80 backdrop-blur-md flex flex-col items-center justify-center text-white">
                        <div className="w-20 h-20 border-4 border-orange-500 border-t-transparent rounded-full animate-spin mb-4"></div>
                        <h2 className="text-xl font-bold">Processing Sync...</h2>
                    </div>
                )}

                {/* KPI & ACTIONS HEADER */}
                <div className="flex flex-col lg:flex-row gap-4 mb-6 items-stretch">
                    <div className="flex-1"><KpiCards data={kpiData} /></div>
                    
                    <div className="flex gap-2 items-center">
                        {/* 🔥 NAYA BUTTON: Navigate to Raw Drill */}
                        <button 
                            onClick={() => setActiveTab('cj74/cji5')} 
                            className="border border-slate-300 border-t-4 border-t-purple-500 bg-white px-5 py-2 shadow-sm hover:shadow-md transition-all flex items-center gap-2 rounded-lg"
                        >
                            <HiOutlineDatabase className="text-purple-600 text-lg" /> 
                            <span className="text-sm font-semibold text-purple-700">Raw Data</span>
                        </button>

                        <button onClick={handleFullExport} className="border border-slate-300 border-t-4 border-t-blue-500 bg-white px-5 py-2 shadow-sm hover:shadow-md transition-all flex items-center gap-2 rounded-lg">
                            <HiOutlineUpload className="text-blue-600" /> 
                            <span className="text-sm font-semibold text-blue-700">Export</span>
                        </button>

                        {/* Existing code ko replace karein is logic se */}
                        {(user?.type === 'admin' || user?.type === 'super_admin') && (
                            <>
                                {/* REVIEW button sirf Super Admin ko dikhega */}
                                {user?.type === 'super_admin' && (
                                    <button onClick={handleReviewClick} className="border border-slate-300 border-t-4 border-t-green-600 bg-white px-5 py-2 shadow-sm hover:shadow-md transition-all flex items-center gap-2 rounded-lg">
                                        <HiOutlineRefresh className="text-green-700" /> 
                                        <span className="text-sm font-semibold text-green-700">Review</span>
                                    </button>
                                )}

                                {/* SYNC DB button Admin aur Super Admin dono ko dikhta rahega (Existing Logic) */}
                                <button onClick={handleFullRefresh} className="border border-slate-300 border-t-4 border-t-slate-800 bg-white px-5 py-2 shadow-sm hover:shadow-md transition-all flex items-center gap-2 rounded-lg">
                                    <span className="text-sm font-semibold text-slate-800">Sync DB</span>
                                </button>
                            </>
                        )}

                        <button onClick={handleSave} className="border border-slate-800 border-t-4 border-t-black-600 bg-white px-5 py-2 shadow-sm hover:shadow-md transition-all flex items-center gap-2 rounded-lg">
                            <HiOutlineSave className="text-black-700" /> 
                            <span className="text-sm font-bold text-black">Save</span>
                        </button>
                    </div>
                </div>

                {/* ─── WBS TYPE BANNER — hamesha visible, dropdown always synced with FilterBar ─── */}
                <div className={`mb-6 p-4 rounded-3xl text-sm flex flex-wrap items-center gap-3 shadow-sm border transition-colors duration-300
                    ${wbsTypeSelected
                        ? 'border-green-200 bg-green-50/80 text-green-800'
                        : 'border-orange-200 bg-orange-50/80 text-orange-800'
                    }`}>
                    <span className="text-lg flex-shrink-0">{wbsTypeSelected ? '✅' : '⚠️'}</span>
                    <div className="flex flex-wrap items-center gap-2 flex-1">
                        {!wbsTypeSelected && (
                            <span className="font-extrabold uppercase tracking-wide">
                                Please select a specific WBS Type to unlock ASBL &amp; Non Committed values.
                            </span>
                        )}
                        {wbsTypeSelected && (
                            <span className="font-bold uppercase tracking-wide">
                                WBS Type selected: <strong>{filters.wbs_type.join(', ')}</strong>
                            </span>
                        )}
                        {/* 🔥 Inline WBS Type Dropdown — HAMESHA visible, FilterBar ke saath fully synced */}
                        <WbsTypeInlineDropdown
                            options={wbsTypeOptions}
                            selected={filters.wbs_type || []}
                            onChange={onFilterChange}
                        />
                    </div>
                </div>

                {/* ─── DATA TABLE ─── */}
                <div className="rounded-[1.5rem] overflow-hidden shadow-xl border border-white bg-white w-full">
                    <DataTable 
                        title="" 
                        columns={tableColumns} 
                        apiUrl={dynamicApiUrl} 
                        filters={filters} 
                        onKpiUpdate={handleKpiUpdate} 
                        collapseView={collapseView} 
                        showSaveButton={false}
                        user={user}
                    />
                </div>

                {/* BOTTOM NOTE */}
                <div className="mt-6 mb-4 px-5 py-4 border border-amber-200 rounded-2xl text-[13px] text-slate-800 bg-amber-50/60 shadow-sm leading-relaxed">
                    <span className="font-black text-amber-800 uppercase tracking-tighter mr-2">Note:</span>
                    This tool is to be used to track Services Cost "– EAC vs ASBL". 
                    Please ignore revenue figures, as these figures are not validated.
                </div>
            </div>

            {/* 🔵 POWER BI STYLE SIDEBAR */}
            <div className={`fixed right-0 top-0 h-full bg-white border-l border-slate-200 transition-all duration-300 z-[2001] shadow-2xl flex ${isSidebarOpen ? 'w-[380px]' : 'w-[40px]'}`}>
                
                <div 
                    onClick={() => setIsSidebarOpen(!isSidebarOpen)}
                    className={`h-full flex flex-col items-center pt-8 cursor-pointer hover:bg-slate-50 transition-colors ${isSidebarOpen ? 'w-[40px] border-r border-slate-100' : 'w-full'}`}
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
                    {isSidebarOpen && <HiChevronRight className="text-slate-300 mt-auto mb-10 text-xl" />}
                </div>

                {isSidebarOpen && (
                    <div className="flex-1 flex flex-col animate-in fade-in duration-300">
                        <div className="p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                            <span className="font-black text-lg text-slate-800 tracking-tight">Filters Pane</span>
                            <button onClick={onResetFilters} className="text-[11px] font-black uppercase text-red-500 hover:underline">Reset All</button>
                        </div>
                        <div className="flex-1 overflow-y-auto p-6 custom-scrollbar">
                            <FilterBar
                                filters={filters}
                                options={options}
                                onFilterChange={onFilterChange}
                                onReset={onResetFilters}
                            />
                        </div>
                        <div className="p-5 border-t border-slate-100 bg-white shadow-[0_-4px_20px_rgba(0,0,0,0.03)]">
                            <button
                                onClick={() => setIsSidebarOpen(false)}
                                className="w-full py-4 bg-blue-600 text-white rounded-2xl font-black text-xs uppercase shadow-lg shadow-blue-100 hover:bg-blue-700 transition-all active:scale-95"
                            >
                                Apply Filters
                            </button>
                        </div>
                    </div>
                )}
            </div>

            <AsblModal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} onSubmit={() => setIsModalOpen(false)} />
        </div>
    );
};

export default SummaryView;