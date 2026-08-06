import React, { useState, useEffect } from 'react';
import { Routes, Route, useLocation } from 'react-router-dom'; 
import Sidebar from './components/Sidebar';
import SummaryView from './pages/SummaryView';
import AddProject from './pages/AddProject';
import PtdAutomation from './pages/PtdAutomation';
import AsblAutomation from './pages/AsblAutomation';
import Dashboard from './pages/Dashboard';
import Login from './pages/Login';
import AdminPanel from './pages/AdminPanel';
import DrillDownPage from './pages/DrillDownPage';
import MyAccess from './pages/MyAccess';
import Logs from './pages/Logs';
import ERPResource from './pages/ERPResource';
import RequestAccess from './pages/RequestAccess';
import AccessRequestsTable from './pages/AccessRequestsTable';
import RawDrill from './pages/RawDrill'

// ============================================================
// DEFAULT FILTERS — ek jagah define karo, dono pages use karein
// ============================================================
const DEFAULT_FILTERS = {
  category_type: ['All'],
  bu: [],
  customer: [],
  loa_id: [],
  loa_name: [],
  wbs_type: [],
  wbs: [],
  wbs_description: [],
  active_inactive: ['Active'],
  period: []
};

function App() {
  const [user, setUser] = useState(null);
  const [activeTab, setActiveTab] = useState('summary');
  const location = useLocation(); 

  // ============================================================
  // 🔥 SHARED FILTERS STATE — lifted up from SummaryView & Dashboard
  // Dono pages ab ek hi filters state share karenge (Power BI style!)
  // ============================================================
  const [sharedFilters, setSharedFilters] = useState(DEFAULT_FILTERS);

  // Single handler — dono pages ko pass hoga
  const handleFilterChange = (name, value) => {
    setSharedFilters(prev => ({ ...prev, [name]: value }));
  };

  // Reset handler — DEFAULT_FILTERS pe wapas le jaata hai
  const handleResetFilters = () => {
    setSharedFilters(DEFAULT_FILTERS);
  };

  useEffect(() => {
    const savedUser = localStorage.getItem('user');
    if (savedUser) setUser(JSON.parse(savedUser));
  }, []);

  // Role Protection
  useEffect(() => {
    if (user) {
      if (user.type === 'user' && activeTab === 'admin') {
        setActiveTab('summary');
      }
    }
  }, [activeTab, user]);

  const handleLogout = () => {
    localStorage.removeItem('user');
    setUser(null);
    setActiveTab('summary');
    setSharedFilters(DEFAULT_FILTERS); // logout pe filters bhi reset
  };

  if (!user) {
    return <Login onLoginSuccess={(userData) => setUser(userData)} />;
  }

  const isDrillDown = location.pathname === '/drilldown';

  return (
    <div className="flex bg-slate-50 min-h-screen">
      {!isDrillDown && (
        <Sidebar activeTab={activeTab} setActiveTab={setActiveTab} user={user} onLogout={handleLogout} />
      )}

      <main className={`flex-1 ${isDrillDown ? 'ml-0' : ''} p-8 bg-[#fcfcfd] min-h-screen overflow-x-hidden`}
        style={!isDrillDown ? { marginLeft: '130px', width: "calc(100vw - 130px)" } : {}}>
        
        {!isDrillDown && (
            <div className="mb-8 text-center">
                <h1 className="text-3xl font-semibold text-slate-800 tracking-tight">
                    NI INDIA Financial Cost Tracker
                </h1>
            </div>
        )}

        <Routes>
          <Route path="/" element={
            <>
              {activeTab === 'summary' && (
                <SummaryView
                  user={user}
                  filters={sharedFilters}
                  onFilterChange={handleFilterChange}
                  onResetFilters={handleResetFilters}
                  setActiveTab={setActiveTab}
                />
              )}
              {activeTab === 'add-project' && <AddProject user={user} />}
              {activeTab === 'ptd' && <PtdAutomation />}
              {activeTab === 'asbl' && <AsblAutomation user={user} />}
              {activeTab === 'dashboard' && (
                <Dashboard
                  user={user}
                  filters={sharedFilters}
                  onFilterChange={handleFilterChange}
                  onResetFilters={handleResetFilters}
                />
              )}
              {activeTab === 'erp_resource' && <ERPResource />}
              {activeTab === 'admin' && (user?.type === 'super_admin' || user?.type === 'admin') && (
                <AdminPanel user={user} onBack={() => setActiveTab('summary')} />
              )}
              {activeTab === 'my-access' && (<MyAccess user={user} />)}
              {activeTab === 'logs' && (<Logs />)}
              {activeTab === 'cj74/cji5' && (
                <RawDrill
                  user={user}
                  filters={sharedFilters} // Same filters jo Summary View mein hain
                  onFilterChange={handleFilterChange}
                  onResetFilters={handleResetFilters}
                />
              )}
              
              {['ftc'].includes(activeTab) && (
                <div className="bg-white p-20 rounded-xl shadow text-center border-2 border-dashed border-gray-200">
                  <h2 className="text-xl font-bold text-gray-700 capitalize">{activeTab} Page</h2>
                  <p className="text-gray-400 mt-2">This module is under development.</p>
                </div>
              )}
            </>
          } />

          <Route path="/drilldown" element={<DrillDownPage />} />
          <Route path="/request-access" element={<RequestAccess />} />
        <Route path="/admin/access-requests" element={<AccessRequestsTable />} />
        </Routes>
      </main>
    </div>
  );
}

export default App;