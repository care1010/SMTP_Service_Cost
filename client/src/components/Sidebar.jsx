import React, { useState } from 'react';
import './Sidebar.css';
import logo from '../assets/OIP.jpg';
import MyAccess from '../pages/MyAccess';
import Swal from 'sweetalert2';

// 🔥 FIX: Yahan brackets ke andar 'user' aur 'onLogout' add kiya
const Sidebar = ({ activeTab, setActiveTab, user, onLogout }) => {
const [openMenu, setOpenMenu] = useState(false);

const handleLogout = () => {
  Swal.fire({
    title: 'Sign Out?',
    text: 'Are you sure you want to sign out?',
    icon: 'warning',
    showCancelButton: true,
    confirmButtonText: 'Yes, Sign Out',
    cancelButtonText: 'Cancel',
    background: '#0f172a',
    color: '#fff',
    confirmButtonColor: '#dc2626',
    cancelButtonColor: '#475569'
  }).then((result) => {
    if (result.isConfirmed) {
      Swal.fire({
        icon: 'success',
        title: 'Signed Out Successfully',
        timer: 1000,
        showConfirmButton: false,
        background: '#0f172a',
        color: '#fff'
      }).then(() => {
        onLogout();
      });
    }
  });
};

  const menuItems = [
    { id: 'dashboard', label: 'Dashboard' },
    { id: 'summary', label: 'Summary View' },
    { id: 'erp_resource', label: 'Cross ERP'},
    { id: 'add-project', label: 'Add Project / WBS' },
    { id: 'asbl', label: 'ASBL' },
    { id: 'cj74/cji5', label: 'CJ74/CJi5' },
    // { id: 'loa-view', label: 'Loa Wise View' },
    // Inhe sirf admin ya super_admin ko dikhao
    ...(user?.type !== 'user' ? [
        { id: 'ptd', label: 'PTD' },
        { id: 'logs', label: 'Logs' }
    ] : [])
];

  return (
    <div className="sidebar-container">
      
      {/* LOGO */}
      <div className="logo-section">
        <img src={logo} alt="Logo" className="logo-image" />
      </div>
      
      {/* MENU */}
      <nav className="nav-menu">
        {menuItems.map((item) => (
          <button
            key={item.id}
            onClick={() => setActiveTab(item.id)}
            className={`nav-btn ${activeTab === item.id ? 'active' : ''}`}
          >
            {item.label}
          </button>
        ))}
      </nav>
      
      {/* PROFILE SECTION */}
      <div className="profile-section">
        <div 
          className="profile-icon"
          onClick={() => setOpenMenu(!openMenu)}
        >
          <span className="user-icon">👤</span>
        </div>

        {openMenu && (
          <div className="profile-card">
            
            <div className="profile-header">
              {/* Dynamic Avatar (Email ka pehla letter) */}
              <div className="avatar">{user?.email?.charAt(0).toUpperCase() || 'U'}</div>
              <div>
                {/* Dynamic Email aur Role */}
                <div className="email">{user?.email || 'User'}</div>
                <div className="role">{user?.type?.toUpperCase() || 'GUEST'}</div>
              </div>
            </div>

            <div className="divider"></div>

            {/* ROLE CHECK: Sirf Admin aur Super Admin ko dikhega */}
            {(user?.type === 'admin' || user?.type === 'super_admin') && (
                <div 
                    className="menu-item" 
                    onClick={() => { setActiveTab('admin'); setOpenMenu(false); }}
                    style={{ cursor: 'pointer' }}
                >
                    🛠 Admin Panel
                </div>
            )}

          {/* My Access Tab everybody can see */}
          <div
              className="menu-item"
              onClick={() => {
                  setActiveTab('my-access');
                  setOpenMenu(false);
              }}
              style={{ cursor: 'pointer' }}
          >
              🔐 My Access
          </div>

            {/*  SIGN OUT BUTTON (onLogout function ab defined hai) */}
            <div className="menu-item" onClick={handleLogout} style={{cursor: 'pointer'}}>
              🚪 Sign Out
            </div>

            <div className="cancel-btn" onClick={() => setOpenMenu(false)}>
              Cancel
            </div>

          </div>
        )}
      </div>
    </div>
  );
};

export default Sidebar;