import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { HashRouter as Router, Routes, Route, useLocation, useNavigate } from 'react-router-dom';
import Sidebar from './app/Sidebar';
import Home from './app/Home';
import LearningCenterView from './app/LearningCenterView';
import CompanyProvider from './app/CompanyProvider';
import CertificationsList from './app/CertificationsList';
import SelfExplore from './app/SelfExplore';
import Assessments from './app/Assessments';
import ClientContentLibrary from './app/ClientContentLibrary';
import AdminPortal from '../../adminAccess/components/AdminPortal';
import { Bell, Settings, UserCircle, Maximize2, Minimize2, AlertTriangle, Globe, CheckCircle2, Shield, CheckCircle, Mail, Cloud } from 'lucide-react';
import type { ILearningCenterProps } from './ILearningCenterProps';
import { SharePointService } from '../services/SharePointService';
import LayoutScrollWrapper from '../../shared/components/LayoutScrollWrapper';
import './app/styles.css';
import './app/Sidebar.css';

function PortalContent({
  sidebarWidth, startResizing, handleResizing, stopResizing, isResizing, isExpanded, setIsExpanded,
  isOwner, canAccessAdmin, adminAccessResolved, userRole, userDisplayName, userEmail, showNotifications, setShowNotifications, showSettings,
  setShowSettings, showProfile, setShowProfile, notifications, onLogout, context, userGroup,
  onMarkRead, onNotificationClick
}: any) {
  const userPhotoUrl = userEmail ? `/_layouts/15/userphoto.aspx?size=M&accountname=${userEmail}` : null;
  const location = useLocation();
  const navigate = useNavigate();
  const isAdmin = location.pathname.includes('admin');
  const unreadCount = notifications.filter((n: any) => !n.read).length;
  const notificationsRef = useRef<HTMLDivElement | null>(null);
  const settingsRef = useRef<HTMLDivElement | null>(null);
  const profileRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      const target = event.target as Node;

      if (showNotifications && notificationsRef.current && !notificationsRef.current.contains(target)) {
        setShowNotifications(false);
      }

      if (showSettings && settingsRef.current && !settingsRef.current.contains(target)) {
        setShowSettings(false);
      }

      if (showProfile && profileRef.current && !profileRef.current.contains(target)) {
        setShowProfile(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [showNotifications, showSettings, showProfile, setShowNotifications, setShowSettings, setShowProfile]);

  // Security check: If trying to access admin but not an owner, redirect or show denied

  return (
    <div className={`sharepoint-learning-portal-root ${isResizing ? 'is-resizing' : ''} ${isExpanded || isAdmin ? 'is-expanded-full' : ''} ${isAdmin ? 'is-admin-view' : ''}`}>
      {/* Sidebar - Hide if in Admin mode for a broader view */}
      {!isAdmin && <Sidebar width={sidebarWidth} canAccessAdmin={adminAccessResolved && canAccessAdmin} />}
      {!isAdmin && (
        <div
          className="sidebar-resizer"
          onMouseDown={startResizing}
        />
      )}

      <main className="portal-main-container" style={{ width: '100%', backgroundColor: '#f8fafc' }}>
        {/* Header - Hide if in Admin mode to match the Admin aesthetic */}
        {!isAdmin && (
          <header className="portal-top-bar" style={{ display: 'flex', justifyContent: 'space-between', padding: '0 2rem' }}>
            <div className="header-actions-unified" style={{ gap: '1.5rem', display: 'flex', alignItems: 'center', marginLeft: 'auto' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '6px 12px', background: '#f0fdf4', color: '#16a34a', borderRadius: '100px', fontSize: '0.75rem', fontWeight: 800, border: '1px solid #dcfce7' }}>
                <Cloud size={14} /> ORGANIZATIONAL SYNC ACTIVE
              </div>
              <button
                className="icon-btn tooltip-trigger"
                onClick={() => setIsExpanded(!isExpanded)}
                title={isExpanded ? "Exit Fullscreen" : "Maximize View"}
                style={{ backgroundColor: isExpanded ? 'var(--primary-light)' : 'transparent', color: isExpanded ? 'var(--primary)' : 'inherit' }}
              >
                {isExpanded ? <Minimize2 size={20} /> : <Maximize2 size={20} />}
              </button>
              <div ref={notificationsRef} style={{ position: 'relative' }}>
                <button className={`icon-btn ${showNotifications ? 'is-active' : ''}`} onClick={() => { setShowNotifications(!showNotifications); setShowSettings(false); }}>
                  <Bell size={20} />
                  {unreadCount > 0 && (
                    <span className="notification-dot" style={{
                      minWidth: '18px',
                      height: '18px',
                      padding: '0 5px',
                      borderRadius: '999px',
                      display: 'inline-flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      fontSize: '0.65rem',
                      fontWeight: 800
                    }}>
                      {unreadCount}
                    </span>
                  )}
                </button>
                {showNotifications && (
                    <div className="dropdown-panel notifications-dropdown fade-in">
                      <div className="dropdown-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <span>Assignments</span>
                        <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                          <button 
                            className="btn-view-latest" 
                            onClick={() => { setShowNotifications(false); navigate('/learning-center'); }}
                            style={{ padding: '4px 10px', fontSize: '0.7rem', background: 'var(--primary)', color: 'white', border: 'none', borderRadius: '100px', cursor: 'pointer', fontWeight: 700 }}
                          >
                            OPEN PATHS
                          </button>
                          <span title="Mark all as read" style={{ display: 'flex' }} onClick={() => onMarkRead && onMarkRead()}><CheckCircle size={18} style={{ color: 'var(--primary)', cursor: 'pointer' }} /></span>
                        </div>
                      </div>
                    <div className="dropdown-content">
                      {notifications.length === 0 ? (
                        <div style={{ padding: '2rem', textAlign: 'center', color: '#94a3b8', fontWeight: 700 }}>
                          No new certifications assigned
                        </div>
                      ) : notifications.map((n: any) => (
                        <div
                          key={n.id}
                          className={`notification-item ${!n.read ? 'is-unread' : ''}`}
                          onClick={() => onNotificationClick && onNotificationClick(n, navigate)}
                          style={{ cursor: 'pointer' }}
                        >
                          <div className="n-icon-box">
                            <CheckCircle size={20} />
                          </div>
                          <div className="n-content">
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                              <div className="n-title">{n.title}</div>
                              <button 
                                className="btn-view-notif" 
                                onClick={(event) => {
                                  event.stopPropagation();
                                  if (onNotificationClick) {
                                    onNotificationClick(n, navigate);
                                  }
                                }}
                                style={{ padding: '2px 8px', fontSize: '0.65rem', background: 'var(--primary-light)', color: 'var(--primary)', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 700 }}
                              >
                                VIEW
                              </button>
                            </div>
                            <div className="n-text">{n.text}</div>
                            <div className="n-time">Assigned: {n.assignedDate || n.time}</div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              <div ref={settingsRef} style={{ position: 'relative' }}>
                <button className={`icon-btn ${showSettings ? 'is-active' : ''}`} onClick={() => { setShowSettings(!showSettings); setShowNotifications(false); }}>
                  <Settings size={18} />
                </button>
                {showSettings && (
                  <div className="dropdown-panel settings-dropdown fade-in">
                    <div className="dropdown-header">System Settings</div>
                    <div className="dropdown-content">
                      <div className="settings-item">
                        <div className="settings-label"><Maximize2 size={18} /> Display Mode</div>
                        <div className="settings-value">{isExpanded ? 'Full Screen' : 'Standard'}</div>
                      </div>
                      <div className="settings-item">
                        <div className="settings-label"><Globe size={18} /> Language</div>
                        <div className="settings-value">English</div>
                      </div>
                      <div className="settings-item">
                        <div className="settings-label"><Shield size={18} /> Security Level</div>
                        <div className="settings-value">{userRole}</div>
                      </div>
                      <div className="settings-item">
                        <div className="settings-label"><Mail size={18} /> Email Support</div>
                      </div>
                      {/* Sign Out removed as per request */}
                    </div>
                  </div>
                )}
              </div>

              <div ref={profileRef} style={{ position: 'relative' }}>
                <div className={`user-pill-unified ${showProfile ? 'is-active' : ''}`} onClick={() => { setShowProfile(!showProfile); setShowNotifications(false); setShowSettings(false); }}>
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', marginRight: '4px' }}>
                    <span style={{ fontSize: '0.85rem', fontWeight: 950, color: '#1e293b', lineHeight: 1 }}>{userDisplayName || 'Profile'}</span>
                    <span style={{ fontSize: '0.65rem', fontWeight: 800, color: '#64748b', textTransform: 'uppercase' }}>{userGroup}</span>
                  </div>
                  {userPhotoUrl ? (
                    <img src={userPhotoUrl} alt="Avatar" style={{ width: '32px', height: '32px', borderRadius: '10px', objectFit: 'cover' }} />
                  ) : (
                    <div style={{ width: '32px', height: '32px', background: 'var(--primary)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'white', borderRadius: '10px' }}>
                      <UserCircle size={24} />
                    </div>
                  )}
                </div>

                {/* Standard div instead of motion.div for debugging */}
                {showProfile && (
                  <>
                    <div
                      className="profile-backdrop"
                      onClick={() => setShowProfile(false)}
                    />
                    <div
                      className="profile-popover-anchor"
                    >
                      <div className="user-account-card">
                        <div className="account-card-header">
                          <h3>User Account</h3>
                        </div>

                        <div className="account-card-profile">
                          <div className="account-card-avatar">
                            {userPhotoUrl ? (
                              <img src={userPhotoUrl} alt="Avatar" />
                            ) : (
                              <div style={{ width: '100%', height: '100%', background: 'var(--primary)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'white', borderRadius: 'inherit' }}>
                                <UserCircle size={32} />
                              </div>
                            )}
                          </div>
                          <div className="account-card-details">
                            <div className="name" style={{ fontSize: '1.2rem' }}>{userDisplayName}</div>
                            <div className="email" style={{ fontSize: '0.9rem' }}>{userEmail}</div>
                          </div>
                        </div>

                        <div className="account-card-stats">
                          <div className="account-card-stat-item">
                            <div className="account-label">
                              <CheckCircle2 size={18} style={{ color: '#10b981' }} />
                              <span>Status</span>
                            </div>
                            <div className="account-value status">Active</div>
                          </div>
                          <div className="account-card-stat-item">
                            <div className="account-label">
                              <Shield size={18} style={{ color: 'var(--primary)' }} />
                              <span>Role Access</span>
                            </div>
                            <div className="account-value">{userRole}</div>
                          </div>
                        </div>

                        <button
                          className="account-signout-btn"
                          onClick={() => { setShowProfile(false); setShowSettings(true); }}
                        >
                          <Settings size={20} />
                          Account Settings
                        </button>

                        {/* Sign Out removed as per request */}
                      </div>
                    </div>
                  </>
                )}
              </div>
            </div>
          </header>
        )}

        <div className="portal-scroll-area" style={{
          height: (isAdmin || isExpanded) ? (isAdmin ? '100vh' : 'calc(100vh - 80px)') : 'calc(100vh - 80px)',
          overflowY: 'auto'
        }}>
          <LayoutScrollWrapper className="portal-layout-scroll-frame" innerClassName="portal-layout-scroll-frame__inner">
            <Routes>
              <Route path="/" element={<Home />} />
              <Route path="/user-dashboard" element={<Home />} />
              <Route path="/learning-center" element={<LearningCenterView />} />
              <Route path="/learning-center/:companyId" element={<CompanyProvider userEmail={userEmail} userDisplayName={userDisplayName} />}>
                <Route index element={<CertificationsList />} />
              </Route>
              <Route path="/assessments" element={<Assessments userDisplayName={userDisplayName} userEmail={userEmail} context={context} />} />
              <Route path="/explore" element={<SelfExplore />} />
              <Route path="/library" element={<ClientContentLibrary />} />
              <Route path="/admin" element={!adminAccessResolved ? <div className="access-denied">
                <div className="denied-card">
                  <AlertTriangle size={48} color="#0ea5e9" />
                  <h2>Checking Access</h2>
                  <p>Validating SharePoint Owners and Members group access for the Admin Portal.</p>
                </div>
              </div> : canAccessAdmin ? <AdminPortal userDisplayName={userDisplayName} userEmail={userEmail} isOwner={userRole === 'Owner'} canAccessAdmin={canAccessAdmin} userRole={userRole} context={context} /> : <div className="access-denied">
                <div className="denied-card">
                  <AlertTriangle size={48} color="#f43f5e" />
                  <h2>Access Restricted</h2>
                  <p>You do not have permission to access the Admin Portal. Only SharePoint Owners and Members are allowed.</p>
                  <button className="btn-primary" onClick={() => navigate('/')}>Back to Safety</button>
                </div>
              </div>} />
              {/* Catch-all to home */}
              <Route path="*" element={<Home />} />
            </Routes>
          </LayoutScrollWrapper>
        </div>
      </main>
    </div>
  );
}

export default function LearningCenter(props: ILearningCenterProps) {
  const [sidebarWidth, setSidebarWidth] = useState(240);
  const [isResizing, setIsResizing] = useState(false);
  const [isExpanded, setIsExpanded] = useState(true);
  const [showNotifications, setShowNotifications] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [showProfile, setShowProfile] = useState(false);

  // Use props directly instead of a mock activeUser
  const determinedEmail = props.userEmail || '';
  const determinedName = props.userDisplayName || 'User';
  
  const [userGroup, setUserGroup] = useState('Checking access...');
  const [userRole, setUserRole] = useState('Checking access...');
  const [canAccessAdmin, setCanAccessAdmin] = useState(false);
  const [adminAccessResolved, setAdminAccessResolved] = useState(false);

  const [isOwner, setIsOwner] = useState(false);

    useEffect(() => {
        const fetchRealStatus = async () => {
            setAdminAccessResolved(false);

            try {
                SharePointService.init(props.context.pageContext.web.absoluteUrl, props.context.spHttpClient, props.context);
                const accessState = await SharePointService.getCurrentUserAdminAccess(determinedEmail, true);
                const resolvedRole = accessState.currentUserRole === 'Unknown' ? 'Visitor' : accessState.currentUserRole;
                const resolvedGroup = resolvedRole === 'Learner' ? 'Standard Learner' : resolvedRole;

                setUserRole(resolvedRole);
                setUserGroup(accessState.accessCheckFailed ? 'Access Check Failed' : resolvedGroup);
                setIsOwner(resolvedRole === 'Owner');
                setCanAccessAdmin(accessState.canAccessAdmin);
                setAdminAccessResolved(true);

                console.log('Real-time Group Status Fetched:', {
                  role: resolvedRole,
                  canAccessAdmin: accessState.canAccessAdmin
                });
            } catch (error) {
                console.error('Failed to fetch group status', error);
                setUserRole('Visitor');
                setUserGroup('Access Check Failed');
                setIsOwner(false);
                setCanAccessAdmin(false);
                setAdminAccessResolved(true);
            }
        };
        fetchRealStatus();
    }, [determinedEmail, props.context]);

  const [userNotifications, setUserNotifications] = useState<any[]>([]);

  useEffect(() => {
    const loadNotifs = async () => {
      if (!determinedEmail) return;
      
      try {
        const spNotifs = await SharePointService.getUserAssignmentNotifications(determinedEmail.toLowerCase());
        setUserNotifications(spNotifs || []);
      } catch (e) {
          console.warn("Could not fetch assignment notifications from SharePoint", e);
          setUserNotifications([]);
      }
    };
    void loadNotifs();
  }, [determinedEmail]);

  const handleMarkRead = async () => {
    try {
      const unread = userNotifications.filter(n => !n.read);
      for (const n of unread) {
        if (n.id) {
          await SharePointService.markAssignmentNotificationAsRead(n.id, n.sourceList);
        }
      }
      setUserNotifications(prev => prev.map(n => ({ ...n, read: true, status: 'Viewed' })));
    } catch (e) {
      console.error("Failed to mark notifications as read", e);
    }
  };

  const getNotificationTargetPath = (notification: any) => {
    const statusText = `${notification?.title || ''} ${notification?.text || ''}`.toLowerCase();
    return statusText.indexOf('assessment') !== -1 ? '/assessments' : '/learning-center';
  };

  const handleNotificationClick = async (notification: any, navigateFn?: (path: string) => void) => {
    try {
      if (!notification.read && notification.id) {
        await SharePointService.markAssignmentNotificationAsRead(notification.id, notification.sourceList);
        setUserNotifications(prev => prev.map(item =>
          item.id === notification.id && item.sourceList === notification.sourceList
            ? { ...item, read: true, status: 'Viewed' }
            : item
        ));
      }
    } catch (error) {
      console.error('Failed to mark assignment notification as read', error);
    } finally {
      setShowNotifications(false);
      const targetPath = getNotificationTargetPath(notification);
      if (navigateFn) {
        navigateFn(targetPath);
      } else {
        window.location.assign(props.context.pageContext.web.absoluteUrl);
      }
    }
  };

  const startResizing = (e: any) => {
    setIsResizing(true);
    e.preventDefault();
  };

  const stopResizing = () => {
    setIsResizing(false);
  };

  const handleResizing = (e: any) => {
    if (isResizing) {
      const newWidth = e.clientX;
      if (newWidth > 200 && newWidth < 500) {
        setSidebarWidth(newWidth);
      }
    }
  };

  useEffect(() => {
    if (isResizing) {
      window.addEventListener('mousemove', handleResizing);
      window.addEventListener('mouseup', stopResizing);
    } else {
      window.removeEventListener('mousemove', handleResizing);
      window.removeEventListener('mouseup', stopResizing);
    }
    return () => {
      window.removeEventListener('mousemove', handleResizing);
      window.removeEventListener('mouseup', stopResizing);
    };
  }, [isResizing]);

  const handleLogout = () => {
    // In SPFx, logout is handled by the browser/SharePoint, but we can redirect or clear local state if needed
    window.location.href = props.context.pageContext.web.absoluteUrl + "/_layouts/15/signout.aspx";
  };

  return (
    <Router>
      <PortalContent
        sidebarWidth={sidebarWidth}
        startResizing={startResizing}
        handleResizing={handleResizing}
        stopResizing={stopResizing}
        isResizing={isResizing}
        isExpanded={isExpanded}
        setIsExpanded={setIsExpanded}
        isOwner={isOwner}
        canAccessAdmin={canAccessAdmin}
        adminAccessResolved={adminAccessResolved}
        userRole={userRole}
        userDisplayName={determinedName}
        userEmail={determinedEmail}
        showNotifications={showNotifications}
        setShowNotifications={setShowNotifications}
        showSettings={showSettings}
        setShowSettings={setShowSettings}
        showProfile={showProfile}
        setShowProfile={setShowProfile}
        notifications={userNotifications}
        onLogout={handleLogout}
        context={props.context}
        userGroup={userGroup}
        onMarkRead={handleMarkRead}
        onNotificationClick={handleNotificationClick}
      />
    </Router>
  );
}
