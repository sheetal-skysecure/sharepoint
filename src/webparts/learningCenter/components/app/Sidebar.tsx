import * as React from 'react';
import { NavLink } from 'react-router-dom';
import { Home, GraduationCap, Compass, ShieldCheck, ClipboardCheck, BookOpen } from 'lucide-react';
import logo from './skysecure_logo_clean_1772790765128.png';
import './Sidebar.css';

export default function Sidebar({ width, canAccessAdmin }: { width?: number, canAccessAdmin?: boolean }) {
    return (
        <aside className="app-sidebar" style={{ width: width || 280 }}>
            <div className="sidebar-brand">
                <img src={logo} alt="SkySecure Logo" />
                <span className="brand-text">skysecure</span>
            </div>

            <nav className="sidebar-nav">
                <div className="nav-group">
                    <span className="nav-label">General</span>
                    <NavLink to="/" className={({ isActive }) => `nav-item ${isActive ? 'active' : ''}`}>
                        <Home size={20} />
                        <span>Home</span>
                    </NavLink>
                </div>

                <div className="nav-group">
                    <span className="nav-label">Success</span>
                    <NavLink to="/learning-center" className={({ isActive }) => `nav-item ${isActive ? 'active' : ''}`}>
                        <GraduationCap size={20} />
                        <span>Certifications</span>
                    </NavLink>
                    <NavLink to="/assessments" className={({ isActive }) => `nav-item ${isActive ? 'active' : ''}`}>
                        <ClipboardCheck size={20} />
                        <span>Assessments</span>
                    </NavLink>
                    <NavLink to="/explore" className={({ isActive }) => `nav-item ${isActive ? 'active' : ''}`}>
                        <Compass size={20} />
                        <span>Self Explore</span>
                    </NavLink>
                    <NavLink to="/library" className={({ isActive }) => `nav-item ${isActive ? 'active' : ''}`}>
                        <BookOpen size={20} />
                        <span>Content Library</span>
                    </NavLink>
                </div>

                {canAccessAdmin && (
                    <div className="nav-group">
                        <span className="nav-label">Administrative</span>
                        <NavLink to="/admin" className={({ isActive }) => `nav-item ${isActive ? 'active' : ''}`}>
                            <ShieldCheck size={20} />
                            <span>Admin Portal</span>
                        </NavLink>
                    </div>
                )}
            </nav>

            <div className="sidebar-footer">
                <p style={{ fontSize: '0.7rem', color: '#94a3b8', fontWeight: 800, textAlign: 'center', letterSpacing: '0.05em' }}>&copy; 2026 SKYSECURE</p>
            </div>
        </aside>
    );
}
