import * as React from 'react';
import { Link, useLocation } from 'react-router-dom';
import { LayoutGrid, Bell, UserCircle, Search, ShieldCheck } from 'lucide-react';

export default function TopNav() {
    const location = useLocation();

    const isActive = (path: string) => location.pathname === path;

    return (
        <nav className="navbar" style={{
            backdropFilter: 'blur(20px)',
            backgroundColor: 'rgba(255, 255, 255, 0.85)',
            borderBottom: '1px solid rgba(0, 0, 0, 0.05)',
            height: '72px'
        }}>
            <div className="container flex items-center justify-between w-full">
                <Link to="/" className="nav-brand" style={{ fontSize: '1.25rem', fontWeight: 900, letterSpacing: '-0.03em' }}>
                    <div style={{
                        width: '32px',
                        height: '32px',
                        background: 'var(--gradient-primary)',
                        borderRadius: '8px',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        color: 'white'
                    }}>
                        <LayoutGrid size={18} />
                    </div>
                    SkySecure<span style={{ color: 'var(--text-muted)' }}>Tech</span>
                </Link>

                <div className="flex items-center gap-8">
                    <div className="flex items-center gap-2">
                        <NavLink to="/" active={isActive('/')} label="Home" />
                        <NavLink to="/learning-center" active={location.pathname.startsWith('/learning-center')} label="Certifications" />
                        <NavLink to="/admin" active={isActive('/admin')} label="Admin" icon={<ShieldCheck size={14} />} />
                    </div>

                    <div className="flex items-center gap-4" style={{
                        borderLeft: '1px solid #f1f5f9',
                        paddingLeft: '2rem',
                        marginLeft: '1rem'
                    }}>
                        <button style={{ color: '#94a3b8', transition: 'all 0.2s' }} className="hover:text-primary"><Search size={20} /></button>
                        <button style={{ color: '#94a3b8', transition: 'all 0.2s', position: 'relative' }} className="hover:text-primary">
                            <Bell size={20} />
                            <span style={{ position: 'absolute', top: '-2px', right: '-2px', width: '8px', height: '8px', backgroundColor: '#ef4444', borderRadius: '50%', border: '2px solid white' }}></span>
                        </button>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '4px 4px 4px 12px', backgroundColor: '#f8fafc', borderRadius: '100px', border: '1px solid #f1f5f9' }}>
                            <div style={{ fontSize: '0.75rem', fontWeight: 700, color: '#475569' }}>Sheetal S.</div>
                            <button style={{ color: 'var(--primary)', display: 'flex' }} className="hover:opacity-80">
                                <UserCircle size={28} />
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </nav>
    );
}

function NavLink({ to, active, label, icon }: { to: string, active: boolean, label: string, icon?: any }) {
    return (
        <Link
            to={to}
            style={{
                fontSize: '0.9rem',
                fontWeight: active ? 800 : 500,
                color: active ? 'var(--primary)' : '#64748b',
                padding: '0.5rem 1rem',
                borderRadius: '8px',
                backgroundColor: active ? 'var(--primary-light)' : 'transparent',
                transition: 'all 0.3s cubic-bezier(0.16, 1, 0.3, 1)',
                display: 'flex',
                alignItems: 'center',
                gap: '8px'
            }}
            className={active ? "" : "hover:bg-gray-50 hover:text-main"}
        >
            {icon}
            {label}
        </Link>
    );
}
