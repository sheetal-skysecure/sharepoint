import * as React from 'react';
import { Outlet, useParams, Link, useLocation, useNavigate } from 'react-router-dom';
import { ChevronRight, ArrowLeft } from 'lucide-react';
import { companies } from './data';

export default function CompanyProvider(props: { userEmail?: string; userDisplayName?: string }) {
    const location = useLocation();
    const navigate = useNavigate();
    const { companyId } = useParams();
    const normalizedCompanyId = (
        (location.state as { provider?: string } | null)?.provider ||
        companyId ||
        'microsoft'
    ).toString().trim().toLowerCase();
    const company = companies.find((c: any) => (c.id || '').toString().trim().toLowerCase() === normalizedCompanyId) || companies[0];

    React.useEffect(() => {
        if ((companyId || '').toString().trim().toLowerCase() !== normalizedCompanyId) {
            navigate(`/learning-center/${normalizedCompanyId}`, {
                replace: true,
                state: { provider: normalizedCompanyId }
            });
        }
    }, [companyId, navigate, normalizedCompanyId]);

    if (!company) {
        return <div className="container" style={{ padding: '3rem' }}>Provider not found.</div>;
    }

    return (
        <div className="container animate-fade-in" style={{ padding: '2rem' }}>
            <nav style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '2rem', fontSize: '0.875rem', color: 'var(--text-muted)' }}>
                <Link to="/learning-center" style={{ display: 'flex', alignItems: 'center', gap: '0.25rem' }} className="hover:text-main">
                    <ArrowLeft size={16} />
                    Learning Center
                </Link>
                <ChevronRight size={14} />
                <span style={{ color: 'var(--text-main)', fontWeight: 500 }}>{company.name} Certifications</span>
            </nav>

            <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginBottom: '3rem' }}>
                <div style={{ backgroundColor: 'white', padding: '1rem', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                    <img src={company.logoUrl} alt={company.name} style={{ height: '40px', objectFit: 'contain' }} />
                </div>
                <div>
                    <h1 style={{ fontSize: '2rem', fontWeight: 600 }}>{company.name} Certification Paths</h1>
                    <p style={{ color: 'var(--text-muted)' }}>Browse available certifications and schedule your study period.</p>
                </div>
            </div>

            <main>
                <Outlet context={{ companyId: normalizedCompanyId, userEmail: props.userEmail || '', userDisplayName: props.userDisplayName || '' }} />
            </main>
        </div>
    );
}
