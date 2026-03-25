import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import { Search, Eye, Download, FileText, Video, Link2, FileArchive, Folder, Activity, BookOpen, Clock, FileSpreadsheet, FileCode, Presentation } from 'lucide-react';
import { LMS_CONTENT_LIBRARY_REFRESH_EVENT, SharePointService } from '../../services/SharePointService';

type FolderGroups = Record<string, any[]>;

export default function ClientContentLibrary() {
    const [assets, setAssets] = useState<any[]>([]);
    const [searchTerm, setSearchTerm] = useState('');
    const [filterType, setFilterType] = useState('ALL');
    const [selectedFolder, setSelectedFolder] = useState('ALL');
    const [isLoading, setIsLoading] = useState(true);

    useEffect(() => {
        let isMounted = true;

        const fetchUserContent = async (showLoadingState: boolean = false) => {
            if (showLoadingState && isMounted) {
                setIsLoading(true);
            }

            try {
                const spAssets = await SharePointService.getContentAssets();
                if (isMounted) {
                    setAssets(spAssets || []);
                }
            } catch (error) {
                console.error('Error loading SharePoint assets:', error);
                if (isMounted) {
                    setAssets([]);
                }
            } finally {
                if (showLoadingState && isMounted) {
                    setIsLoading(false);
                }
            }
        };

        void fetchUserContent(true);
        const intervalId = window.setInterval(() => {
            void fetchUserContent();
        }, 5000);
        const handleContentRefresh = () => {
            void fetchUserContent();
        };
        window.addEventListener(LMS_CONTENT_LIBRARY_REFRESH_EVENT, handleContentRefresh);

        return () => {
            isMounted = false;
            window.clearInterval(intervalId);
            window.removeEventListener(LMS_CONTENT_LIBRARY_REFRESH_EVENT, handleContentRefresh);
        };
    }, []);

    const getIcon = (type: string) => {
        switch (type) {
            case 'VIDEO': return <Video size={20} />;
            case 'PDF': return <FileText size={20} />;
            case 'EXCEL': return <FileSpreadsheet size={20} />;
            case 'DOC': return <FileCode size={20} />;
            case 'PPT': return <Presentation size={20} />;
            case 'LINK': return <Link2 size={20} />;
            case 'SCORM': return <FileArchive size={20} />;
            default: return <Folder size={20} />;
        }
    };

    const normalizedFolders = useMemo(() => {
        return Array.from(new Set(
            assets
                .map((asset) => (asset.folderName || 'Others').toString().trim())
                .filter((folder) => !!folder)
        )).sort((left, right) => left.localeCompare(right));
    }, [assets]);

    const filteredAssets = useMemo(() => {
        return assets.filter((asset) => {
            const folderName = (asset.folderName || 'Others').toString().trim();
            const normalizedSearch = searchTerm.toLowerCase();
            const matchesSearch =
                (asset.name || '').toLowerCase().includes(normalizedSearch) ||
                (asset.description || '').toLowerCase().includes(normalizedSearch) ||
                folderName.toLowerCase().includes(normalizedSearch);
            const matchesType = filterType === 'ALL' || asset.type === filterType;
            const matchesFolder = selectedFolder === 'ALL' || folderName.toLowerCase() === selectedFolder.toLowerCase();

            return matchesSearch && matchesType && matchesFolder;
        });
    }, [assets, filterType, searchTerm, selectedFolder]);

    const groupedAssets = useMemo(() => {
        return filteredAssets.reduce((acc: FolderGroups, asset) => {
            const folder = (asset.folderName || 'Others').toString().trim() || 'Others';
            if (!acc[folder]) {
                acc[folder] = [];
            }
            acc[folder].push(asset);
            return acc;
        }, {});
    }, [filteredAssets]);

    const groupedEntries = useMemo(() => {
        return Object.keys(groupedAssets)
            .sort((left: string, right: string) => left.localeCompare(right))
            .map((folderName: string) => [folderName, groupedAssets[folderName]] as [string, any[]]);
    }, [groupedAssets]);

    const buildAssetActionUrl = (candidate: string, openInSharePointViewer: boolean): string => {
        const normalizedCandidate = (candidate || '').toString().trim();
        if (!normalizedCandidate) {
            return '';
        }

        try {
            const resolvedUrl = new URL(normalizedCandidate, window.location.origin);
            if (openInSharePointViewer) {
                resolvedUrl.searchParams.set('web', '1');
            } else {
                resolvedUrl.searchParams.delete('web');
            }

            return resolvedUrl.toString();
        } catch (error) {
            if (!openInSharePointViewer) {
                return normalizedCandidate;
            }

            return normalizedCandidate.indexOf('?') >= 0
                ? `${normalizedCandidate}&web=1`
                : `${normalizedCandidate}?web=1`;
        }
    };

    const handleOpenAsset = (asset: any, openInSharePointViewer: boolean = false) => {
        const finalUrl = (asset.url || asset.path || '').toString().trim();

        if (!finalUrl) {
            alert('This document cannot be opened because no valid SharePoint file URL was found.');
            return;
        }

        const targetUrl = buildAssetActionUrl(finalUrl, openInSharePointViewer);
        if (!targetUrl) {
            alert('This document cannot be opened because no valid SharePoint file URL was found.');
            return;
        }

        window.open(targetUrl, '_blank', 'noopener,noreferrer');
    };

    return (
        <div className="container animate-fade-in" style={{ padding: '3rem 2rem' }}>
            <header style={{ marginBottom: '3rem' }}>
                <h1 style={{ fontSize: '2.5rem', fontWeight: 800, marginBottom: '0.75rem', letterSpacing: '-0.02em', color: '#1e293b' }}>Content Library</h1>
                <p style={{ color: 'var(--text-muted)', fontSize: '1.15rem' }}>Documents are loaded from SharePoint `Documents1` and grouped by their assigned folder.</p>
            </header>

            <div style={{ display: 'flex', gap: '1rem', flexWrap: 'wrap', marginBottom: '1rem' }}>
                <div style={{ flex: 1, minWidth: '300px', position: 'relative' }}>
                    <Search style={{ position: 'absolute', left: '1rem', top: '50%', transform: 'translateY(-50%)', color: '#94a3b8' }} size={20} />
                    <input
                        type="text"
                        placeholder="Search materials..."
                        value={searchTerm}
                        onChange={e => setSearchTerm(e.target.value)}
                        style={{
                            width: '100%',
                            padding: '1rem 1rem 1rem 3rem',
                            borderRadius: '16px',
                            border: '1.5px solid #e2e8f0',
                            fontSize: '1rem',
                            outline: 'none',
                            fontWeight: 600
                        }}
                    />
                </div>

                <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'center', background: 'white', padding: '0.5rem', borderRadius: '16px', border: '1.5px solid #e2e8f0', flexWrap: 'wrap' }}>
                    {['ALL', 'VIDEO', 'PDF', 'EXCEL', 'DOC', 'PPT', 'SCORM', 'LINK'].map(t => (
                        <button
                            key={t}
                            onClick={() => setFilterType(t)}
                            style={{
                                padding: '0.5rem 1rem',
                                borderRadius: '12px',
                                fontSize: '0.85rem',
                                fontWeight: 800,
                                background: filterType === t ? 'var(--primary)' : 'transparent',
                                color: filterType === t ? 'white' : '#64748b',
                                border: 'none',
                                cursor: 'pointer',
                                transition: 'all 0.2s'
                            }}
                        >
                            {t}
                        </button>
                    ))}
                </div>
            </div>

            <div style={{ display: 'flex', gap: '0.75rem', flexWrap: 'wrap', marginBottom: '2rem' }}>
                <button
                    onClick={() => setSelectedFolder('ALL')}
                    style={{
                        padding: '0.65rem 1rem',
                        borderRadius: '999px',
                        border: '1.5px solid',
                        borderColor: selectedFolder === 'ALL' ? 'var(--primary)' : '#dbe4f0',
                        background: selectedFolder === 'ALL' ? 'rgba(37, 99, 235, 0.08)' : 'white',
                        color: selectedFolder === 'ALL' ? 'var(--primary)' : '#64748b',
                        fontWeight: 800,
                        cursor: 'pointer'
                    }}
                >
                    All
                </button>
                {normalizedFolders.map((folder) => (
                    <button
                        key={folder}
                        onClick={() => setSelectedFolder(folder)}
                        style={{
                            padding: '0.65rem 1rem',
                            borderRadius: '999px',
                            border: '1.5px solid',
                            borderColor: selectedFolder === folder ? 'var(--primary)' : '#dbe4f0',
                            background: selectedFolder === folder ? 'rgba(37, 99, 235, 0.08)' : 'white',
                            color: selectedFolder === folder ? 'var(--primary)' : '#64748b',
                            fontWeight: 800,
                            textTransform: 'capitalize',
                            cursor: 'pointer'
                        }}
                    >
                        {folder}
                    </button>
                ))}
            </div>

            {isLoading ? (
                <div style={{ textAlign: 'center', padding: '5rem 0', background: '#f8fafc', borderRadius: '24px', border: '1.5px solid #e2e8f0' }}>
                    <Activity size={48} className="animate-spin" style={{ color: 'var(--primary)', marginBottom: '1rem' }} />
                    <h3 style={{ fontSize: '1.5rem', fontWeight: 800, color: '#1e293b', margin: '0 0 0.5rem' }}>Synchronizing Library...</h3>
                    <p style={{ color: '#64748b', margin: 0, fontWeight: 600 }}>Fetching latest files from SharePoint</p>
                </div>
            ) : groupedEntries.length === 0 ? (
                <div style={{ textAlign: 'center', padding: '5rem 0', background: '#f8fafc', borderRadius: '24px', border: '2px dashed #e2e8f0' }}>
                    <BookOpen size={48} color="#cbd5e1" style={{ marginBottom: '1rem' }} />
                    <h3 style={{ fontSize: '1.5rem', fontWeight: 800, color: '#64748b', margin: '0 0 0.5rem' }}>No Materials Found</h3>
                    <p style={{ color: '#94a3b8', margin: 0, fontWeight: 600 }}>Try adjusting your search criteria, file type, or folder filter.</p>
                </div>
            ) : (
                <div style={{ display: 'grid', gap: '2rem' }}>
                    {groupedEntries.map(([folderName, folderAssets]: [string, any[]]) => (
                        <section key={folderName} style={{ display: 'grid', gap: '1rem' }}>
                            <div style={{
                                display: 'flex',
                                alignItems: 'center',
                                gap: '0.75rem',
                                padding: '1rem 1.25rem',
                                background: 'white',
                                borderRadius: '20px',
                                border: '1.5px solid #e2e8f0'
                            }}>
                                <Folder size={22} color="#2563eb" />
                                <div style={{ display: 'flex', alignItems: 'baseline', gap: '0.75rem', flexWrap: 'wrap' }}>
                                    <h2 style={{ margin: 0, fontSize: '1.2rem', fontWeight: 900, color: '#1e293b', textTransform: 'capitalize' }}>{folderName}</h2>
                                    <span style={{ fontSize: '0.85rem', color: '#64748b', fontWeight: 700 }}>{folderAssets.length} file{folderAssets.length === 1 ? '' : 's'}</span>
                                </div>
                            </div>

                            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))', gap: '1.5rem' }}>
                                {folderAssets.map((asset: any) => (
                                    <div key={`${folderName}-${asset.id || asset.url || asset.name}`} style={{
                                        background: 'white',
                                        padding: '1.5rem',
                                        borderRadius: '20px',
                                        border: '1.5px solid #e2e8f0',
                                        boxShadow: '0 4px 15px rgba(0,0,0,0.03)',
                                        display: 'flex',
                                        flexDirection: 'column',
                                        gap: '1rem'
                                    }}>
                                        <div style={{ display: 'flex', alignItems: 'flex-start', gap: '1rem' }}>
                                            <div style={{ width: '48px', height: '48px', background: '#f8fafc', borderRadius: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--primary)' }}>
                                                {getIcon(asset.type)}
                                            </div>
                                            <div style={{ flex: 1 }}>
                                                <h3 style={{ fontSize: '1.1rem', fontWeight: 800, color: '#1e293b', marginBottom: '0.25rem', lineHeight: '1.3' }}>{asset.name}</h3>
                                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: '#64748b', fontSize: '0.8rem', fontWeight: 600, flexWrap: 'wrap' }}>
                                                    <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>{asset.type}</span>
                                                    {asset.size ? <span>{asset.size}</span> : null}
                                                    <span style={{ textTransform: 'capitalize' }}>{folderName}</span>
                                                </div>
                                            </div>
                                        </div>

                                        <p style={{ color: '#475569', fontSize: '0.9rem', lineHeight: '1.5', margin: 0, flex: 1 }}>
                                            {asset.description || 'No description available for this resource.'}
                                        </p>

                                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', paddingTop: '1rem', borderTop: '1px solid #f1f5f9' }}>
                                            <div style={{ display: 'flex', alignItems: 'center', gap: '6px', color: '#94a3b8', fontSize: '0.8rem', fontWeight: 600 }}>
                                                <Clock size={14} /> {asset.dateAdded}
                                            </div>
                                            <div style={{ display: 'flex', gap: '0.5rem' }}>
                                                <button
                                                    onClick={() => handleOpenAsset(asset, true)}
                                                    style={{ background: '#f1f5f9', color: 'var(--primary)', border: 'none', padding: '0.5rem 1rem', borderRadius: '10px', display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 700, cursor: 'pointer' }}
                                                >
                                                    <Eye size={16} /> View
                                                </button>
                                                <button
                                                    onClick={() => handleOpenAsset(asset, false)}
                                                    style={{ background: 'var(--primary)', color: 'white', border: 'none', padding: '0.5rem 1rem', borderRadius: '10px', display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 700, cursor: 'pointer' }}
                                                >
                                                    <Download size={16} /> Open
                                                </button>
                                            </div>
                                        </div>
                                    </div>
                                ))}
                            </div>
                        </section>
                    ))}
                </div>
            )}
        </div>
    );
}
