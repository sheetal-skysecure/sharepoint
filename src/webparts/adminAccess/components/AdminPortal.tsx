import * as React from 'react';
import { useState, useEffect, useMemo, useRef } from 'react';
import { useInRouterContext, useNavigate } from 'react-router-dom';
import * as XLSX from 'xlsx';
import {
    LayoutGrid,
    ShieldCheck,
    BookOpen,
    Users,
    AlertTriangle,
    Search,
    TrendingUp,
    X,
    Award,
    Tag,
    PlusCircle,
    CheckCircle,
    Settings,
    Trash2,
    Bell,
    UserCircle,
    LogOut,
    Info,
    CheckCircle2,
    Loader2,
    Shield,
    Activity,
    BarChart3,
    ClipboardCheck,
    FileQuestion,
    Upload,
    Filter,
    Edit,
    Plus,
    Briefcase,
    MapPin,
    Hash,
    Video,
    FileText,
    Link2,
    FileArchive,
    Folder,
    Clock,
    Eye,
    Download,
    Sparkles,
    Cloud,
    Globe,
    FileSpreadsheet,
    FileCode,
    Presentation,
    Send
} from 'lucide-react';

import { IAssessmentAssignmentRecord, IAssessmentTrackerItem, ICertificationMaxSeatsItem, IDepartmentProgressLearner, IDepartmentProgressSummary, IUpcomingRenewalRecord, LMS_CONTENT_LIBRARY_REFRESH_EVENT, LMS_ENROLLMENTS_REFRESH_EVENT, SharePointService } from '../../learningCenter/services/SharePointService';
import LayoutScrollWrapper from '../../shared/components/LayoutScrollWrapper';
import logo from './skysecure_logo_clean_1772790765128.png';
import './AdminPortal.css';

const LMS_AUDIT_REFRESH_EVENT = 'lms-audit-refresh';
const dispatchEnrollmentRefreshSignal = (): void => {
    window.dispatchEvent(new Event(LMS_ENROLLMENTS_REFRESH_EVENT));
};

const getEnrollmentAssignmentErrorMessage = (error: unknown): string => {
    const rawMessage = error instanceof Error ? error.message : '';
    if (!rawMessage) {
        return 'Error assigning certification';
    }

    if (rawMessage.toLowerCase().indexOf('already enrolled') !== -1) {
        return 'Already enrolled';
    }

    if (rawMessage.toLowerCase().indexOf('certification not found') !== -1) {
        return 'Certification not found';
    }

    return rawMessage;
};

const resolveCertificationAssignmentDetails = async (cert: any): Promise<{ certificationId: number; certName: string; certCode: string; pathId: string; }> => {
    const certificationTitle = (cert?.name || cert?.certName || cert?.title || '').toString().trim();
    const certificationCode = (cert?.code || cert?.certCode || '').toString().trim();
    if (!certificationTitle) {
        throw new Error('Certification not found');
    }

    const certificationDetails = await SharePointService.getCertificationDetailsByCodeOrTitle(certificationCode, certificationTitle, true);
    if (!certificationDetails) {
        throw new Error('Certification not found');
    }

    const canonicalTitle = (certificationDetails.title || certificationTitle).toString().trim();
    return {
        certificationId: Number(certificationDetails.id || 0),
        certName: canonicalTitle,
        certCode: (certificationDetails.code || cert?.code || canonicalTitle).toString().trim(),
        pathId: (certificationDetails.code || cert?.pathId || cert?.code || canonicalTitle).toString().trim()
    };
};

interface ICertificationWorkbookRow {
    Title: string;
    CertificationCode: string;
    Provider?: string;
    SourceRow?: {
        sheetName: string;
        rowIndex: number;
        values: any[];
    };
}

const REQUIRED_CERTIFICATION_FIELDS: Array<'Title' | 'CertificationCode'> = [
    'Title',
    'CertificationCode'
];

const validateCertificationWorkbookRow = (row: Partial<ICertificationWorkbookRow>): boolean => {
    if (!row.Title) {
        return false;
    }

    if (!row.CertificationCode) {
        return false;
    }

    return REQUIRED_CERTIFICATION_FIELDS.every((field) => (
        row[field] !== undefined &&
        row[field] !== null &&
        row[field] !== ''
    ));
};

const normalizeExcelHeader = (value: string): string => {
    return (value || '')
        .toString()
        .trim()
        .toLowerCase()
        .replace(/[_-]+/g, ' ')
        .replace(/\s+/g, ' ');
};

const normalizeProviderLabel = (value: string): string => {
    const normalizedValue = (value || '').toString().trim().toLowerCase();

    if (!normalizedValue) {
        return '';
    }

    if (normalizedValue.includes('google') || normalizedValue.includes('gcp')) {
        return 'Google';
    }

    if (normalizedValue.includes('aws') || normalizedValue.includes('amazon')) {
        return 'AWS';
    }

    if (
        normalizedValue.includes('microsoft') ||
        normalizedValue.includes('azure') ||
        normalizedValue.includes('teams') ||
        normalizedValue.includes('m365') ||
        normalizedValue.includes('office 365') ||
        normalizedValue.includes('dynamics') ||
        normalizedValue.includes('power platform') ||
        normalizedValue.includes('power bi') ||
        normalizedValue.includes('fabric') ||
        normalizedValue.includes('sharepoint') ||
        normalizedValue.includes('entra')
    ) {
        return 'Microsoft';
    }

    if (normalizedValue.includes('other')) {
        return 'Other';
    }

    return '';
};

const inferProviderFromWorkbookContext = (sheetName: string, values: any[], title: string, code: string): string => {
    const combinedText = [
        sheetName,
        ...(values || []).map((value) => (value || '').toString()),
        title,
        code
    ].join(' ');

    return normalizeProviderLabel(combinedText) || 'Other';
};

const detectCertificationHeaderIndexes = (values: any[]): { titleIndex: number; codeIndex: number } | null => {
    const normalizedValues = values.map((value) => normalizeExcelHeader((value || '').toString()));
    const titleIndex = normalizedValues.findIndex((value) => value === 'certification name');
    const codeIndex = normalizedValues.findIndex((value) => (
        value === 'exam code' ||
        value === 'certification code / id' ||
        value === 'certification code/id' ||
        value === 'certification code' ||
        value === 'certificationcode/id'
    ));

    if (titleIndex === -1 || codeIndex === -1) {
        return null;
    }

    return { titleIndex, codeIndex };
};

const collectWorkbookRows = (workbook: XLSX.WorkBook): ICertificationWorkbookRow[] => {
    let allRows: ICertificationWorkbookRow[] = [];

    (workbook.SheetNames || []).forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
            return;
        }

        const sheetRows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: '' });
        console.log('Sheet:', sheetName, 'Rows:', sheetRows.length);

        let activeTitleIndex = -1;
        let activeCodeIndex = -1;

        sheetRows.forEach((rawRow, rowIndex) => {
            const values = Array.isArray(rawRow) ? rawRow : [];
            console.log('ROW:', { sheetName, rowIndex, values });
            console.log('VALUES:', values);

            const headerIndexes = detectCertificationHeaderIndexes(values);
            if (headerIndexes) {
                activeTitleIndex = headerIndexes.titleIndex;
                activeCodeIndex = headerIndexes.codeIndex;
                return;
            }

            if (activeTitleIndex === -1 || activeCodeIndex === -1) {
                return;
            }

            const title = (values[activeTitleIndex] || '').toString().trim();
            const code = (values[activeCodeIndex] || '').toString().trim().toUpperCase();
            const provider = inferProviderFromWorkbookContext(sheetName, values, title, code);

            if (!title && !code) {
                return;
            }

            allRows.push({
                Title: title,
                CertificationCode: code,
                Provider: provider,
                SourceRow: {
                    sheetName,
                    rowIndex,
                    values
                }
            });
        });
    });

    console.log('Total Rows:', allRows.length);
    return allRows;
};

function RouterConnectedBrand() {
    const navigate = useNavigate();

    const handleClick = () => {
        navigate('/user-dashboard');
    };

    const handleKeyDown = (event: React.KeyboardEvent<HTMLButtonElement>) => {
        if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault();
            handleClick();
        }
    };

    return (
        <button
            type="button"
            className="sidebar-brand sidebar-brand-button"
            onClick={handleClick}
            onKeyDown={handleKeyDown}
            title="Go to User Dashboard"
        >
            <img src={logo} alt="SkySecure Logo" />
            <span className="brand-text">skysecure</span>
        </button>
    );
}

function PortalBrand(props: { onFallbackNavigate: () => void }) {
    const isInRouter = useInRouterContext();

    const handleFallbackKeyDown = (event: React.KeyboardEvent<HTMLButtonElement>) => {
        if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault();
            props.onFallbackNavigate();
        }
    };

    if (isInRouter) {
        return <RouterConnectedBrand />;
    }

    return (
        <button
            type="button"
            className="sidebar-brand sidebar-brand-button"
            onClick={props.onFallbackNavigate}
            onKeyDown={handleFallbackKeyDown}
            title="Go to User Dashboard"
        >
            <img src={logo} alt="SkySecure Logo" />
            <span className="brand-text">skysecure</span>
        </button>
    );
}

export default function AdminPortal(props: any) {
    // Safety check for props
    const userDisplayName = props?.userDisplayName || 'Admin Portal';
    const userEmail = props?.userEmail || '';
    const userPhotoUrl = userEmail ? `/_layouts/15/userphoto.aspx?size=M&accountname=${userEmail}` : null;
    const context = props?.context;

    const [view, setView] = useState<'DASHBOARD' | 'MANAGEMENT' | 'DETAILS' | 'TRACKER' | 'SECURITY' | 'AUDIT' | 'REPORTS' | 'CONFIG' | 'ASSESSMENTS' | 'USERS' | 'TAXONOMY' | 'CONTENT' | 'ASSIGNMENTS'>('DASHBOARD');
    const [selectedCert, setSelectedCert] = useState<any>(null);
    const [certificationSearchText, setCertificationSearchText] = useState('');
    const [realEnrollments, setRealEnrollments] = useState<any[]>([]);
    const [customCerts, setCustomCerts] = useState<any[]>([]);
    const [pathLibraryState, setPathLibraryState] = useState<{ loading: boolean; refreshing: boolean; error: string | null }>({
        loading: true,
        refreshing: false,
        error: null
    });
    const [certificationMaxSeatsMap, setCertificationMaxSeatsMap] = useState<Map<string, number>>(new Map());
    const [certificationCatalogItems, setCertificationCatalogItems] = useState<ICertificationMaxSeatsItem[]>([]);
    const [showCreateCertificationModal, setShowCreateCertificationModal] = useState(false);
    const [editingCertificationId, setEditingCertificationId] = useState<number | null>(null);
    const [selectedCertificationItem, setSelectedCertificationItem] = useState<ICertificationMaxSeatsItem | null>(null);
    const [newCertificationTitle, setNewCertificationTitle] = useState('');
    const [newCertificationCode, setNewCertificationCode] = useState('');
    const [newCertificationProvider, setNewCertificationProvider] = useState('');
    const [newCertificationLink, setNewCertificationLink] = useState('');
    const [isCreatingCertification, setIsCreatingCertification] = useState(false);
    const [isDeletingCertification, setIsDeletingCertification] = useState(false);
    const [isUploadingCertificationWorkbook, setIsUploadingCertificationWorkbook] = useState(false);
    const [sidebarWidth, setSidebarWidth] = useState(250);
    const [isResizing, setIsResizing] = useState(false);
    const [accessUsers, setAccessUsers] = useState<any[]>([]);
    const [showAddUser, setShowAddUser] = useState(false);
    const [newUserData, setNewUserData] = useState({ name: '', email: '', role: 'Member', bu: '', dept: '' });
    const [toast, setToast] = useState<{ message: string, type: 'success' | 'info' | 'error' } | null>(null);
    const [showProfileOverlay, setShowProfileOverlay] = useState(false);
    const [showAdminNotifications, setShowAdminNotifications] = useState(false);
    const [adminNotifications, setAdminNotifications] = useState<any[]>([
        { id: 1, title: 'System Diagnostics', text: 'SharePoint connection verified.', time: 'Just now', type: 'info' },
        { id: 2, title: 'New User Access', text: 'A new member was authorized for the portal.', time: '10m ago', type: 'success' }
    ]);
    const notificationsMenuRef = useRef<HTMLDivElement | null>(null);
    const profileMenuRef = useRef<HTMLDivElement | null>(null);
    const certificationWorkbookInputRef = useRef<HTMLInputElement | null>(null);
    const dashboardSeatRefreshInFlight = useRef(false);

    // User & Taxonomy States
    const [allUsers, setAllUsers] = useState<any[]>([]);
    const [directorySyncState, setDirectorySyncState] = useState<{ users: any[]; loading: boolean; error: string | null }>({
        users: [],
        loading: false,
        error: null
    });
    const [isDataLoaded, setIsDataLoaded] = useState(false);
    const isInitialDataLoadInFlight = useRef(false);
    const [taxonomyData, setTaxonomyData] = useState<any>({
        departments: [],
        businessUnits: [],
        roles: [],
        locations: [],
        groups: [],
        categories: []
    });
    const [showAddUserModal, setShowAddUserModal] = useState(false);
    const [showTaxonomyModal, setShowTaxonomyModal] = useState(false);
    const [activeTaxonomyTab, setActiveTaxonomyTab] = useState('departments');
    
    // Assign Certificate State
    const [showAssignCertModal, setShowAssignCertModal] = useState(false);
    const [certToAssign, setCertToAssign] = useState<any>(null);
    const [selectedUsersForCert, setSelectedUsersForCert] = useState<number[]>([]);
    const [certUserSearchTerm, setCertUserSearchTerm] = useState('');
    const [directoryUsers, setDirectoryUsers] = useState<any[]>([]);
    const [assignModalLearners, setAssignModalLearners] = useState<any[]>([]);
    const [filteredAssignModalLearners, setFilteredAssignModalLearners] = useState<any[]>([]);
    const [isLoadingAssignLearners, setIsLoadingAssignLearners] = useState(false);
    const [isAssigningCert, setIsAssigningCert] = useState(false);
    const [certExamScheduledDate, setCertExamScheduledDate] = useState(new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]);

    const normalizePathLookupValue = (value: any): string =>
        (value === null || value === undefined ? '' : value.toString()).trim().toLowerCase();

    const enrollmentCountByPath = useMemo(() => {
        const seatMap = new Map<string, Set<string>>();
        const registerSeatKey = (key: any, email: string): void => {
            const normalizedKey = normalizePathLookupValue(key);
            if (!normalizedKey || !email) {
                return;
            }

            if (!seatMap.has(normalizedKey)) {
                seatMap.set(normalizedKey, new Set<string>());
            }

            seatMap.get(normalizedKey)!.add(email);
        };

        (realEnrollments || []).forEach((enrollment: any) => {
            const normalizedEmail = (enrollment.email || enrollment.userEmail || '').toString().trim().toLowerCase();
            if (!normalizedEmail) {
                return;
            }

            [
                enrollment.pathId,
                enrollment.code,
                enrollment.certCode,
                enrollment.name,
                enrollment.certName,
                enrollment.certificateName
            ].forEach((key) => registerSeatKey(key, normalizedEmail));
        });

        return Array.from(seatMap.entries()).reduce((acc: Record<string, number>, [pathId, emails]) => {
            acc[pathId] = emails.size;
            return acc;
        }, {});
    }, [realEnrollments]);

    const certificationCatalogLookup = useMemo(() => {
        const lookup = new Map<string, ICertificationMaxSeatsItem>();

        (certificationCatalogItems || []).forEach((item) => {
            const normalizedTitle = normalizePathLookupValue(item?.title);
            if (normalizedTitle) {
                lookup.set(`title:${normalizedTitle}`, item);
            }

            const normalizedCode = normalizePathLookupValue(item?.code);
            if (normalizedCode) {
                lookup.set(`code:${normalizedCode}`, item);
            }
        });

        return lookup;
    }, [certificationCatalogItems]);

    const findCertificationCatalogItem = (pathItem: any): ICertificationMaxSeatsItem | null => {
        const itemId = Number(pathItem?.id || 0);
        if (Number.isFinite(itemId) && itemId > 0) {
            const matchedById = (certificationCatalogItems || []).find((item) => Number(item.id) === itemId);
            if (matchedById) {
                return matchedById;
            }
        }

        const normalizedCode = normalizePathLookupValue(pathItem?.code || pathItem?.certCode);
        if (normalizedCode) {
            const matchedByCode = certificationCatalogLookup.get(`code:${normalizedCode}`);
            if (matchedByCode) {
                return matchedByCode;
            }
        }

        const normalizedTitle = normalizePathLookupValue(pathItem?.name || pathItem?.certName || pathItem?.title);
        if (normalizedTitle) {
            const matchedByTitle = certificationCatalogLookup.get(`title:${normalizedTitle}`);
            if (matchedByTitle) {
                return matchedByTitle;
            }
        }

        return null;
    };

    const getSeatSummaryForPath = (pathItem: any): any => {
        const certificationCatalogItem = findCertificationCatalogItem(pathItem);
        const certificationTitleKey = normalizePathLookupValue(certificationCatalogItem?.title || pathItem?.name || pathItem?.certName || pathItem?.title);
        const certificationCodeKey = normalizePathLookupValue(certificationCatalogItem?.code || pathItem?.code || pathItem?.certCode);
        const rawPathId = (certificationCatalogItem?.code || pathItem?.pathId || pathItem?.code || certificationCatalogItem?.title || pathItem?.name || '').toString().trim();
        const normalizedPathId = normalizePathLookupValue(rawPathId);
        const certificationCountMap = certificationMaxSeatsMap;
        const certificationListCount = certificationCatalogItem
            ? Number(
                certificationCatalogItem.assignedLearnerCount ??
                certificationCatalogItem.enrolledCount ??
                certificationCatalogItem.maxSeats ??
                0
            )
            : certificationCodeKey
                ? Number(certificationCountMap.get(certificationCodeKey) || 0)
                : certificationTitleKey
                    ? Number(certificationCountMap.get(certificationTitleKey) || 0)
                    : 0;
        const seatCountKey = certificationCodeKey || certificationTitleKey || normalizedPathId || normalizePathLookupValue(pathItem?.code);
        const liveEnrollmentCount = seatCountKey ? (enrollmentCountByPath[seatCountKey] || 0) : 0;
        const assignedLearnerCount = Math.max(
            Number(certificationListCount || liveEnrollmentCount || pathItem?.assignedLearnerCount || pathItem?.enrolledCount || pathItem?.maxSeats || 0),
            0
        );
        const enrollments = (realEnrollments || []).filter((enrollment: any) => {
            if (certificationCodeKey) {
                return normalizePathLookupValue(enrollment.code || enrollment.certCode || enrollment.pathId) === certificationCodeKey;
            }

            if (certificationTitleKey) {
                return normalizePathLookupValue(enrollment.certName || enrollment.certificateName || enrollment.name) === certificationTitleKey;
            }

            if (normalizedPathId) {
                return normalizePathLookupValue(enrollment.pathId) === normalizedPathId;
            }

            return normalizePathLookupValue(enrollment.code) === normalizePathLookupValue(pathItem?.code) ||
                normalizePathLookupValue(enrollment.name) === normalizePathLookupValue(pathItem?.name);
        });

        return {
            ...pathItem,
            id: certificationCatalogItem?.id || pathItem?.id,
            name: certificationCatalogItem?.title || pathItem?.name || pathItem?.title,
            code: certificationCatalogItem?.code || pathItem?.code || pathItem?.certCode || '',
            pathId: rawPathId,
            occupiedSeats: assignedLearnerCount,
            enrolledCount: assignedLearnerCount,
            assignedLearnerCount,
            maxSeats: assignedLearnerCount,
            enrollments,
            isSharePointManaged: !!certificationCatalogItem,
            status: 'OPEN'
        };
    };

    const handleOpenAssignModal = (cert: any) => {
        const seatManagedCert = getSeatSummaryForPath(cert);

        if (!seatManagedCert?.name) {
            showToast('Certification not found', 'error');
            return;
        }

        setCertToAssign({
            ...seatManagedCert
        });
        setShowAssignCertModal(true);
        setSelectedUsersForCert([]);
        setCertUserSearchTerm('');
        setCertExamScheduledDate(new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]);
    };

    const handleAssignCert = async () => {
        if (selectedUsersForCert.length === 0 || !certToAssign || isAssigningCert) return;
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const selectedExamDate = new Date(certExamScheduledDate);

        if (!certExamScheduledDate || Number.isNaN(selectedExamDate.getTime()) || selectedExamDate < today) {
            showToast('Select a future exam date before pushing the certification.', 'error');
            return;
        }

        setIsAssigningCert(true);
        const users = assignModalLearners.length > 0
            ? assignModalLearners
            : getAssignableLearners(
                normalizeLearnerSelectionUsers(directoryUsers.length > 0 ? directoryUsers : directorySyncState.users)
            );

        if (users.length === 0) {
            showToast('Learners are still syncing from SharePoint. Wait a moment and try again.', 'info');
            setIsAssigningCert(false);
            return;
        }
        const selectedLookup = new Set(selectedUsersForCert);
        const targetUsers = users.filter((user: any, index: number) => selectedLookup.has(getLearnerSelectionId(user, index)));

        try {
            const assignmentDetails = await resolveCertificationAssignmentDetails(certToAssign);

            const duplicateChecks = await Promise.all(
                targetUsers.map(async (userObj: any) => ({
                    userObj,
                    alreadyEnrolled: await SharePointService.hasEnrollmentForUserCertification(userObj.email, assignmentDetails.certName, assignmentDetails.certCode)
                }))
            );

            const usersToAssign = duplicateChecks.filter((entry) => !entry.alreadyEnrolled).map((entry) => entry.userObj);
            const skippedUsers = duplicateChecks
                .filter((entry) => entry.alreadyEnrolled)
                .map((entry) => entry.userObj.name || (entry.userObj.email || '').toLowerCase());

            if (usersToAssign.length === 0) {
                showToast('Already enrolled', 'info');
                return;
            }

            const assignedUsers = usersToAssign.map((userObj: any) => userObj.name || (userObj.email || '').toLowerCase());
            const chunkSize = usersToAssign.length > 20 ? 10 : Math.max(usersToAssign.length, 1);
            for (let index = 0; index < usersToAssign.length; index += chunkSize) {
                const chunk = usersToAssign.slice(index, index + chunkSize);
                await Promise.all(
                    chunk.map((userObj: any) => SharePointService.createEnrollmentForCertificationAssignment({
                        userEmail: userObj.email,
                        userName: userObj.name,
                        certCode: assignmentDetails.certCode,
                        certName: assignmentDetails.certName,
                        pathId: assignmentDetails.pathId,
                        assignedByName: userDisplayName,
                        examScheduledDate: certExamScheduledDate
                    }))
                );

                if (usersToAssign.length > 20 && index + chunkSize < usersToAssign.length) {
                    await new Promise((resolve) => window.setTimeout(resolve, 150));
                }
            }

            const refreshedEnrollments = await refreshEnrollments();
            setRealEnrollments(refreshedEnrollments);
            await loadCertificationMaxSeatsData(true);
            dispatchEnrollmentRefreshSignal();
            window.setTimeout(() => {
                window.dispatchEvent(new Event(LMS_AUDIT_REFRESH_EVENT));
            }, 500);

            if (assignedUsers.length > 0) {
                showToast(`Pushed ${assignmentDetails.certCode || assignmentDetails.certName} to ${assignedUsers.length} learner${assignedUsers.length > 1 ? 's' : ''}.`);
            }

            if (skippedUsers.length > 0) {
                console.info('Skipped direct learner assignments because they already existed:', skippedUsers);
            }

            setShowAssignCertModal(false);
            setCertToAssign(null);
            setSelectedUsersForCert([]);
            setCertUserSearchTerm('');
            setCertExamScheduledDate(new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]);
        } catch (error) {
            console.error('Failed to assign certification from the learner picker', error);
            showToast(getEnrollmentAssignmentErrorMessage(error), 'error');
        } finally {
            setIsAssigningCert(false);
        }
    };

    const [newLmsUserData, setNewLmsUserData] = useState({
        employeeId: '',
        name: '',
        email: '',
        department: '',
        businessUnit: '',
        role: '',
        location: '',
        status: 'Active'
    });
    const [newTaxonomyItemData, setNewTaxonomyItemData] = useState('');

    const [config, setConfig] = useState({
        darkMode: false,
        notifications: true,
        autoArchive: false,
        accentColor: 'blue',
        maintenanceMode: false
    });

    // Initialize SharePointService with context on component mount
    useEffect(() => {
        if (context && context.spHttpClient && context.pageContext?.web?.absoluteUrl) {
            try {
                console.log('Initializing SharePointService in AdminPortal with URL:', context.pageContext.web.absoluteUrl);
                SharePointService.init(context.pageContext.web.absoluteUrl, context.spHttpClient, context);
            } catch (e) {
                console.error('Failed to initialize SharePointService:', e);
            }
        } else {
            console.warn('AdminPortal: context, spHttpClient, or absoluteUrl missing for SharePointService initialization');
        }
    }, [context]);

    // Load config and notifications from SP & local storage
    useEffect(() => {
        const savedConfig = localStorage.getItem('portalAdminConfig');
        if (savedConfig) {
            try { setConfig(JSON.parse(savedConfig)); } catch (e) { }
        }
        
        const loadSPNotifs = async () => {
            try {
                const spNotifs = await SharePointService.getNotifications('Admin');
                if (spNotifs && spNotifs.length > 0) {
                    setAdminNotifications(spNotifs);
                } else {
                    const savedNotifs = localStorage.getItem('adminNotifications');
                    if (savedNotifs) setAdminNotifications(JSON.parse(savedNotifs));
                }
            } catch (e) {
                console.error('Failed to load SharePoint admin notifications, falling back to local storage:', e);
                const savedNotifs = localStorage.getItem('adminNotifications');
                if (savedNotifs) setAdminNotifications(JSON.parse(savedNotifs));
            }
        };

        void loadSPNotifs();
    }, []);

    useEffect(() => {
        const handleClickOutside = (event: MouseEvent) => {
            const target = event.target as Node;

            if (showAdminNotifications && notificationsMenuRef.current && !notificationsMenuRef.current.contains(target)) {
                setShowAdminNotifications(false);
            }

            if (showProfileOverlay && profileMenuRef.current && !profileMenuRef.current.contains(target)) {
                setShowProfileOverlay(false);
            }
        };

        document.addEventListener('mousedown', handleClickOutside);
        return () => document.removeEventListener('mousedown', handleClickOutside);
    }, [showAdminNotifications, showProfileOverlay]);

    const updateAdminNotifications = async (newNotif: any) => {
        const updated = [newNotif, ...adminNotifications].slice(0, 20);
        setAdminNotifications(updated);
        localStorage.setItem('adminNotifications', JSON.stringify(updated));

        // PUSH TO SHAREPOINT FOR REAL-TIME PERMANENCE
        try {
            await SharePointService.addNotification({
                title: newNotif.title,
                text: newNotif.text,
                targetEmail: newNotif.targetEmail || 'Admin',
                type: newNotif.type || 'info',
                time: new Date().toISOString(),
                read: false
            });
        } catch (e) {
            console.warn("Failed to push notification to SharePoint", e);
        }
    };

    const handleMarkAdminRead = async () => {
        const unreadNotifications = adminNotifications.filter(n => !n.read);

        if (unreadNotifications.length === 0) {
            showToast('All notifications are already marked as read.', 'info');
            return;
        }

        const updatedNotifications = adminNotifications.map(n => ({ ...n, read: true }));
        setAdminNotifications(updatedNotifications);
        localStorage.setItem('adminNotifications', JSON.stringify(updatedNotifications));

        try {
            const syncResults = await Promise.all(
                unreadNotifications
                    .filter(n => typeof n.id === 'number')
                    .map(async n => {
                        try {
                            await SharePointService.markNotificationAsRead(n.id);
                            return { ok: true };
                        } catch (error) {
                            return { ok: false, error };
                        }
                    })
            );

            const failedUpdates = syncResults.filter(result => !result.ok);
            if (failedUpdates.length > 0) {
                console.error('Failed to update one or more admin notifications in SharePoint', failedUpdates);
                showToast('Notifications marked as read locally. SharePoint sync partially failed.', 'info');
                return;
            }

            showToast('All admin notifications marked as read.');
        } catch (e) {
            console.error("Failed to mark admin notifications as read", e);
            showToast('Notifications marked as read locally. SharePoint sync failed.', 'info');
        }
    };

    const handleLogoNavigate = () => {
        const fallbackUrl = context?.pageContext?.web?.absoluteUrl || window.location.origin;
        window.location.assign(fallbackUrl);
    };

    const updateConfig = (newConfig: any) => {
        const updated = { ...config, ...newConfig };
        setConfig(updated);
        localStorage.setItem('portalAdminConfig', JSON.stringify(updated));
        showToast('Settings successfully updated');
    };

    // --- Resizing Handlers ---
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
            if (newWidth > 180 && newWidth < 500) {
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

    // --- Toast Handler ---
    const showToast = (message: string, type: 'success' | 'info' | 'error' = 'success') => {
        setToast({ message, type });
        setTimeout(() => setToast(null), 3000);
    };

    const hasChanged = (currentValue: any, nextValue: any): boolean =>
        JSON.stringify(currentValue) !== JSON.stringify(nextValue);

    const getLearnerSelectionId = (user: any, fallbackIndex: number = 0): number => {
        const rawId = Number(user?.userId || user?.Id || user?.id);
        if (!isNaN(rawId) && rawId > 0) {
            return rawId;
        }

        const stableKey = ((user?.email || user?.login || user?.employeeId || `learner-${fallbackIndex}`) as string).toLowerCase();
        let hash = 0;

        for (let i = 0; i < stableKey.length; i += 1) {
            hash = ((hash << 5) - hash + stableKey.charCodeAt(i)) | 0;
        }

        return Math.abs(hash) || fallbackIndex + 1;
    };

    const getCachedDirectoryUsers = (): any[] => {
        if (directoryUsers.length > 0) {
            return directoryUsers;
        }

        if (directorySyncState.users.length > 0) {
            return directorySyncState.users;
        }

        return allUsers;
    };

    const normalizeLearnerSelectionUsers = (users: any[]): any[] => {
        const deduped = new Map<string, any>();
        const rolePriority: Record<string, number> = {
            Owner: 0,
            Member: 1,
            Visitor: 2
        };

        (users || []).forEach((user: any, index: number) => {
            const rawEmail = (user?.email || user?.Email || '').toString().trim();
            if (!rawEmail) {
                return;
            }

            const normalizedEmail = rawEmail.toLowerCase();
            const role = (user?.role || '').toString().trim() || (
                user?.siteGroup === 'Owners' ? 'Owner' :
                    user?.siteGroup === 'Visitors' ? 'Visitor' :
                        'Member'
            );
            const siteGroup = (user?.siteGroup || user?.group || '').toString().trim() || (
                role === 'Owner' ? 'Owners' :
                    role === 'Visitor' ? 'Visitors' :
                        'Members'
            );

            const normalizedUser = {
                ...user,
                id: user?.id || user?.Id || rawEmail || `learner-selection-${index}`,
                name: user?.name || user?.Title || rawEmail,
                email: rawEmail,
                role,
                siteGroup
            };

            const existing = deduped.get(normalizedEmail);
            if (!existing || (rolePriority[role] ?? 99) < (rolePriority[existing.role] ?? 99)) {
                deduped.set(normalizedEmail, normalizedUser);
            }
        });

        return Array.from(deduped.values()).sort((left: any, right: any) => {
            const roleDelta = (rolePriority[left.role] ?? 99) - (rolePriority[right.role] ?? 99);
            if (roleDelta !== 0) {
                return roleDelta;
            }

            return (left.name || '').localeCompare(right.name || '');
        });
    };

    const persistDirectoryUsers = (users: any[]): any[] => {
        const deduped = new Map<string, any>();
        const rolePriority: Record<string, number> = {
            Owner: 0,
            Member: 1,
            Visitor: 2
        };

        (users || []).forEach((user: any, index: number) => {
            const email = (user?.email || user?.Email || user?.userPrincipalName || '').trim();
            const login = (user?.login || user?.LoginName || '').trim();
            const dedupeKey = (email || login).toLowerCase();

            if (!dedupeKey) {
                return;
            }

            const role = user?.role || 'Visitor';
            const normalizedUser = {
                id: getLearnerSelectionId(user, index),
                employeeId: user?.employeeId || `EMP-${user?.id || index + 1}`,
                name: user?.name || user?.displayName || user?.Title || email || login,
                email,
                login,
                department: user?.department || '',
                businessUnit: user?.businessUnit || user?.officeLocation || '',
                role,
                siteGroup: user?.siteGroup || user?.group || `${role}s`,
                jobTitle: user?.JobTitle || user?.jobTitle || '',
                officeLocation: user?.officeLocation || '',
                status: user?.status || 'Active',
                progress: user?.progress || 0
            };

            const existingUser = deduped.get(dedupeKey);
            if (!existingUser || (rolePriority[role] ?? 99) < (rolePriority[existingUser.role] ?? 99)) {
                deduped.set(dedupeKey, normalizedUser);
            }
        });

        const normalizedUsers = Array.from(deduped.values()).sort((a, b) => {
            const roleDelta = (rolePriority[a.role] ?? 99) - (rolePriority[b.role] ?? 99);
            if (roleDelta !== 0) {
                return roleDelta;
            }

            return (a.name || '').localeCompare(b.name || '');
        });

        setAllUsers((prev: any[]) => hasChanged(prev, normalizedUsers) ? normalizedUsers : prev);
        setDirectoryUsers((prev: any[]) => hasChanged(prev, normalizedUsers) ? normalizedUsers : prev);
        setDirectorySyncState({
            users: normalizedUsers,
            loading: false,
            error: null
        });
        return normalizedUsers;
    };

    const loadDirectoryUsers = async (
        forceRefresh: boolean = false,
        options: { showLoadingIndicator?: boolean } = {}
    ): Promise<{ users: any[]; error: string | null }> => {
        const cachedUsers = getCachedDirectoryUsers();

        if (!forceRefresh && isDataLoaded && cachedUsers.length > 0) {
            return {
                users: cachedUsers,
                error: directorySyncState.error
            };
        }

        if (cachedUsers.length > 0) {
            setAllUsers((prev: any[]) => hasChanged(prev, cachedUsers) ? cachedUsers : prev);
            setDirectoryUsers((prev: any[]) => hasChanged(prev, cachedUsers) ? cachedUsers : prev);
            setDirectorySyncState((prev) => ({
                users: prev.users.length > 0 ? prev.users : cachedUsers,
                loading: false,
                error: prev.error
            }));
        }

        const siteUrl = (context?.pageContext?.web?.absoluteUrl || '').toString().trim();
        if (!context?.spHttpClient || !siteUrl || siteUrl.toLowerCase().indexOf('localhost') !== -1) {
            const contextError = 'SharePoint tenant context is unavailable. Open this web part on the production SharePoint site instead of localhost/workbench.';

            setDirectorySyncState((prev) => ({
                users: prev.users.length > 0 ? prev.users : cachedUsers,
                loading: false,
                error: contextError
            }));
            setIsDataLoaded(true);
            isInitialDataLoadInFlight.current = false;

            return {
                users: cachedUsers,
                error: contextError
            };
        }

        const shouldShowLoading = options.showLoadingIndicator ?? cachedUsers.length === 0;
        isInitialDataLoadInFlight.current = !forceRefresh;
        setDirectorySyncState((prev) => ({
            users: prev.users.length > 0 ? prev.users : cachedUsers,
            loading: shouldShowLoading,
            error: null
        }));

        try {
            console.log('[AdminPortal] Loading learner directory users', {
                forceRefresh,
                siteUrl
            });

            const users = await SharePointService.getLearnerDirectoryUsers(forceRefresh);
            const normalizedUsers = persistDirectoryUsers(users || []);
            setIsDataLoaded(true);

            if (normalizedUsers.length > 0) {
                return {
                    users: normalizedUsers,
                    error: null
                };
            }

            const emptyMessage = 'No users were returned from the SharePoint Learners directory source.';
            console.warn('[AdminPortal] Learner directory source returned an empty response', {
                siteUrl,
                forceRefresh
            });

            setAllUsers((prev: any[]) => hasChanged(prev, []) ? [] : prev);
            setDirectoryUsers((prev: any[]) => hasChanged(prev, []) ? [] : prev);
            setDirectorySyncState({
                users: [],
                loading: false,
                error: emptyMessage
            });
            setIsDataLoaded(true);

            return {
                users: [],
                error: emptyMessage
            };
        } catch (e) {
            const errorMessage = e instanceof Error ? e.message : 'Failed to load users from the SharePoint Learners directory source.';
            console.error('[AdminPortal] Failed to load learner directory users', {
                error: e,
                siteUrl,
                forceRefresh
            });

            setDirectorySyncState((prev) => ({
                users: prev.users.length > 0 ? prev.users : cachedUsers,
                loading: false,
                error: errorMessage
            }));
            setIsDataLoaded(true);

            return {
                users: cachedUsers,
                error: errorMessage
            };
        } finally {
            if (!forceRefresh) {
                isInitialDataLoadInFlight.current = false;
            }
        }
    };

    const refreshEnrollments = async (): Promise<any[]> => {
        try {
            const spEnrollments = await SharePointService.getEnrollments('');
            const mapped = (spEnrollments || []).map(e => ({
                id: e.id,
                name: e.certName,
                code: e.certCode || e.certName,
                certificationId: e.certificationId,
                userName: e.userName,
                email: e.userEmail,
                userId: e.userId,
                pathId: e.pathId,
                startDate: e.startDate,
                endDate: e.endDate,
                status: e.status,
                progress: e.progress,
                assignedDate: e.assignedDate,
                assignedByName: e.assignedByName,
                assignedByAdmin: e.assignedByAdmin,
                examScheduledDate: e.examScheduledDate,
                rescheduledDate: e.rescheduledDate,
                completionDate: e.completionDate,
                examCode: e.examCode,
                listStatus: e.listStatus
            }));

            setRealEnrollments((prev: any[]) => hasChanged(prev, mapped) ? mapped : prev);
            return mapped;
        } catch (e) {
            console.warn('Failed to fetch enrollments from SharePoint', e);
            return [];
        }
    };

    const loadCertificationMaxSeatsData = async (forceRefresh: boolean = false): Promise<Map<string, number>> => {
        const siteUrl = (context?.pageContext?.web?.absoluteUrl || '').toString().trim();
        if (!context?.spHttpClient || !siteUrl || siteUrl.toLowerCase().indexOf('localhost') !== -1) {
            console.warn('[Certifications] Skipping MaxSeats sync because SharePoint site context is unavailable.', {
                siteUrl,
                hasSpHttpClient: !!context?.spHttpClient
            });
            return certificationMaxSeatsMap;
        }

        try {
            setPathLibraryState((prev) => ({
                ...prev,
                loading: prev.loading && !forceRefresh,
                refreshing: forceRefresh,
                error: null
            }));

            const certificationItems = await SharePointService.fetchCertificationMaxSeats(forceRefresh);
            const nextMap = SharePointService.createCertificationMaxSeatsMap(certificationItems);
            setCertificationCatalogItems((prev) => hasChanged(prev, certificationItems) ? certificationItems : prev);
            setCertificationMaxSeatsMap((prev) => {
                const prevEntries = JSON.stringify(Array.from(prev.entries()));
                const nextEntries = JSON.stringify(Array.from(nextMap.entries()));
                return prevEntries !== nextEntries ? new Map(nextMap) : prev;
            });
            setPathLibraryState({
                loading: false,
                refreshing: false,
                error: null
            });
            return nextMap;
        } catch (error) {
            console.error('[Certifications] Failed to fetch MaxSeats map from SharePoint', error);
            setPathLibraryState({
                loading: false,
                refreshing: false,
                error: error instanceof Error ? error.message : 'Failed to load the SharePoint Certifications list.'
            });
            return certificationMaxSeatsMap;
        }
    };

    // --- Data Management & Sync ---
    useEffect(() => {
        if (!context?.spHttpClient || !context?.pageContext?.web?.absoluteUrl || isDataLoaded || isInitialDataLoadInFlight.current) {
            return;
        }

        let disposed = false;
        isInitialDataLoadInFlight.current = true;

        const loadInitialData = async (): Promise<void> => {
            try {
                const initialEnrollments = await refreshEnrollments();
                await SharePointService.ensureLearnerEditAccessForEnrollments(initialEnrollments).catch((error) => {
                    console.warn('[Enrollments] Failed to sync learner edit access on existing admin assignments.', error);
                });
                await loadCertificationMaxSeatsData(false);

                const custom = localStorage.getItem('selfExploreCerts');
                if (!disposed && custom) {
                    setCustomCerts(JSON.parse(custom) || []);
                }

                await loadDirectoryUsers(false, { showLoadingIndicator: true });

                try {
                    const spTax = await SharePointService.getTaxonomy();
                    if (disposed) {
                        return;
                    }

                    if (spTax) {
                        setTaxonomyData((prev: any) => hasChanged(prev, spTax) ? spTax : prev);
                        localStorage.setItem('lmsTaxonomyData', JSON.stringify(spTax));
                    } else {
                        const taxData = localStorage.getItem('lmsTaxonomyData');
                        if (taxData) {
                            const parsedTax = JSON.parse(taxData);
                            setTaxonomyData((prev: any) => hasChanged(prev, parsedTax) ? parsedTax : prev);
                        }
                    }
                } catch (e) {
                    const taxData = localStorage.getItem('lmsTaxonomyData');
                    if (!disposed && taxData) {
                        const parsedTax = JSON.parse(taxData);
                        setTaxonomyData((prev: any) => hasChanged(prev, parsedTax) ? parsedTax : prev);
                    }
                }
            } catch (e) {
                console.error("Initial SharePoint data load error:", e);
            } finally {
                if (!disposed) {
                    setIsDataLoaded(true);
                }
                isInitialDataLoadInFlight.current = false;
            }
        };

        void loadInitialData();

        return () => {
            disposed = true;
            isInitialDataLoadInFlight.current = false;
        };
    }, [context, isDataLoaded]);

    useEffect(() => {
        if (!isDataLoaded || !context?.spHttpClient || !context?.pageContext?.web?.absoluteUrl) {
            return;
        }

        const handleEnrollmentRefresh = () => {
            void refreshEnrollments();
        };

        window.addEventListener(LMS_ENROLLMENTS_REFRESH_EVENT, handleEnrollmentRefresh);

        return () => {
            window.removeEventListener(LMS_ENROLLMENTS_REFRESH_EVENT, handleEnrollmentRefresh);
        };
    }, [context?.pageContext?.web?.absoluteUrl, context?.spHttpClient, isDataLoaded]);

    useEffect(() => {
        if (
            view !== 'DASHBOARD' ||
            !isDataLoaded ||
            !context?.spHttpClient ||
            !context?.pageContext?.web?.absoluteUrl
        ) {
            return;
        }

        let disposed = false;

        const refreshDashboardSeatData = async (): Promise<void> => {
            if (disposed || dashboardSeatRefreshInFlight.current) {
                return;
            }

            dashboardSeatRefreshInFlight.current = true;
            try {
                await Promise.all([
                    refreshEnrollments(),
                    loadCertificationMaxSeatsData(true)
                ]);
            } catch (error) {
                console.warn('[Dashboard] Failed to refresh live seat allocation data', error);
            } finally {
                dashboardSeatRefreshInFlight.current = false;
            }
        };

        void refreshDashboardSeatData();
        const intervalId = window.setInterval(() => {
            void refreshDashboardSeatData();
        }, 5000);

        return () => {
            disposed = true;
            window.clearInterval(intervalId);
        };
    }, [context?.pageContext?.web?.absoluteUrl, context?.spHttpClient, isDataLoaded, view]);

    const allCerts = useMemo(() => {
        const seenCertifications = new Set<string>();

        return (certificationCatalogItems || []).reduce((list: any[], certificationItem) => {
            const seatManagedCert = getSeatSummaryForPath({
                id: certificationItem.id,
                name: certificationItem.title,
                title: certificationItem.title,
                code: certificationItem.code,
                maxSeats: certificationItem.maxSeats,
                enrolledCount: certificationItem.enrolledCount,
                assignedLearnerCount: certificationItem.assignedLearnerCount,
                isDeletable: false,
                isSharePointManaged: true
            });
            const dedupeKey = normalizePathLookupValue(seatManagedCert.code || seatManagedCert.name || seatManagedCert.pathId);
            if (dedupeKey && seenCertifications.has(dedupeKey)) {
                return list;
            }

            if (dedupeKey) {
                seenCertifications.add(dedupeKey);
            }

            list.push({
                ...seatManagedCert,
                category: certificationItem.category || 'Others',
                provider: certificationItem.provider || '',
                level: certificationItem.level || '',
                link: certificationItem.link || '',
                isDeletable: false,
                isSharePointManaged: true
            });
            return list;
        }, []);
    }, [certificationCatalogItems, realEnrollments, enrollmentCountByPath, certificationMaxSeatsMap]);

    const filteredCerts = useMemo(() => {
        const term = (certificationSearchText || '').toLowerCase().trim();
        return (allCerts || []).filter(c =>
            (c.name || '').toLowerCase().includes(term) ||
            (c.code || '').toLowerCase().includes(term) ||
            (c.provider || '').toLowerCase().includes(term)
        );
    }, [allCerts, certificationSearchText]);

    const stats = useMemo(() => {
        const activeEmails = new Set((allUsers || []).map(u => (u.email || '').toLowerCase()));
        const filteredEnrollments = (realEnrollments || []).filter(e => activeEmails.has((e.email || '').toLowerCase()));
        
        const total = filteredEnrollments.length;
        const completedCount = filteredEnrollments.filter(e => e.status === 'completed').length;
        const rate = total > 0 ? (completedCount / total) * 100 : 0;

        return {
            totalCerts: allCerts.length,
            totalEnrolled: total,
            inProgress: filteredEnrollments.filter(e => e.status === 'scheduled').length,
            completed: completedCount,
            completionRate: rate.toFixed(1),
            totalLearners: allUsers.length
        };
    }, [allCerts, realEnrollments, allUsers]);

    const handleDeleteEnrollment = async (id: any) => {
        if (!window.confirm("Permanently revoke this user's enrollment?")) return;

        const enrollmentToDelete = (realEnrollments || []).find((item: any) => Number(item.id) === Number(id));

        try {
            await SharePointService.deleteEnrollment(Number(id));

            if (enrollmentToDelete?.email || enrollmentToDelete?.userEmail) {
                await SharePointService.addAuditLogEntry({
                    title: 'Enrollment Deleted',
                    learnerEmail: (enrollmentToDelete.email || enrollmentToDelete.userEmail || '').toString(),
                    learnerName: (enrollmentToDelete.userName || enrollmentToDelete.name || '').toString(),
                    action: 'Deleted',
                    assignmentName: (enrollmentToDelete.name || enrollmentToDelete.certName || '').toString(),
                    pathId: (enrollmentToDelete.pathId || enrollmentToDelete.code || enrollmentToDelete.certCode || '').toString(),
                    assignmentDate: new Date().toISOString(),
                    status: 'Deleted'
                });
            }

            const refreshed = await refreshEnrollments();
            setRealEnrollments(refreshed);
            await loadCertificationMaxSeatsData(true);
            window.setTimeout(() => {
                window.dispatchEvent(new Event(LMS_AUDIT_REFRESH_EVENT));
            }, 300);
            showToast('Enrollment successfully revoked.', 'info');
        } catch (error) {
            console.error('Failed to revoke enrollment from SharePoint', error);
            showToast('Enrollment revoke failed.', 'error');
        }
    };

    const openEditModal = (path: any): void => {
        const seatManagedPath = getSeatSummaryForPath(path);
        const selectedCertification = findCertificationCatalogItem(seatManagedPath);

        if (!selectedCertification?.id) {
            showToast('This certification was not found in the SharePoint Certifications list.', 'error');
            return;
        }

        setEditingCertificationId(Number(selectedCertification.id));
        setSelectedCertificationItem({
            ...selectedCertification,
            title: (selectedCertification.title || seatManagedPath.name || '').toString().trim(),
            code: (selectedCertification.code || seatManagedPath.code || '').toString().trim().toUpperCase(),
            provider: normalizeProviderLabel(selectedCertification.provider || seatManagedPath.provider || 'Other') || 'Other',
            link: (selectedCertification.link || '').toString().trim(),
            maxSeats: Math.max(Number(
                selectedCertification.assignedLearnerCount ??
                selectedCertification.enrolledCount ??
                selectedCertification.maxSeats ??
                seatManagedPath.assignedLearnerCount ??
                seatManagedPath.enrolledCount ??
                seatManagedPath.maxSeats ??
                0
            ), 0)
        });
        setIsDeletingCertification(false);
        setShowCreateCertificationModal(true);
    };

    const openCreateCertificationModal = (): void => {
        setEditingCertificationId(null);
        setSelectedCertificationItem(null);
        setNewCertificationTitle('');
        setNewCertificationCode('');
        setNewCertificationProvider('Microsoft');
        setNewCertificationLink('');
        setIsDeletingCertification(false);
        setShowCreateCertificationModal(true);
    };

    const resetCreateCertificationModal = (): void => {
        setShowCreateCertificationModal(false);
        setEditingCertificationId(null);
        setSelectedCertificationItem(null);
        setNewCertificationTitle('');
        setNewCertificationCode('');
        setNewCertificationProvider('Microsoft');
        setNewCertificationLink('');
        setIsDeletingCertification(false);
    };

    const closeCreateCertificationModal = (): void => {
        if (isCreatingCertification || isDeletingCertification) {
            return;
        }

        resetCreateCertificationModal();
    };

    useEffect(() => {
        if (!selectedCertificationItem) {
            return;
        }

        setNewCertificationTitle((selectedCertificationItem.title || '').toString().trim());
        setNewCertificationCode((selectedCertificationItem.code || '').toString().trim().toUpperCase());
        setNewCertificationProvider(normalizeProviderLabel(selectedCertificationItem.provider || 'Other') || 'Other');
        setNewCertificationLink((selectedCertificationItem.link || '').toString().trim());
    }, [selectedCertificationItem]);

    const syncCertificationCatalogItemLocally = (
        certificationId: number,
        values: { title: string; code: string; provider: string; link: string; assignedLearnerCount: number; }
    ): void => {
        const normalizedId = Number(certificationId || 0);
        if (normalizedId <= 0) {
            return;
        }

        const normalizedProvider = (values.provider || 'Other').toString().trim().toLowerCase();
        setCertificationCatalogItems((prev: any[]) => {
            let hasMatch = false;
            const next = (prev || []).map((item: any) => {
                if (Number(item?.id || 0) !== normalizedId) {
                    return item;
                }

                hasMatch = true;
                return {
                    ...item,
                    title: values.title,
                    code: values.code,
                    provider: normalizedProvider,
                    link: values.link,
                    maxSeats: values.assignedLearnerCount,
                    enrolledCount: values.assignedLearnerCount,
                    assignedLearnerCount: values.assignedLearnerCount
                };
            });

            if (!hasMatch) {
                next.push({
                    id: normalizedId,
                    title: values.title,
                    code: values.code,
                    provider: normalizedProvider,
                    link: values.link,
                    maxSeats: values.assignedLearnerCount,
                    enrolledCount: values.assignedLearnerCount,
                    assignedLearnerCount: values.assignedLearnerCount,
                    category: 'Others',
                    level: '',
                    fileUrl: ''
                });
                next.sort((left: any, right: any) => (left?.title || '').localeCompare(right?.title || ''));
            }

            return hasChanged(prev, next) ? next : prev;
        });

        setSelectedCert((prev: any) => {
            if (!prev || Number(prev?.id || 0) !== normalizedId) {
                return prev;
            }

            const next = {
                ...prev,
                name: values.title,
                title: values.title,
                code: values.code,
                provider: normalizedProvider,
                link: values.link,
                maxSeats: values.assignedLearnerCount,
                enrolledCount: values.assignedLearnerCount,
                assignedLearnerCount: values.assignedLearnerCount
            };

            return hasChanged(prev, next) ? next : prev;
        });
    };

    const saveNewCertification = async (): Promise<void> => {
        const title = newCertificationTitle.trim();
        const code = newCertificationCode.trim().toUpperCase();
        const provider = normalizeProviderLabel(newCertificationProvider) || 'Other';
        const link = newCertificationLink.trim();

        if (!title) {
            showToast('Certification title is required.', 'error');
            return;
        }

        if (!code) {
            showToast('Certification code is required.', 'error');
            return;
        }

        setIsCreatingCertification(true);
        try {
            const latestCertification = editingCertificationId
                ? await SharePointService.getCertificationDetailsById(editingCertificationId, true)
                : null;
            const assignedLearnerCount = Math.max(Number(
                latestCertification?.assignedLearnerCount ??
                latestCertification?.enrolledCount ??
                latestCertification?.maxSeats ??
                selectedCertificationItem?.assignedLearnerCount ??
                selectedCertificationItem?.enrolledCount ??
                selectedCertificationItem?.maxSeats ??
                0
            ), 0);

            console.log('OLD:', selectedCertificationItem?.provider || '');
            console.log('NEW:', newCertificationProvider);
            console.log('[Certifications] Updating certification draft', {
                id: editingCertificationId,
                title,
                code,
                provider,
                link,
                assignedLearnerCount
            });

            if (editingCertificationId) {
                await SharePointService.updateCertificationItem(editingCertificationId, title, assignedLearnerCount, code, { provider, link });
                syncCertificationCatalogItemLocally(editingCertificationId, {
                    title,
                    code,
                    provider,
                    link,
                    assignedLearnerCount
                });
            } else {
                const createdCertificationId = await SharePointService.createCertificationItem(title, 0, code, { provider, link });
                syncCertificationCatalogItemLocally(createdCertificationId, {
                    title,
                    code,
                    provider,
                    link,
                    assignedLearnerCount: 0
                });
            }

            try {
                await Promise.all([
                    loadCertificationMaxSeatsData(true),
                    refreshEnrollments()
                ]);
            } catch (refreshError) {
                console.warn('[Certifications] Certification save succeeded but the immediate refresh failed.', refreshError);
            }

            showToast(editingCertificationId
                ? 'Certification updated in the SharePoint Certifications list.'
                : 'Certification created in the SharePoint Certifications list.');
            resetCreateCertificationModal();
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : (editingCertificationId ? 'Failed to update certification.' : 'Failed to create certification.');
            console.error('[Certifications] Failed to save certification', {
                title,
                code,
                provider,
                link,
                id: editingCertificationId,
                error
            });
            showToast(errorMessage, 'error');
        } finally {
            setIsCreatingCertification(false);
        }
    };

    const deleteSelectedCertification = async (): Promise<void> => {
        if (!editingCertificationId) {
            return;
        }

        const confirmationAccepted = window.confirm(
            `Delete certification "${newCertificationTitle.trim() || newCertificationCode.trim() || 'this certification'}"?`
        );

        if (!confirmationAccepted) {
            return;
        }

        setIsDeletingCertification(true);
        try {
            await SharePointService.deleteCertificationItem(
                editingCertificationId,
                newCertificationTitle.trim(),
                newCertificationCode.trim()
            );
            await loadCertificationMaxSeatsData(true);
            showToast('Certification deleted from the SharePoint Certifications list.');
            resetCreateCertificationModal();
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : 'Failed to delete certification.';
            console.error('[Certifications] Failed to delete certification', {
                id: editingCertificationId,
                title: newCertificationTitle,
                code: newCertificationCode,
                error
            });
            showToast(errorMessage, 'error');
        } finally {
            setIsDeletingCertification(false);
        }
    };

    const triggerCertificationWorkbookUpload = (): void => {
        if (isUploadingCertificationWorkbook) {
            return;
        }

        certificationWorkbookInputRef.current?.click();
    };

    const handleCertificationWorkbookUpload = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
        const file = event.target.files?.[0];
        event.target.value = '';

        if (!file) {
            return;
        }

        if (!context?.pageContext?.web?.absoluteUrl) {
            showToast('SharePoint context is unavailable. Open the admin portal on the production SharePoint site.', 'error');
            return;
        }

        setIsUploadingCertificationWorkbook(true);
        try {
            const workbookData = await file.arrayBuffer();
            const workbook = XLSX.read(workbookData, { type: 'array' });
            const parsedRows = collectWorkbookRows(workbook);

            if (parsedRows.length === 0) {
                throw new Error('The Excel file does not contain any rows across its sheets.');
            }

            const validRows = parsedRows.filter((parsedRow) => {
                console.log('Mapped Row:', parsedRow);
                const isValid = validateCertificationWorkbookRow(parsedRow);
                if (!isValid) {
                    console.warn('Skipping row due to missing certification name or exam code value:', parsedRow.SourceRow || parsedRow);
                }

                return isValid;
            });

            console.log('Valid Rows:', validRows.length);

            if (validRows.length === 0) {
                throw new Error("The Excel file does not contain usable certification name and exam code values. Expected 'Certification Name'/'Exam Code' headers or the malformed Excel fallback columns.");
            }

            const uploadResult = await SharePointService.bulkUpsertCertificationItems(
                validRows.map((row) => ({
                    title: (row.Title || '').trim(),
                    code: (row.CertificationCode || '').trim().toUpperCase(),
                    maxSeats: 0,
                    provider: normalizeProviderLabel(row.Provider || '') || 'Other'
                }))
            );

            await loadCertificationMaxSeatsData(true);
            showToast(
                `Excel upload complete. Created ${uploadResult.createdCount}, updated ${uploadResult.updatedCount}, skipped ${uploadResult.skippedCount}, processed ${uploadResult.totalProcessed}.`,
                'success'
            );
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : 'Failed to upload certifications Excel file.';
            console.error('[Certifications] Excel upload failed', {
                fileName: file.name,
                error
            });
            showToast(errorMessage, 'error');
        } finally {
            setIsUploadingCertificationWorkbook(false);
        }
    };

    const getAssignableLearners = (users: any[]): any[] => {
        return (users || [])
            .filter((user: any) => !!(user?.email || user?.Email))
            .sort((a: any, b: any) => (a.name || '').localeCompare(b.name || ''));
    };

    const filterAssignableLearners = (users: any[], rawSearch: string): any[] => {
        const normalizedSearch = (rawSearch || '').trim().toLowerCase();
        if (!normalizedSearch) {
            return users;
        }

        return users.filter((user: any) =>
            [
                user.name || user.Title,
                user.email || user.Email,
                user.employeeId,
                user.jobTitle || user.JobTitle
            ].some((value: string) => (value || '').toLowerCase().includes(normalizedSearch))
        );
    };

    useEffect(() => {
        if (!showAssignCertModal) {
            return;
        }

        let isCancelled = false;

        const cachedUsers = directoryUsers.length > 0
            ? directoryUsers
            : directorySyncState.users.length > 0
                ? directorySyncState.users
                : allUsers;
        const cachedLearners = getAssignableLearners(normalizeLearnerSelectionUsers(cachedUsers));

        setAssignModalLearners(cachedLearners);
        setFilteredAssignModalLearners(filterAssignableLearners(cachedLearners, ''));
        setIsLoadingAssignLearners(true);

        const loadAssignableLearners = async (): Promise<void> => {
            try {
                const learners = await SharePointService.getAssessmentAssignmentLearners();
                const normalizedLearners = getAssignableLearners(normalizeLearnerSelectionUsers(learners));

                if (isCancelled) {
                    return;
                }

                setAssignModalLearners(normalizedLearners);
                setFilteredAssignModalLearners(filterAssignableLearners(normalizedLearners, ''));
                setIsLoadingAssignLearners(false);
            } catch (error) {
                console.error('[Certifications] Failed to load direct assignment learners', error);

                if (isCancelled) {
                    return;
                }

                setAssignModalLearners(cachedLearners);
                setFilteredAssignModalLearners(filterAssignableLearners(cachedLearners, ''));
                setIsLoadingAssignLearners(false);
            }
        };

        void loadAssignableLearners();

        return () => {
            isCancelled = true;
        };
    }, [showAssignCertModal, directoryUsers, directorySyncState.users, allUsers]);

    useEffect(() => {
        if (!showAssignCertModal) {
            return;
        }

        setFilteredAssignModalLearners(filterAssignableLearners(assignModalLearners, certUserSearchTerm));
    }, [assignModalLearners, certUserSearchTerm, showAssignCertModal]);

    const selectedLearnersForCert = useMemo(() => {
        const selectedLookup = new Set(selectedUsersForCert);
        return assignModalLearners.filter((user: any, index: number) =>
            selectedLookup.has(getLearnerSelectionId(user, index))
        );
    }, [assignModalLearners, selectedUsersForCert]);

    const toggleCertUserSelection = (id: number) => {
        setSelectedUsersForCert((prev) =>
            prev.indexOf(id) !== -1
                ? prev.filter((selectedId) => selectedId !== id)
                : [...prev, id]
        );
    };

    const clearCertUserSearch = () => {
        setCertUserSearchTerm('');
    };


    return (
        <div className={`admin-portal-wrapper ${isResizing ? 'is-resizing' : ''}`}>
            <input
                ref={certificationWorkbookInputRef}
                type="file"
                accept=".xlsx,.xls"
                style={{ display: 'none' }}
                onChange={(event) => { void handleCertificationWorkbookUpload(event); }}
            />
            <aside className="portal-sidebar" style={{ width: sidebarWidth }}>
                <PortalBrand onFallbackNavigate={handleLogoNavigate} />
                <nav className="nav-list">
                    <div className="nav-section">
                        <div className="nav-header">INSIGHTS</div>
                        <NavBtn active={view === 'DASHBOARD'} onClick={() => setView('DASHBOARD')} icon={<LayoutGrid size={20} />} label="Overview" />
                        <NavBtn active={view === 'REPORTS'} onClick={() => setView('REPORTS')} icon={<BarChart3 size={20} />} label="Detailed Reports" />
                    </div>

                    <div className="nav-section">
                        <div className="nav-header">MANAGEMENT</div>
                        <NavBtn active={view === 'MANAGEMENT'} onClick={() => setView('MANAGEMENT')} icon={<BookOpen size={20} />} label="Certifications" />
                        <NavBtn active={view === 'CONTENT'} onClick={() => setView('CONTENT')} icon={<Upload size={20} />} label="Content Library" />
                        <NavBtn active={view === 'ASSESSMENTS'} onClick={() => setView('ASSESSMENTS')} icon={<FileQuestion size={20} />} label="Assessments" />
                        <NavBtn active={view === 'USERS'} onClick={() => setView('USERS')} icon={<Users size={20} />} label="Learners" />
                        <NavBtn active={view === 'TRACKER'} onClick={() => setView('TRACKER')} icon={<ClipboardCheck size={20} />} label="Enrollments" />
                    </div>

                    <div className="nav-section">
                        <div className="nav-header">SYSTEM</div>
                        <NavBtn active={view === 'AUDIT'} onClick={() => setView('AUDIT')} icon={<Activity size={20} />} label="Audit Logs" />

                        <NavBtn active={view === 'CONFIG'} onClick={() => setView('CONFIG')} icon={<Settings size={20} />} label="Settings" />
                    </div>
                </nav>

                <div className="sidebar-footer-stat">
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '4px' }}>
                        <CheckCircle size={14} /> LIVE SYNC ACTIVE
                    </div>
                    {realEnrollments.length} sessions detected
                </div>
            </aside>

            <div className="sidebar-resizer" onMouseDown={startResizing} />

            <main className="portal-main">
                <header className="portal-header">
                    <div className="header-meta">
                        <div ref={notificationsMenuRef} style={{ position: 'relative' }}>
                            <button className={`icon-button ${showAdminNotifications ? 'active' : ''}`} onClick={() => { setShowAdminNotifications(!showAdminNotifications); setShowProfileOverlay(false); }} title="System Notifications">
                                <Bell size={20} />
                                {adminNotifications.some(n => !n.read) && <span className="notification-dot" style={{ position: 'absolute', top: '8px', right: '8px', width: '8px', height: '8px', background: '#e11d48', borderRadius: '50%', border: '2px solid white' }}></span>}
                            </button>
                            {showAdminNotifications && (
                                <div className="dropdown-panel notifications-dropdown fade-in" style={{ position: 'absolute', top: '100%', right: 0, width: '320px', background: 'white', borderRadius: '20px', boxShadow: '0 20px 50px rgba(0,0,0,0.15)', border: '1.5px solid var(--border)', zIndex: 3000, marginTop: '1rem', padding: '1.5rem' }}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem', borderBottom: '1px solid #f1f5f9', paddingBottom: '1rem' }}>
                                        <h4 style={{ margin: 0, fontWeight: 950, fontSize: '1.1rem' }}>Admin Activity</h4>
                                        <button onClick={handleMarkAdminRead} style={{ background: 'none', border: 'none', color: 'var(--primary)', fontSize: '0.75rem', fontWeight: 800, cursor: 'pointer' }}>Mark all as read</button>
                                    </div>
                                    <div style={{ display: 'grid', gap: '1rem', maxHeight: '350px', overflowY: 'auto' }}>
                                        {adminNotifications.length > 0 ? adminNotifications.map(n => (
                                            <div key={n.id} style={{ display: 'flex', gap: '12px', padding: '0.75rem', borderRadius: '12px', background: !n.read ? '#eff6ff' : '#f8fafc', borderLeft: !n.read ? '4px solid var(--primary)' : 'none' }}>
                                                <div style={{ width: '36px', height: '36px', borderRadius: '10px', background: n.type === 'success' ? '#ecfdf5' : '#eff6ff', display: 'flex', alignItems: 'center', justifyContent: 'center', color: n.type === 'success' ? '#059669' : '#2563eb' }}>
                                                    {n.type === 'success' ? <CheckCircle size={18} /> : <Info size={18} />}
                                                </div>
                                                <div style={{ flex: 1 }}>
                                                    <div style={{ fontSize: '0.85rem', fontWeight: 900, color: '#1e293b' }}>{n.title}</div>
                                                    <div style={{ fontSize: '0.75rem', color: '#64748b', fontWeight: 600, marginTop: '2px' }}>{n.text}</div>
                                                    <div style={{ fontSize: '0.65rem', color: '#94a3b8', fontWeight: 700, marginTop: '4px' }}>{n.time}</div>
                                                </div>
                                            </div>
                                        )) : (
                                            <div style={{ padding: '2rem', textAlign: 'center', color: '#94a3b8', fontWeight: 700 }}>No new messages</div>
                                        )}
                                    </div>
                                </div>
                            )}
                        </div>
                        <button className="icon-button" onClick={() => setView('CONFIG')} title="System Settings"><Settings size={20} /></button>
                        <div ref={profileMenuRef} style={{ position: 'relative' }}>
                            <div className="user-badge" onClick={() => setShowProfileOverlay(!showProfileOverlay)} style={{ cursor: 'pointer', paddingRight: '1.5rem' }}>
                                <div className="user-avatar-small" style={{ width: '36px', height: '36px', borderRadius: '12px', overflow: 'hidden', border: '2px solid white', boxShadow: '0 4px 6px rgba(0,0,0,0.1)' }}>
                                    {userPhotoUrl ? (
                                        <img src={userPhotoUrl} alt="Avatar" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
                                    ) : (
                                        <div style={{ width: '100%', height: '100%', background: 'var(--primary)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'white' }}>
                                            <UserCircle size={24} />
                                        </div>
                                    )}
                                </div>
                                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
                                    <span style={{ fontSize: '0.9rem', fontWeight: 950, color: '#1e293b' }}>{userDisplayName}</span>
                                    <span style={{ fontSize: '0.7rem', fontWeight: 800, color: '#64748b', textTransform: 'uppercase' }}>{props.userRole || (props.isOwner ? 'Owner' : 'Member')}</span>
                                </div>
                            </div>

                            {showProfileOverlay && (
                                <>
                                    <div
                                        className="profile-backdrop"
                                        onClick={() => setShowProfileOverlay(false)}
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
                                                    <div className="name">{userDisplayName}</div>
                                                    <div className="email">{userEmail}</div>
                                                </div>
                                            </div>

                                            <div className="account-card-stats">
                                                <div className="account-card-stat-item">
                                                    <div className="account-label">
                                                        <CheckCircle2 size={18} style={{ color: '#10b981' }} />
                                                        <span>Account Status</span>
                                                    </div>
                                                    <div className="account-value status">Active</div>
                                                </div>
                                                <div className="account-card-stat-item">
                                                    <div className="account-label">
                                                        <Shield size={18} style={{ color: 'var(--primary)' }} />
                                                        <span>Role Level</span>
                                                    </div>
                                                    <div className="account-value">{props.userRole || (props.isOwner ? 'Owner' : 'Member')}</div>
                                                </div>
                                            </div>

                                            <button
                                                className="account-signout-btn"
                                                style={{ background: '#f8fafc', color: '#64748b' }}
                                                onClick={() => { setShowProfileOverlay(false); handleLogoNavigate(); }}
                                            >
                                                <LogOut size={20} />
                                                Return to User Page
                                            </button>
                                        </div>
                                    </div>
                                </>
                            )}
                        </div>
                    </div>
                </header>

                <div className="portal-content-scroll">
                    <LayoutScrollWrapper className="portal-layout-scroll-frame" innerClassName="portal-layout-scroll-frame__inner">
                        {view === 'DASHBOARD' && <DashboardView stats={stats} customCerts={customCerts} setView={setView} />}
                        {view === 'MANAGEMENT' && <ManagementView
                            filteredCerts={filteredCerts}
                            certificationSearchText={certificationSearchText}
                            setCertificationSearchText={setCertificationSearchText}
                            setSelectedCert={setSelectedCert}
                            setView={setView}
                            handleOpenAssignModal={handleOpenAssignModal}
                            openEditModal={openEditModal}
                            openCreateCertificationModal={openCreateCertificationModal}
                            triggerCertificationWorkbookUpload={triggerCertificationWorkbookUpload}
                            isUploadingCertificationWorkbook={isUploadingCertificationWorkbook}
                            pathLibraryState={pathLibraryState}
                            onRefreshPathLibrary={() => { void Promise.all([loadCertificationMaxSeatsData(true), refreshEnrollments()]); }}
                        />}
                        {view === 'ASSESSMENTS' && <AssessmentsView allUsers={allUsers} updateAdminNotifications={updateAdminNotifications} />}
                        {view === 'DETAILS' && <DetailsView selectedCert={selectedCert} setView={setView} />}
                        {view === 'TRACKER' && <EnrollmentTrackerView realEnrollments={realEnrollments} handleDeleteEnrollment={handleDeleteEnrollment} />}
                        {view === 'USERS' && <UsersView
                            allUsers={allUsers}
                            setAllUsers={setAllUsers}
                            taxonomyData={taxonomyData}
                            setShowAddUserModal={setShowAddUserModal}
                            userEmail={userEmail}
                            updateAdminNotifications={updateAdminNotifications}
                            seatManagedCerts={allCerts}
                            context={context}
                            realEnrollments={realEnrollments}
                            onEnrollmentsChanged={refreshEnrollments}
                            onCertificationCountsChanged={loadCertificationMaxSeatsData}
                            directorySyncState={directorySyncState}
                        />}
                        {view === 'CONTENT' && <ContentLibraryView showToast={showToast} userEmail={userEmail} context={props.context} updateAdminNotifications={updateAdminNotifications} />}
                        {view === 'ASSIGNMENTS' && <AssignmentsView taxonomyData={taxonomyData} allUsers={allUsers} seatManagedCerts={allCerts} userEmail={userEmail} updateAdminNotifications={updateAdminNotifications} realEnrollments={realEnrollments} onEnrollmentsChanged={refreshEnrollments} onCertificationCountsChanged={loadCertificationMaxSeatsData} context={context} />}
                        {view === 'TAXONOMY' && <TaxonomyView
                            taxonomyData={taxonomyData}
                            setTaxonomyData={setTaxonomyData}
                            activeTab={activeTaxonomyTab}
                            setActiveTab={setActiveTaxonomyTab}
                            setShowTaxonomyModal={setShowTaxonomyModal}
                        />}
                        {view === 'REPORTS' && <ReportsView realEnrollments={realEnrollments} allUsers={allUsers} />}
                        {view === 'AUDIT' && <AuditView />}
                        {view === 'SECURITY' && <SecurityView
                            accessUsers={accessUsers}
                            setAccessUsers={setAccessUsers}
                            setShowAddUser={setShowAddUser}
                        />}
                        {view === 'CONFIG' && <ConfigView config={config} updateConfig={updateConfig} />}
                    </LayoutScrollWrapper>
                </div>
            </main>

            {
                showCreateCertificationModal && (
                    <div className="modal-overlay" style={{
                        position: 'fixed', inset: 0, backgroundColor: 'rgba(15, 23, 42, 0.4)', backdropFilter: 'blur(12px)', zIndex: 2000,
                        display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '1.5rem'
                    }}>
                        <div className="modal-card fade-in" style={{
                            backgroundColor: 'white', padding: '2.5rem', borderRadius: '32px', width: '100%', maxWidth: '520px',
                            boxShadow: '0 25px 70px -12px rgba(0,0,0,0.3)', border: '1.5px solid var(--border)'
                        }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '2rem' }}>
                                <div>
                                    <h2 style={{ fontSize: '1.85rem', fontWeight: 950, color: '#1e293b', margin: 0, letterSpacing: '-0.04em' }}>
                                        {editingCertificationId ? 'Edit ' : 'Add '}<span style={{ color: 'var(--primary)' }}>Certification</span>
                                    </h2>
                                    <p style={{ color: 'var(--text-muted)', fontSize: '0.95rem', fontWeight: 600, marginTop: '0.4rem' }}>
                                        {editingCertificationId
                                            ? 'Update the selected row in the SharePoint Certifications list.'
                                            : 'Create a new row in the SharePoint Certifications list.'}
                                    </p>
                                </div>
                                <button onClick={closeCreateCertificationModal} className="btn-icon" disabled={isCreatingCertification || isDeletingCertification}><X size={24} /></button>
                            </div>

                            <div style={{ display: 'grid', gap: '1.25rem' }}>
                                <div className="form-group">
                                    <label style={{ display: 'block', fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase' }}>
                                        Certification Title
                                    </label>
                                    <input
                                        type="text"
                                        className="input-field"
                                        value={newCertificationTitle}
                                        onChange={(event: React.ChangeEvent<HTMLInputElement>) => setNewCertificationTitle(event.target.value)}
                                        disabled={isCreatingCertification || isDeletingCertification}
                                    />
                                </div>

                                <div className="form-group">
                                    <label style={{ display: 'block', fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase' }}>
                                        Certification Code
                                    </label>
                                    <input
                                        type="text"
                                        className="input-field"
                                        value={newCertificationCode}
                                        onChange={(event: React.ChangeEvent<HTMLInputElement>) => setNewCertificationCode(event.target.value.toUpperCase())}
                                        disabled={isCreatingCertification || isDeletingCertification}
                                    />
                                </div>

                                <div className="form-group">
                                    <label style={{ display: 'block', fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase' }}>
                                        Provider
                                    </label>
                                    <select
                                        className="input-field"
                                        value={newCertificationProvider}
                                        onChange={(event: React.ChangeEvent<HTMLSelectElement>) => setNewCertificationProvider(event.target.value)}
                                        disabled={isCreatingCertification || isDeletingCertification}
                                    >
                                        <option value="Microsoft">Microsoft</option>
                                        <option value="Google">Google</option>
                                        <option value="AWS">AWS</option>
                                        <option value="Other">Other</option>
                                    </select>
                                </div>

                                <div className="form-group">
                                    <label style={{ display: 'block', fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase' }}>
                                        Certification Link
                                    </label>
                                    <input
                                        type="url"
                                        className="input-field"
                                        placeholder="https://learn.microsoft.com/..."
                                        value={newCertificationLink}
                                        onChange={(event: React.ChangeEvent<HTMLInputElement>) => setNewCertificationLink(event.target.value)}
                                        disabled={isCreatingCertification || isDeletingCertification}
                                    />
                                </div>

                                <div className="seat-capacity-actions" style={{ justifyContent: editingCertificationId ? 'space-between' : 'flex-end' }}>
                                    {editingCertificationId && (
                                        <button
                                            type="button"
                                            className="seat-capacity-btn seat-capacity-btn--delete"
                                            onClick={() => { void deleteSelectedCertification(); }}
                                            disabled={isCreatingCertification || isDeletingCertification}
                                        >
                                            {isDeletingCertification ? 'Deleting...' : 'Delete Certification'}
                                        </button>
                                    )}
                                    <div style={{ display: 'flex', gap: '0.75rem', justifyContent: 'flex-end' }}>
                                    <button
                                        type="button"
                                        className="seat-capacity-btn seat-capacity-btn--cancel"
                                        onClick={closeCreateCertificationModal}
                                        disabled={isCreatingCertification || isDeletingCertification}
                                    >
                                        Cancel
                                    </button>
                                    <button
                                        type="button"
                                        className="seat-capacity-btn seat-capacity-btn--save"
                                        onClick={() => { void saveNewCertification(); }}
                                        disabled={isCreatingCertification || isDeletingCertification}
                                    >
                                        {isCreatingCertification
                                            ? (editingCertificationId ? 'Saving...' : 'Creating...')
                                            : (editingCertificationId ? 'Save Changes' : 'Add Certification')}
                                    </button>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                )
            }

            {/* Add User Modal */}
            {
                showAddUser && (
                    <div className="modal-overlay" style={{
                        position: 'fixed', inset: 0, backgroundColor: 'rgba(15, 23, 42, 0.4)', backdropFilter: 'blur(12px)', zIndex: 2000,
                        display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '1.5rem'
                    }}>
                        <div className="modal-card fade-in" style={{
                            backgroundColor: 'white', padding: '2.5rem', borderRadius: '32px', width: '100%', maxWidth: '480px',
                            boxShadow: '0 25px 70px -12px rgba(0,0,0,0.3)', border: '1.5px solid var(--border)'
                        }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '2rem' }}>
                                <div>
                                    <h2 style={{ fontSize: '1.85rem', fontWeight: 950, color: '#1e293b', margin: 0, letterSpacing: '-0.04em' }}>Authorize <span style={{ color: 'var(--primary)' }}>Member</span></h2>
                                    <p style={{ color: 'var(--text-muted)', fontSize: '0.95rem', fontWeight: 600, marginTop: '0.4rem' }}>Grant administrative portal access.</p>
                                </div>
                                <button onClick={() => setShowAddUser(false)} className="btn-icon"><X size={24} /></button>
                            </div>
                            <div style={{ display: 'grid', gap: '1.75rem' }}>
                                <div className="form-group">
                                    <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase' }}>Full Name</label>
                                    <input type="text" className="input-field" value={newUserData.name} onChange={e => setNewUserData({ ...newUserData, name: e.target.value })} placeholder="e.g. Michael Chen" />
                                </div>
                                <div className="form-group">
                                    <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase' }}>Corporate Email</label>
                                    <input type="email" className="input-field" value={newUserData.email} onChange={e => setNewUserData({ ...newUserData, email: e.target.value })} placeholder="m.chen@skysecure.com" />
                                </div>
                                <div className="form-group">
                                    <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase' }}>Access Role</label>
                                    <select className="input-field" style={{ width: '100%', padding: '0.85rem' }} value={newUserData.role} onChange={e => setNewUserData({ ...newUserData, role: e.target.value })}>
                                        <option value="Member">Member (Standard Access)</option>
                                        <option value="Owner">Owner (Manager Access)</option>
                                    </select>
                                </div>
                                <div className="responsive-two-column-grid">
                                    <div className="form-group">
                                        <label style={{ fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase' }}>Business Unit</label>
                                        <input type="text" className="input-field" value={newUserData.bu} onChange={e => setNewUserData({ ...newUserData, bu: e.target.value })} placeholder="e.g. Sales" />
                                    </div>
                                    <div className="form-group">
                                        <label style={{ fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase' }}>Department</label>
                                        <input type="text" className="input-field" value={newUserData.dept} onChange={e => setNewUserData({ ...newUserData, dept: e.target.value })} placeholder="e.g. Solutions" />
                                    </div>
                                </div>
                                <button className="btn-primary" style={{ width: '100%', marginTop: '1rem', justifyContent: 'center', padding: '1.1rem' }} onClick={() => {
                                    if (!newUserData.name.trim() || !newUserData.email.trim()) {
                                        alert("Please enter the Member's Full Name AND Corporate Email address before granting access.");
                                        return;
                                    }
                                    const updated = [...accessUsers, { ...newUserData, id: Date.now(), status: 'Active' }];
                                    setAccessUsers(updated);
                                    localStorage.setItem('portalAccessUsers', JSON.stringify(updated));
                                    setShowAddUser(false);
                                    setNewUserData({ name: '', email: '', role: 'Member', bu: '', dept: '' });
                                    showToast(`Access granted to ${newUserData.name}`);
                                }}>Grant Portal Access <ShieldCheck size={20} /></button>
                            </div>
                        </div>
                    </div>
                )
            }

            {/* Add LMS Learner Modal */}
            {
                showAddUserModal && (
                    <div className="modal-overlay" style={{
                        position: 'fixed', inset: 0, backgroundColor: 'rgba(15, 23, 42, 0.4)', backdropFilter: 'blur(12px)', zIndex: 2000,
                        display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '1.5rem'
                    }}>
                        <div className="modal-card fade-in" style={{
                            backgroundColor: 'white', padding: '2.5rem', borderRadius: '32px', width: '100%', maxWidth: '600px',
                            boxShadow: '0 25px 70px -12px rgba(0,0,0,0.3)', border: '1.5px solid var(--border)'
                        }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '2rem' }}>
                                <div>
                                    <h2 style={{ fontSize: '1.85rem', fontWeight: 950, color: '#1e293b', margin: 0, letterSpacing: '-0.04em' }}>Register <span style={{ color: 'var(--primary)' }}>New Learner</span></h2>
                                    <p style={{ color: 'var(--text-muted)', fontSize: '0.95rem', fontWeight: 600, marginTop: '0.4rem' }}>Manual entry for the organizational LMS.</p>
                                </div>
                                <button onClick={() => setShowAddUserModal(false)} className="btn-icon"><X size={24} /></button>
                            </div>
                            <div className="responsive-two-column-grid" style={{ gap: '1.5rem' }}>
                                <div className="form-group">
                                    <label style={{ fontSize: '0.8rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Employee ID</label>
                                    <input type="text" className="input-field" value={newLmsUserData.employeeId} onChange={e => setNewLmsUserData({ ...newLmsUserData, employeeId: e.target.value })} placeholder="EMP-000" />
                                </div>
                                <div className="form-group">
                                    <label style={{ fontSize: '0.8rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Full Name</label>
                                    <input type="text" className="input-field" value={newLmsUserData.name} onChange={e => setNewLmsUserData({ ...newLmsUserData, name: e.target.value })} placeholder="John Wick" />
                                </div>
                                <div className="form-group" style={{ gridColumn: 'span 2' }}>
                                    <label style={{ fontSize: '0.8rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>LMS Login Email</label>
                                    <input type="email" className="input-field" value={newLmsUserData.email} onChange={e => setNewLmsUserData({ ...newLmsUserData, email: e.target.value })} placeholder="j.wick@skysecure.ai" />
                                </div>
                                <div className="form-group">
                                    <label style={{ fontSize: '0.8rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Department</label>
                                    <select className="input-field" style={{ width: '100%' }} value={newLmsUserData.department} onChange={e => setNewLmsUserData({ ...newLmsUserData, department: e.target.value })}>
                                        <option value="">Select Dept</option>
                                        {taxonomyData.departments.map((d: string) => <option key={d} value={d}>{d}</option>)}
                                    </select>
                                </div>
                                <div className="form-group">
                                    <label style={{ fontSize: '0.8rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Designation</label>
                                    <select className="input-field" style={{ width: '100%' }} value={newLmsUserData.role} onChange={e => setNewLmsUserData({ ...newLmsUserData, role: e.target.value })}>
                                        <option value="">Select Role</option>
                                        {taxonomyData.roles.map((r: string) => <option key={r} value={r}>{r}</option>)}
                                    </select>
                                </div>
                                <div className="form-group">
                                    <label style={{ fontSize: '0.8rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Business Unit</label>
                                    <select className="input-field" style={{ width: '100%' }} value={newLmsUserData.businessUnit} onChange={e => setNewLmsUserData({ ...newLmsUserData, businessUnit: e.target.value })}>
                                        <option value="">Select BU</option>
                                        {taxonomyData.businessUnits.map((b: string) => <option key={b} value={b}>{b}</option>)}
                                    </select>
                                </div>
                                <div className="form-group">
                                    <label style={{ fontSize: '0.8rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Base Location</label>
                                    <select className="input-field" style={{ width: '100%' }} value={newLmsUserData.location} onChange={e => setNewLmsUserData({ ...newLmsUserData, location: e.target.value })}>
                                        <option value="">Select Location</option>
                                        {taxonomyData.locations.map((l: string) => <option key={l} value={l}>{l}</option>)}
                                    </select>
                                </div>
                                <button className="btn-primary" style={{ gridColumn: 'span 2', marginTop: '1rem', justifyContent: 'center', padding: '1.1rem' }} onClick={() => {
                                    if (!newLmsUserData.name || !newLmsUserData.email || !newLmsUserData.employeeId) {
                                        alert("Please fill name, email and employee ID.");
                                        return;
                                    }
                                    const updated = [...allUsers, { ...newLmsUserData, id: `LMS_${Date.now()}`, progress: 0 }];
                                    setAllUsers(updated);
                                    const audit = JSON.parse(localStorage.getItem('lmsAuditLogs') || '[]');
                                    audit.unshift({ id: Date.now(), user: props.userEmail || 'Admin', action: 'CREATE', detail: `Registered new learner ${newLmsUserData.email}`, timestamp: new Date().toISOString() });
                                    localStorage.setItem('lmsAuditLogs', JSON.stringify(audit.slice(0, 50)));

                                    setShowAddUserModal(false);
                                    setNewLmsUserData({ employeeId: '', name: '', email: '', department: '', businessUnit: '', role: '', location: '', status: 'Active' });
                                    showToast(`${newLmsUserData.name} added to the system.`);
                                }}>Create User Account <Plus size={20} /></button>
                            </div>
                        </div>
                    </div>
                )
            }

            {/* Add Taxonomy Modal */}
            {
                showTaxonomyModal && (
                    <div className="modal-overlay" style={{
                        position: 'fixed', inset: 0, backgroundColor: 'rgba(15, 23, 42, 0.4)', backdropFilter: 'blur(12px)', zIndex: 2000,
                        display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '1.5rem'
                    }}>
                        <div className="modal-card fade-in" style={{
                            backgroundColor: 'white', padding: '2.5rem', borderRadius: '32px', width: '100%', maxWidth: '420px',
                            boxShadow: '0 25px 70px -12px rgba(0,0,0,0.3)', border: '1.5px solid var(--border)'
                        }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '2rem' }}>
                                <div>
                                    <h2 style={{ fontSize: '1.85rem', fontWeight: 950, color: '#1e293b', margin: 0, letterSpacing: '-0.04em' }}>Add <span style={{ color: 'var(--primary)' }}>{activeTaxonomyTab.replace(/([A-Z])/g, ' $1')}</span></h2>
                                    <p style={{ color: 'var(--text-muted)', fontSize: '0.95rem', fontWeight: 600, marginTop: '0.4rem' }}>Extend the platform metadata.</p>
                                </div>
                                <button onClick={() => setShowTaxonomyModal(false)} className="btn-icon"><X size={24} /></button>
                            </div>
                            <div style={{ display: 'grid', gap: '1.75rem' }}>
                                <div className="form-group">
                                    <label style={{ fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase' }}>Entry Name</label>
                                    <input type="text" className="input-field" value={newTaxonomyItemData} onChange={e => setNewTaxonomyItemData(e.target.value)} placeholder={`e.g. New ${activeTaxonomyTab.slice(0, -1)}`} />
                                </div>
                                <button className="btn-primary" style={{ width: '100%', marginTop: '1rem', justifyContent: 'center', padding: '1.1rem' }} onClick={async () => {
                                    if (!newTaxonomyItemData) return;
                                    const updatedItems = [...(taxonomyData[activeTaxonomyTab] || []), newTaxonomyItemData];
                                    const updated = {
                                        ...taxonomyData,
                                        [activeTaxonomyTab]: updatedItems
                                    };
                                    setTaxonomyData(updated);
                                    localStorage.setItem('lmsTaxonomyData', JSON.stringify(updated));
                                    
                                    // Sync to SharePoint
                                    try {
                                        await SharePointService.updateTaxonomy(activeTaxonomyTab, updatedItems);
                                    } catch (e) {
                                        console.warn("Taxonomy sync failed", e);
                                    }

                                    setShowTaxonomyModal(false);
                                    setNewTaxonomyItemData('');
                                    showToast(`${newTaxonomyItemData} added to ${activeTaxonomyTab}`);
                                }}>Confirm Addition <Plus size={20} /></button>
                            </div>
                        </div>
                    </div>
                )
            }
            {
                showAssignCertModal && (
                    <div className="modal-overlay" style={{
                        position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
                        background: 'rgba(15, 23, 42, 0.4)', backdropFilter: 'blur(8px)',
                        display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000
                    }}>
                        <div className="fade-in direct-learner-modal">
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '2rem' }}>
                                <div>
                                    <h2 style={{ fontSize: '1.85rem', fontWeight: 950, color: '#1e293b', margin: 0, letterSpacing: '-0.04em' }}>Direct Learner</h2>
                                    <p style={{ color: 'var(--text-muted)', fontSize: '0.95rem', fontWeight: 600, marginTop: '0.4rem' }}>Enroll a user into {certToAssign?.code}</p>
                                </div>
                                <button onClick={() => setShowAssignCertModal(false)} className="btn-icon"><X size={24} /></button>
                            </div>
                            <div className="direct-learner-modal-body">
                                <div className="form-group">
                                    <label style={{ fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase' }}>Search Learners</label>
                                    <div className="direct-learner-search">
                                        <Search size={18} />
                                        <input
                                            className="input-field direct-learner-search-input"
                                            value={certUserSearchTerm}
                                            onChange={e => setCertUserSearchTerm(e.target.value)}
                                            placeholder="Search by name, email, designation, or employee ID"
                                        />
                                    </div>
                                    <div className="direct-learner-search-meta">
                                        <span>{filteredAssignModalLearners.length} result{filteredAssignModalLearners.length === 1 ? '' : 's'}</span>
                                        {certUserSearchTerm && (
                                            <button
                                                type="button"
                                                className="direct-learner-clear"
                                                onClick={clearCertUserSearch}
                                            >
                                                Clear search
                                            </button>
                                        )}
                                    </div>
                                    {isLoadingAssignLearners && (
                                        <div className="direct-learner-loading">
                                            <Loader2 size={18} className="direct-learner-spinner" />
                                            <span>Loading learners...</span>
                                        </div>
                                    )}
                                </div>
                                <div className="form-group">
                                    <label style={{ fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase' }}>
                                        Scheduled Exam Date
                                    </label>
                                    <input
                                        type="date"
                                        className="input-field"
                                        value={certExamScheduledDate}
                                        min={new Date().toISOString().split('T')[0]}
                                        onChange={(event) => setCertExamScheduledDate(event.target.value)}
                                    />
                                </div>
                                <div>
                                    <div className="direct-learner-section-header">
                                        <label style={{ fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase', display: 'block' }}>
                                            Select Learners
                                        </label>
                                        <div className="selected-count">
                                            {selectedUsersForCert.length} learner{selectedUsersForCert.length === 1 ? '' : 's'} selected
                                        </div>
                                    </div>
                                    <div className="learner-list">
                                        {isLoadingAssignLearners ? (
                                            <div className="direct-learner-empty">
                                                Loading learners...
                                            </div>
                                        ) : filteredAssignModalLearners.length === 0 ? (
                                            <div className="direct-learner-empty">
                                                No matching learners found
                                            </div>
                                        ) : filteredAssignModalLearners.map((user: any) => {
                                            const learnerId = getLearnerSelectionId(user);
                                            const isSelected = selectedUsersForCert.indexOf(learnerId) !== -1;

                                            return (
                                                <div
                                                    key={learnerId}
                                                    className={`learner-card ${isSelected ? 'selected' : ''}`}
                                                    onClick={() => toggleCertUserSelection(learnerId)}
                                                    onKeyDown={(event) => {
                                                        if (event.key === 'Enter' || event.key === ' ') {
                                                            event.preventDefault();
                                                            toggleCertUserSelection(learnerId);
                                                        }
                                                    }}
                                                    role="button"
                                                    tabIndex={0}
                                                    aria-pressed={isSelected}
                                                >
                                                    <div className="learner-card-header">
                                                        <input
                                                            type="checkbox"
                                                            checked={isSelected}
                                                            onChange={() => toggleCertUserSelection(learnerId)}
                                                            onClick={(event) => event.stopPropagation()}
                                                            className="learner-card-checkbox"
                                                        />
                                                        <div className="learner-card-info">
                                                            <strong>{user.name}</strong>
                                                            <p>{user.email}</p>
                                                            <span className="learner-card-designation">{user.jobTitle || 'No designation'}</span>
                                                            <span className="learner-card-meta">{user.employeeId || 'No Employee ID'} | {user.department || 'Not Available'}</span>
                                                        </div>
                                                        {isSelected && (
                                                            <div className="learner-card-checkmark">
                                                                <CheckCircle2 size={18} />
                                                            </div>
                                                        )}
                                                    </div>
                                                </div>
                                            );
                                        })}
                                    </div>
                                    {selectedLearnersForCert.length > 0 && (
                                        <div className="selected-learner-chips">
                                            {selectedLearnersForCert.map((user: any) => (
                                                <span key={user.email} className="selected-learner-chip">
                                                    {user.name}
                                                </span>
                                            ))}
                                        </div>
                                    )}
                                </div>
                                <button 
                                    className="btn-primary push-cert-btn"
                                    style={{ marginTop: '1rem' }} 
                                    onClick={handleAssignCert}
                                    disabled={
                                        selectedUsersForCert.length === 0 ||
                                        isLoadingAssignLearners ||
                                        isAssigningCert
                                    }
                                >
                                    {isAssigningCert
                                        ? `Pushing Certification (${selectedUsersForCert.length})...`
                                        : `Push Certification (${selectedUsersForCert.length})`} <Send size={20} />
                                </button>
                            </div>
                        </div>
                    </div>
                )
            }
            {
                toast && (
                    <div className="toast-notification">
                        <CheckCircle size={20} color="var(--primary)" />
                        <span>{toast.message}</span>
                    </div>
                )
            }
        </div >
    );
}

// --- Sub-View Components (Moved outside for stability) ---

function DashboardView({ stats, customCerts, setView }: any) {
    return (
        <div className="fade-in">
            <div className="view-header">
                <h1 className="view-title">Admin Insights</h1>
                <div className="header-actions">
                    <button className="btn-primary" onClick={() => setView('MANAGEMENT')}><PlusCircle size={18} /> Manage Certs</button>
                </div>
            </div>

            <div className="stats-grid">
                <StatCard icon={<ShieldCheck size={28} />} title="Certifications" value={stats.totalCerts} trend="+2 new" color="#4f46e5" />
                <StatCard icon={<Users size={28} />} title="Enrolled" value={stats.totalEnrolled} trend="+12% YoY" color="#0ea5e9" />
                <StatCard icon={<TrendingUp size={28} />} title="In Progress" value={stats.inProgress} sub={`${stats.completed} completed`} color="#10b981" />
            </div>

            <div className="main-display-grid">
                <div className="activity-feed">
                    <h3 style={{ fontSize: '1.2rem', fontWeight: 900, marginBottom: '2rem' }}>User Feed</h3>
                    <div style={{ display: 'flex', flexDirection: 'column' }}>
                        {[...(customCerts || [])].reverse().slice(0, 5).map(c => (
                            <div key={c.id} className="feed-item">
                                <div className="user-avatar" style={{ backgroundColor: '#4f46e5' }}>{(c.provider || 'C').substring(0, 1)}</div>
                                <div className="feed-info">
                                    <p><strong>{c.userName || 'Member'}</strong> added entry: {c.name}</p>
                                    <span>{c.provider || 'Custom'} â€¢ {c.dateAdded || 'Recently'}</span>
                                </div>
                            </div>
                        ))}
                        {(!customCerts || customCerts.length === 0) && <div style={{ padding: '4rem 1rem', textAlign: 'center' }}><p style={{ color: '#94a3b8', fontSize: '0.95rem', fontWeight: 700 }}>No recent activities detected</p></div>}
                    </div>
                </div>
            </div>
        </div>
    );
}

function ManagementView({ filteredCerts, certificationSearchText, setCertificationSearchText, setSelectedCert, setView, handleOpenAssignModal, openEditModal, openCreateCertificationModal, triggerCertificationWorkbookUpload, isUploadingCertificationWorkbook, pathLibraryState, onRefreshPathLibrary }: any) {
    const groupedCerts = (filteredCerts || []).reduce((acc: Record<string, any[]>, cert: any) => {
        const category = (cert.category || 'Others').toString().trim() || 'Others';
        if (!acc[category]) {
            acc[category] = [];
        }
        acc[category].push(cert);
        return acc;
    }, {});

    const orderedCategories = Object.keys(groupedCerts).sort((a, b) => a.localeCompare(b));

    return (
        <div className="fade-in">
            <div className="view-header cert-management-header">
                <div className="cert-management-title">
                    <h1 className="view-title">Cert Management</h1>
                    <p style={{ color: 'var(--text-muted)', fontWeight: 600 }}>
                        SharePoint Certifications is the single source of truth. Upload one Excel workbook to import certifications from all sheets into SharePoint.
                    </p>
                    <p style={{ color: '#64748b', fontWeight: 700, fontSize: '0.85rem', marginTop: '0.5rem' }}>
                        Total Certifications: {(filteredCerts || []).length}
                    </p>
                    <div className="search-box-unified" style={{ marginTop: '1rem', maxWidth: '460px' }}>
                        <Search size={18} />
                        <input
                            type="text"
                            placeholder="Search certifications by title, code, or provider..."
                            value={certificationSearchText}
                            onChange={(event) => setCertificationSearchText(event.target.value)}
                        />
                    </div>
                </div>
                <div className="header-actions cert-management-actions">
                    <button className="btn-secondary cert-management-action-button" onClick={triggerCertificationWorkbookUpload} disabled={isUploadingCertificationWorkbook}>
                        <Upload size={18} /> {isUploadingCertificationWorkbook ? 'Uploading Excel...' : 'Upload Excel'}
                    </button>
                    <button className="btn-primary cert-management-action-button" onClick={openCreateCertificationModal}>
                        <PlusCircle size={18} /> Add Certification
                    </button>
                    <button className="btn-secondary cert-management-action-button" onClick={onRefreshPathLibrary} disabled={pathLibraryState.refreshing}>
                        {pathLibraryState.refreshing ? 'Refreshing...' : 'Refresh SharePoint'}
                    </button>
                </div>
            </div>

            {pathLibraryState.error && (
                <div style={{ marginBottom: '1rem', padding: '0.9rem 1rem', borderRadius: '14px', border: '1px solid #fecaca', background: '#fef2f2', color: '#b91c1c', fontWeight: 700 }}>
                    {pathLibraryState.error}
                </div>
            )}

            {pathLibraryState.loading && (
                <div style={{ marginBottom: '1rem', padding: '0.9rem 1rem', borderRadius: '14px', border: '1px solid #bfdbfe', background: '#eff6ff', color: '#1d4ed8', fontWeight: 700 }}>
                    Loading the latest certification data from SharePoint.
                </div>
            )}

            {!pathLibraryState.loading && pathLibraryState.refreshing && (
                <div style={{ marginBottom: '1rem', padding: '0.9rem 1rem', borderRadius: '14px', border: '1px solid #bfdbfe', background: '#eff6ff', color: '#1d4ed8', fontWeight: 700 }}>
                    Refreshing the latest certification changes from SharePoint without clearing the table.
                </div>
            )}

            {orderedCategories.map((category) => (
                <div key={category} className="table-container" style={{ marginBottom: '1.5rem' }}>
                    <div style={{ padding: '1rem 1.25rem', borderBottom: '1px solid #e2e8f0', background: '#f8fafc', fontWeight: 900, color: '#1e293b', fontSize: '1rem' }}>
                        {category}
                    </div>
                    <table className="admin-table">
                        <thead>
                            <tr>
                                <th>Certification Details</th>
                                <th>Status</th>
                                <th>Assigned Learners</th>
                                <th style={{ textAlign: 'right' }}>Management</th>
                            </tr>
                        </thead>
                        <tbody>
                            {groupedCerts[category].map((cert: any) => {
                                const canEditCapacity = !!(cert.code || cert.name);

                                return (
                                    <tr key={cert.id}>
                                        <td>
                                            <div className="cert-info">
                                                <div className="cert-name">{cert.name}</div>
                                                <div className="cert-code">{cert.code || 'NO-CODE'}</div>
                                            </div>
                                        </td>
                                        <td><StatusBadge status={cert.status} /></td>
                                        <td>
                                            <div className="seat-capacity-display">
                                                <div className="seat-capacity-value">{cert.assignedLearnerCount ?? cert.enrolledCount ?? cert.occupiedSeats ?? 0}</div>
                                                <div className="seat-capacity-meta">
                                                    Live assigned learner count synced from SharePoint.
                                                </div>
                                            </div>
                                        </td>
                                        <td style={{ textAlign: 'right' }}>
                                            <div className="action-btns">
                                                <button title="Assign certification" className="btn-icon" onClick={() => handleOpenAssignModal(cert)}><Send size={18} /></button>
                                                <button title="Manage Enrollments" className="btn-icon" onClick={() => { setSelectedCert(cert); setView('DETAILS'); }}><Users size={18} /></button>
                                                {cert.link && (
                                                    <a
                                                        href={cert.link}
                                                        target="_blank"
                                                        rel="noopener noreferrer"
                                                        className="cert-edit-btn cert-view-btn"
                                                        title="View Certification"
                                                    >
                                                        <Globe size={16} />
                                                        <span>View Certification</span>
                                                    </a>
                                                )}
                                                {canEditCapacity && (
                                                    <button title="Edit Certification" className="cert-edit-btn" onClick={() => openEditModal(cert)}>
                                                        <Edit size={16} />
                                                        <span>Edit</span>
                                                    </button>
                                                )}
                                            </div>
                                        </td>
                                    </tr>
                                );
                            })}
                        </tbody>
                    </table>
                </div>
            ))}

            {orderedCategories.length === 0 && !pathLibraryState.loading && (
                <div className="table-container" style={{ padding: '4rem', textAlign: 'center', color: '#94a3b8', fontWeight: 700 }}>
                    No certifications found in the SharePoint Certifications list.
                </div>
            )}
        </div>
    );
}

function DetailsView({ selectedCert, setView }: any) {
    return (
        <div className="fade-in">
            <div className="view-header">
                <div>
                    <button className="btn-text" style={{ marginBottom: '1rem', paddingLeft: 0 }} onClick={() => setView('MANAGEMENT')}>â† Back to Library</button>
                    <h1 className="view-title">{selectedCert?.name}</h1>
                </div>
                <div className="header-actions"><StatusBadge status={selectedCert?.status} /></div>
            </div>
            <div className="table-container">
                <table className="admin-table">
                    <thead>
                        <tr><th>Candidate</th><th>Domain</th><th>Enrollment Date</th><th>Verification</th><th style={{ textAlign: 'right' }}>Actions</th></tr>
                    </thead>
                    <tbody>
                        {(selectedCert?.enrollments || []).map((u: any) => (
                            <tr key={u.id}>
                                <td><div className="user-info"><div className="user-name">{u.name || 'Anonymous User'}</div><div className="user-email">{u.email}</div></div></td>
                                <td><span className="pill approved">{u.category || 'N/A'}</span></td>
                                <td>{u.startDate}</td>
                                <td><StatusBadge status="OPEN" /></td>
                                <td style={{ textAlign: 'right' }}><button className="btn-text reject">Revoke</button><button className="btn-text approve">Verify</button></td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        </div>
    );
}

function EnrollmentTrackerView({ realEnrollments, handleDeleteEnrollment }: any) {
    const [showOverdueOnly, setShowOverdueOnly] = useState(false);

    const isOverdue = (endDate: string) => {
        if (!endDate) return false;
        return new Date(endDate) < new Date();
    };

    const filteredEnrollments = showOverdueOnly ? realEnrollments.filter((e: any) => isOverdue(e.endDate) && e.status !== 'completed') : realEnrollments;

    return (
        <div className="fade-in">
            <div className="view-header">
                <div>
                    <h1 className="view-title">Enrollment Tracker</h1>
                    <p style={{ color: 'var(--text-muted)', fontWeight: 600 }}>Real-time monitoring of all learning journeys.</p>
                </div>
                <div style={{ display: 'flex', gap: '0.75rem' }}>
                    <button className="btn-secondary" onClick={() => alert("Downloading PDF Enrollment Report...")} style={{ fontSize: '0.8rem' }}>Export PDF</button>
                    <button className="btn-secondary" onClick={() => setShowOverdueOnly(!showOverdueOnly)} style={{ fontSize: '0.8rem', background: showOverdueOnly ? 'var(--primary)' : 'white', color: showOverdueOnly ? 'white' : 'var(--text-primary)' }}>{showOverdueOnly ? 'Show All' : 'Filter Overdue'}</button>
                </div>
            </div>
            <div className="table-container">
                <table className="admin-table">
                    <thead>
                        <tr>
                            <th>Candidate</th>
                            <th>Certification Path</th>
                            <th>Start Date</th>
                            <th>Target Date</th>
                            <th>Compliance</th>
                            <th style={{ textAlign: 'center' }}>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
                        {(filteredEnrollments || []).length > 0 ? (filteredEnrollments || []).map((e: any, idx: number) => {
                            const overdue = isOverdue(e.endDate) && e.status !== 'completed';
                            return (
                                <tr key={idx}>
                                    <td>
                                        <div className="user-name">{e.userName || 'Unknown User'}</div>
                                        <div className="user-email">{e.email}</div>
                                    </td>
                                    <td style={{ fontWeight: 800 }}>
                                        {e.name}
                                        <div style={{ fontSize: '0.7rem', color: '#64748b', fontWeight: 600 }}>{e.code}</div>
                                    </td>
                                    <td>{e.startDate}</td>
                                    <td style={{ color: overdue ? '#ef4444' : 'inherit', fontWeight: overdue ? 800 : 500 }}>
                                        {e.endDate || 'No Target'}
                                    </td>
                                    <td>
                                        {overdue ? (
                                            <span className="pill" style={{ backgroundColor: '#fff1f2', color: '#e11d48', border: '1px solid #fee2e2' }}>OVERDUE</span>
                                        ) : e.status === 'completed' ? (
                                            <span className="pill approved">COMPLIANT</span>
                                        ) : (
                                            <span className="pill" style={{ backgroundColor: '#eff6ff', color: '#1d4ed8' }}>ON TRACK</span>
                                        )}
                                    </td>
                                    <td style={{ textAlign: 'center' }}>
                                        <button className="btn-icon" style={{ color: '#ef4444', marginLeft: 'auto', marginRight: 'auto' }} onClick={() => handleDeleteEnrollment(e.id)} title="Revoke Enrollment"><Trash2 size={16} /></button>
                                    </td>
                                </tr>
                            );
                        }) : (
                            <tr>
                                <td colSpan={6} style={{ textAlign: 'center', padding: '4rem', color: '#64748b' }}>
                                    <AlertTriangle size={32} style={{ marginBottom: '1rem', opacity: 0.5 }} />
                                    <div style={{ fontWeight: 800 }}>No Active Enrollments Detected</div>
                                    <div style={{ fontSize: '0.85rem' }}>User activities will appear here in real-time.</div>
                                </td>
                            </tr>
                        )}
                    </tbody>
                </table>
            </div>
        </div>
    );
}

function SecurityView({ accessUsers, setAccessUsers, setShowAddUser }: any) {
    return (
        <div className="fade-in">
            <header className="view-header">
                <div><h1 className="view-title">Access Control</h1><p style={{ color: 'var(--text-muted)', fontWeight: 600 }}>Manage portal permissions.</p></div>
                <button className="btn-primary" onClick={() => setShowAddUser(true)}><PlusCircle size={18} /> Add Portal Member</button>
            </header>
            <div className="table-container" style={{ marginTop: '2rem' }}>
                <table className="admin-table">
                    <thead><tr><th>Member Profile</th><th>Corporate Identity</th><th>Org Unit</th><th>Access Level</th><th>Status</th><th>Operations</th></tr></thead>
                    <tbody>
                        {(accessUsers || []).map((user: any) => (
                            <tr key={user.id}>
                                <td><div className="user-info-cell"><div className="user-avatar-small"><UserCircle size={20} /></div><span style={{ fontWeight: 850 }}>{user.name}</span></div></td>
                                <td style={{ color: '#64748b' }}>{user.email}</td>
                                <td>
                                    <div style={{ fontSize: '0.85rem', fontWeight: 700 }}>{user.bu || 'Corporate'}</div>
                                    <div style={{ fontSize: '0.75rem', color: '#94a3b8' }}>{user.dept || 'General'}</div>
                                </td>
                                <td><span className={`pill ${user.role?.toLowerCase()}`}>{user.role}</span></td>
                                <td><span className="status-badge active">{user.status}</span></td>
                                <td><button className="icon-button" style={{ color: '#ef4444' }} onClick={() => setAccessUsers(accessUsers.filter((u: any) => u.id !== user.id))}><Trash2 size={18} /></button></td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        </div>
    );
}

function ConfigView({ config, updateConfig }: any) {
    const handleFactoryReset = () => {
        if (window.confirm('WARNING: This will clear all local enrollments and customizations. Continue?')) {
            localStorage.clear();
            window.location.reload();
        }
    };

    return (
        <div className="fade-in">
            <div className="view-header">
                <div>
                    <h1 className="view-title">System Configurations</h1>
                    <p style={{ color: 'var(--text-muted)', fontWeight: 600 }}>Adjust administrative environment parameters.</p>
                </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(400px, 1fr))', gap: '2rem' }}>
                {/* Engine Settings */}
                <div className="chart-container" style={{ padding: '2.5rem' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '2.5rem' }}>
                        <Settings className="text-gradient" size={24} />
                        <h3 style={{ margin: 0, fontWeight: 950 }}>Core Engine</h3>
                    </div>

                    <div style={{ display: 'grid', gap: '1.25rem' }}>
                        <ToggleItem
                            label="Real-time LocalSync"
                            sub="Sync data across browser tabs instantly"
                            active={true}
                            disabled
                        />
                        <ToggleItem
                            label="Auto-Archive Operations"
                            sub="Archive enrollments older than 90 days"
                            active={config.autoArchive}
                            onToggle={() => updateConfig({ autoArchive: !config.autoArchive })}
                        />
                        <ToggleItem
                            label="Maintenance Mode"
                            sub="Disable public access to Learning Center during updates"
                            active={config.maintenanceMode}
                            onToggle={() => updateConfig({ maintenanceMode: !config.maintenanceMode })}
                        />
                    </div>
                </div>

                {/* Appearance & UX */}
                <div className="chart-container" style={{ padding: '2.5rem' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '2.5rem' }}>
                        <LayoutGrid className="text-gradient" size={24} />
                        <h3 style={{ margin: 0, fontWeight: 950 }}>Interface & UX</h3>
                    </div>

                    <div style={{ display: 'grid', gap: '1.25rem' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '1rem', background: '#f8fafc', borderRadius: '20px', border: '1px solid var(--border)' }}>
                            <div>
                                <div style={{ fontWeight: 850, fontSize: '0.95rem' }}>System Accent Color</div>
                                <div style={{ fontSize: '0.8rem', color: '#64748b', fontWeight: 600 }}>Choose primary brand color</div>
                            </div>
                            <div style={{ display: 'flex', gap: '0.5rem' }}>
                                <button
                                    onClick={() => updateConfig({ accentColor: 'blue' })}
                                    style={{ width: 24, height: 24, borderRadius: '50%', background: '#0ea5e9', border: config.accentColor === 'blue' ? '3px solid white' : 'none', boxShadow: '0 0 0 2px #0ea5e9', cursor: 'pointer' }}
                                />
                                <button
                                    onClick={() => updateConfig({ accentColor: 'green' })}
                                    style={{ width: 24, height: 24, borderRadius: '50%', background: '#10b981', border: config.accentColor === 'green' ? '3px solid white' : 'none', boxShadow: '0 0 0 2px #10b981', cursor: 'pointer' }}
                                />
                            </div>
                        </div>

                        <ToggleItem
                            label="Enhanced Glassmorphism"
                            sub="Apply Gaussian blur to all dashboard panels"
                            active={true}
                            disabled
                        />
                    </div>

                    <div style={{ marginTop: '2.5rem' }}>
                        <h4 style={{ fontSize: '0.8rem', color: '#ef4444', fontWeight: 900, marginBottom: '1rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Criticial Danger Zone</h4>
                        <button
                            className="btn-secondary"
                            style={{ color: '#ef4444', borderColor: '#fee2e2', background: '#fff1f2', width: '100%', justifyContent: 'center' }}
                            onClick={handleFactoryReset}
                        >
                            <Trash2 size={18} /> Perform Factory System Reset
                        </button>
                    </div>
                </div>
            </div>
        </div>
    );
}

function ToggleItem({ label, sub, active, onToggle, disabled }: any) {
    return (
        <div style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            padding: '1rem 1.25rem',
            background: active ? 'rgba(14, 165, 233, 0.04)' : '#f8fafc',
            borderRadius: '20px',
            border: '1.5px solid',
            borderColor: active ? 'rgba(14, 165, 233, 0.2)' : 'var(--border)',
            opacity: disabled ? 0.7 : 1
        }}>
            <div>
                <div style={{ fontWeight: 850, fontSize: '0.95rem', color: active ? 'var(--primary)' : '#1e293b' }}>{label}</div>
                <div style={{ fontSize: '0.8rem', color: '#64748b', fontWeight: 600 }}>{sub}</div>
            </div>
            <div
                onClick={!disabled ? onToggle : undefined}
                style={{
                    width: 48,
                    height: 24,
                    background: active ? 'var(--gradient-primary)' : '#cbd5e1',
                    borderRadius: 24,
                    position: 'relative',
                    cursor: disabled ? 'default' : 'pointer',
                    transition: 'all 0.3s'
                }}
            >
                <div style={{
                    position: 'absolute',
                    top: 4,
                    left: active ? 28 : 4,
                    width: 16,
                    height: 16,
                    background: 'white',
                    borderRadius: '50%',
                    transition: 'all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275)',
                    boxShadow: '0 2px 4px rgba(0,0,0,0.2)'
                }} />
            </div>
        </div>
    );
}

// --- Basic Helper Components ---

function NavBtn({ active, icon, label, onClick }: any) {
    return <button className={`nav-btn ${active ? 'active' : ''}`} onClick={onClick}>{icon} {label}</button>;
}

function StatCard({ icon, title, value, trend, sub, color }: any) {
    return (
        <div className="stat-card">
            <div className="stat-icon" style={{ backgroundColor: `${color}20`, color: color }}>{icon}</div>
            <div className="stat-content">
                <span className="stat-label">{title}</span>
                <div className="stat-value">{value}</div>
                {trend && <span className="stat-trend">{trend}</span>}
                {sub && <span className="stat-sub">{sub}</span>}
            </div>
        </div>
    );
}

function AssessmentsView({ allUsers, updateAdminNotifications }: { allUsers: any[]; updateAdminNotifications?: (notification: any) => Promise<void> | void }) {
    const defaultAssessments = [
        { id: 1, title: 'Identity Fundamentals Quiz', path: 'AZ-900', questions: 10, threshold: 80, isPublished: true },
        { id: 2, title: 'Teams Admin Final Exam', path: 'MS-700', questions: 25, threshold: 75, isPublished: true },
        { id: 3, title: 'Draft: New Security Model', path: 'SC-900', questions: null, threshold: 85, isPublished: false }
    ];

    const [creating, setCreating] = useState(false);
    const [subTab, setSubTab] = useState('results'); // 'management' | 'results'
    const [assessmentResults, setAssessmentResults] = useState<IAssessmentTrackerItem[]>([]);
    const [adminAssessments, setAdminAssessments] = useState<any[]>([]);
    const [searchTerm, setSearchTerm] = useState('');
    const [assessmentResultsLoading, setAssessmentResultsLoading] = useState(false);
    const [assessmentResultsError, setAssessmentResultsError] = useState<string | null>(null);
    const [questions, setQuestions] = useState<any[]>([]);
    const [isGenerating, setIsGenerating] = useState(false);
    const [numToGenerate, setNumToGenerate] = useState(5);
    const [pushingAsmt, setPushingAsmt] = useState<any>(null);
    const [showPushModal, setShowPushModal] = useState(false);
    const [isAssigningAssessment, setIsAssigningAssessment] = useState(false);
    const [selectedUsers, setSelectedUsers] = useState<any[]>([]);
    const [assessmentUserSearchTerm, setAssessmentUserSearchTerm] = useState('');
    const [assessmentScheduledDate, setAssessmentScheduledDate] = useState(new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]);
    const [assessmentLearnerState, setAssessmentLearnerState] = useState<{ users: any[]; loading: boolean }>({
        users: [],
        loading: true
    });

    const autoGenerateLibrary: any = {
        'SC-300': [
            { id: 1, q: "Which tool is used to monitor Azure AD sign-in activity and identify risky sign-ins?", options: ["Azure Monitor", "Azure AD Identity Protection", "Microsoft Sentinel", "Microsoft Defender for Cloud"], correct: 1 },
            { id: 2, q: "What is the primary purpose of Azure AD Privileged Identity Management (PIM)?", options: ["To manage user passwords", "To provide just-in-time privileged access", "To synchronize local AD with cloud", "To create guest user accounts"], correct: 1 },
            { id: 3, q: "Which authentication method provides the highest level of security in Azure AD?", options: ["Password only", "SMS-based MFA", "Certificate-based authentication", "FIDO2 security keys"], correct: 3 },
            { id: 4, q: "How can you enforce Conditional Access policies based on the physical location of the user?", options: ["Using IP ranges in Named Locations", "Using the user's home address", "Using GPS tracking on the device", "Using the user's ISP name"], correct: 0 },
            { id: 5, q: "What is the role of a 'managed identity' in Azure?", options: ["To allow users to manage their own accounts", "To provide an identity for Azure services to authenticate to other resources", "To manage external guest users", "To store certificates securely"], correct: 1 }
        ],
        'AZ-900': [
            { id: 1, q: "Define 'Cloud Computing'.", options: ["Running applications on a local server", "The delivery of computing services over the internet", "A way to store files on a USB drive", "Using a supercomputer for gaming"], correct: 1 },
            { id: 2, q: "Which cloud model is a combination of public and private clouds?", options: ["Public Cloud", "Private Cloud", "Hybrid Cloud", "Community Cloud"], correct: 2 },
            { id: 3, q: "What does 'SaaS' stand for?", options: ["Server as a Service", "Software as a Service", "System as a Service", "Storage as a Service"], correct: 1 },
            { id: 4, q: "Which Azure service provides a platform for serverless code execution?", options: ["Azure Virtual Machines", "Azure App Service", "Azure Functions", "Azure Kubernetes Service"], correct: 2 },
            { id: 5, q: "What is the Azure 'Free Account' duration for major services?", options: ["1 month", "6 months", "12 months", "Forever"], correct: 2 }
        ],
        'MS-700': [
            { id: 1, q: "Which portal is primarily used to manage Microsoft Teams settings?", options: ["Exchange Admin Center", "M365 Admin Center", "Teams Admin Center", "Azure Active Directory Portal"], correct: 2 },
            { id: 2, q: "What is the maximum number of members in a standard Microsoft Team?", options: ["5,000", "10,000", "25,000", "Unlimited"], correct: 2 },
            { id: 3, q: "Which tool can be used to troubleshoot Teams call quality for a specific user?", options: ["Teams Usage Report", "Call Quality Dashboard (CQD)", "Call Analytics", "Microsoft 365 Health Dashboard"], correct: 2 },
            { id: 4, q: "What role is required to manage all aspects of Microsoft Teams?", options: ["Global Administrator", "Teams Administrator", "User Administrator", "Billing Administrator"], correct: 1 },
            { id: 5, q: "How do you enable guest access in Microsoft Teams?", options: ["It is enabled by default for all", "In the Teams Admin Center under Org-wide settings", "By adding a guest to a private channel", "Using a PowerShell command only"], correct: 1 }
        ]
    };

    const handleAIAutoGenerate = () => {
        setIsGenerating(true);
        setTimeout(() => {
            const code = newPath.toUpperCase().trim();
            let pool = autoGenerateLibrary[code] || [
                { id: 101, q: `Identify the core security concepts relevant to ${newPath}.`, options: ["Confidentiality", "Integrity", "Availability", "All of the above"], correct: 3 },
                { id: 102, q: `What is the primary responsibility of a ${newPath} specialist?`, options: ["System Monitoring", "User Training", "Implementation & Configuration", "Cost Management"], correct: 2 },
                { id: 103, q: `Which tool is most commonly used for ${newPath} management?`, options: ["Microsoft Sentinel", "Microsoft Purview", "Entra ID", "Power Platform"], correct: 0 },
                { id: 104, q: `What is the first step in a ${newPath} deployment project?`, options: ["Budget Approval", "Needs Analysis", "Buying Hardware", "Hiring Staff"], correct: 1 },
                { id: 105, q: `Which framework is best for ${newPath} governance?`, options: ["NIST", "ISO 27001", "ITIL", "COBIT"], correct: 1 }
            ];

            // Shuffle pool
            const shuffled = [...pool].sort(() => 0.5 - Math.random());
            const selected = shuffled.slice(0, Math.min(numToGenerate, shuffled.length));

            setQuestions(selected.map((q, i) => ({ ...q, id: Date.now() + i })));
            setIsGenerating(false);
            alert(`AI Analysis Complete: Generated ${selected.length} unique questions for ${newPath}.`);
        }, 1500);
    };

    const handleBulkUpload = (event: any) => {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e: any) => {
            const text = e.target.result;
            const rows = text.split('\n');
            const newQuestions = rows.map((row: string, index: number) => {
                const parts = row.split(',');
                if (parts.length < 6) return null;
                const [q, o1, o2, o3, o4, correct] = parts;
                return {
                    id: Date.now() + index,
                    q: q.trim().replace(/^"|"$/g, ''),
                    options: [o1, o2, o3, o4].map(o => o.trim().replace(/^"|"$/g, '')),
                    correct: parseInt(correct.trim())
                };
            }).filter((q: any) => q !== null);

            setQuestions([...questions, ...newQuestions]);
            alert(`Successfully uploaded ${newQuestions.length} questions.`);
        };
        reader.readAsText(file);
    };

    // Form State for new assessment
    const [newTitle, setNewTitle] = useState('');
    const [newThreshold, setNewThreshold] = useState<number>(80);
    const [newPath, setNewPath] = useState('General');

    const normalizeAssessmentLearners = (users: any[]): any[] => {
        const deduped = new Map<string, any>();
        const rolePriority: Record<string, number> = {
            Owner: 0,
            Member: 1,
            Visitor: 2
        };

        (users || []).forEach((user: any, index: number) => {
            const rawEmail = (user?.email || user?.Email || '').toString().trim();
            if (!rawEmail) {
                return;
            }

            const normalizedEmail = rawEmail.toLowerCase();
            const role = (user?.role || '').toString().trim() || (
                user?.siteGroup === 'Owners' ? 'Owner' :
                    user?.siteGroup === 'Visitors' ? 'Visitor' :
                        'Member'
            );
            const siteGroup = (user?.siteGroup || user?.group || '').toString().trim() || (
                role === 'Owner' ? 'Owners' :
                    role === 'Visitor' ? 'Visitors' :
                        'Members'
            );

            const normalizedUser = {
                ...user,
                id: user?.id || user?.Id || rawEmail || `assessment-learner-${index}`,
                name: user?.name || user?.Title || rawEmail,
                email: rawEmail,
                role,
                siteGroup
            };

            const existing = deduped.get(normalizedEmail);
            if (!existing || (rolePriority[role] ?? 99) < (rolePriority[existing.role] ?? 99)) {
                deduped.set(normalizedEmail, normalizedUser);
            }
        });

        return Array.from(deduped.values()).sort((left: any, right: any) => {
            const roleDelta = (rolePriority[left.role] ?? 99) - (rolePriority[right.role] ?? 99);
            if (roleDelta !== 0) {
                return roleDelta;
            }

            return (left.name || '').localeCompare(right.name || '');
        });
    };

    const filteredAssessmentLearners = useMemo(() => {
        const normalizedSearch = assessmentUserSearchTerm.trim().toLowerCase();
        if (!normalizedSearch) {
            return assessmentLearnerState.users;
        }

        return assessmentLearnerState.users.filter((user: any) => {
            return [
                user?.name,
                user?.email,
                user?.jobTitle,
                user?.siteGroup,
                user?.employeeId
            ].some((value) => (value || '').toString().toLowerCase().includes(normalizedSearch));
        });
    }, [assessmentLearnerState.users, assessmentUserSearchTerm]);

    const selectedAssessmentUserEmails = useMemo(() => {
        return new Set(
            selectedUsers
                .map((user: any) => (user?.email || '').toString().trim().toLowerCase())
                .filter((email: string) => !!email)
        );
    }, [selectedUsers]);

    const handleOpenPushModal = (assessment: any) => {
        setPushingAsmt(assessment);
        setShowPushModal(true);
        setSelectedUsers([]);
        setAssessmentUserSearchTerm('');
        setAssessmentScheduledDate(new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]);
    };

    const handleClosePushModal = () => {
        if (isAssigningAssessment) {
            return;
        }

        setShowPushModal(false);
        setPushingAsmt(null);
        setSelectedUsers([]);
        setAssessmentUserSearchTerm('');
        setAssessmentScheduledDate(new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]);
    };

    const toggleSelectedUser = (user: any) => {
        const normalizedEmail = (user?.email || user?.Email || '').toString().trim().toLowerCase();
        if (!normalizedEmail) {
            return;
        }

        setSelectedUsers((prev) => {
            const alreadySelected = prev.some((selectedUser: any) =>
                (selectedUser?.email || '').toString().trim().toLowerCase() === normalizedEmail
            );

            if (alreadySelected) {
                return prev.filter((selectedUser: any) =>
                    (selectedUser?.email || '').toString().trim().toLowerCase() !== normalizedEmail
                );
            }

            return [
                ...prev,
                {
                    ...user,
                    email: user?.email || user?.Email || normalizedEmail,
                    name: user?.name || user?.Title || normalizedEmail
                }
            ];
        });
    };

    const loadAssessmentResults = React.useCallback(async (options?: { silent?: boolean }) => {
        if (!options?.silent) {
            setAssessmentResultsLoading(true);
        }

        try {
            const results = await SharePointService.getAssessmentTrackerItems();
            setAssessmentResults((prev) => JSON.stringify(prev) === JSON.stringify(results) ? prev : results);
            setAssessmentResultsError(null);
        } catch (error) {
            console.error('[Assessments] Failed to load learner results tracker data', error);
            setAssessmentResults((prev) => prev.length === 0 ? prev : []);
            setAssessmentResultsError('Unable to load assessment records from SharePoint right now.');
        } finally {
            if (!options?.silent) {
                setAssessmentResultsLoading(false);
            }
        }
    }, []);

    useEffect(() => {
        const load = () => {
            const custom = localStorage.getItem('lmsAdminAssessments');
            if (custom) {
                try { setAdminAssessments(JSON.parse(custom)); } catch (e) { }
            }
        };
        load();
        void loadAssessmentResults();
        window.addEventListener('storage', load);
        return () => window.removeEventListener('storage', load);
    }, [loadAssessmentResults]);

    useEffect(() => {
        let isCancelled = false;

        const loadAssessmentLearners = async () => {
            try {
                const learners = await SharePointService.getAssessmentAssignmentLearners();

                if (!isCancelled) {
                    setAssessmentLearnerState({
                        users: normalizeAssessmentLearners(learners),
                        loading: false
                    });
                }
            } catch (error) {
                console.error('[Assessments] Failed to load direct assignment learners', error);

                if (!isCancelled) {
                    setAssessmentLearnerState({
                        users: normalizeAssessmentLearners(allUsers),
                        loading: false
                    });
                }
            }
        };

        void loadAssessmentLearners();

        return () => {
            isCancelled = true;
        };
    }, []);

    const handleCreateAssessment = () => {
        if (!newTitle) { alert('Please provide an assessment title.'); return; }

        const newAsmt = {
            id: 'admin_' + Date.now(),
            title: newTitle,
            path: newPath,
            questions: questions.length || 10,
            questionsArr: questions.length > 0 ? questions : null,
            threshold: newThreshold,
            isPublished: true,
            certCode: newPath
        };

        const updated = [...adminAssessments, newAsmt];
        setAdminAssessments(updated);
        localStorage.setItem('lmsAdminAssessments', JSON.stringify(updated));

        // reset & close
        setNewTitle('');
        setNewPath('General');
        setNewThreshold(80);
        setQuestions([]);
        setCreating(false);
        setSubTab('management');
        alert("Assessment successfully published and synced with User Portal.");
    };
    const handleDeleteResult = async (resId: number) => {
        if (!window.confirm("Remove this learner's assessment attempt?")) return;

        try {
            await SharePointService.deleteAssessmentTrackerItem(Number(resId));
            setAssessmentResults((prev) => prev.filter((item) => Number(item.id) !== Number(resId)));
        } catch (error) {
            console.error('[Assessments] Failed to delete assessment tracker item', error);
            alert('Unable to delete the assessment right now.');
        }
    };

    const handleDeleteAssessment = (id: any) => {
        if (!window.confirm("Delete this assessment system-wide?")) return;
        const up = adminAssessments.filter(a => a.id !== id);
        setAdminAssessments(up);
        localStorage.setItem('lmsAdminAssessments', JSON.stringify(up));
    };

    const handlePushAssignment = async () => {
        if (!pushingAsmt || isAssigningAssessment) {
            if (!pushingAsmt) {
                alert('Select an assessment before assigning it.');
            }
            return;
        }

        if (selectedUsers.length === 0) {
            alert('Select at least one learner before pushing the assessment.');
            return;
        }

        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const selectedAssignmentDate = new Date(assessmentScheduledDate);
        if (!assessmentScheduledDate || Number.isNaN(selectedAssignmentDate.getTime()) || selectedAssignmentDate < today) {
            alert('Select a valid future scheduled date before assigning the assessment.');
            return;
        }

        try {
            setIsAssigningAssessment(true);

            const assignmentSummary = await SharePointService.assignAssessmentToSelectedLearners({
                id: pushingAsmt.id,
                title: pushingAsmt.title,
                assessmentName: pushingAsmt.path || pushingAsmt.certCode || pushingAsmt.title,
                certCode: pushingAsmt.certCode || pushingAsmt.path,
                threshold: pushingAsmt.threshold,
                questions: pushingAsmt.questions,
                questionsArr: pushingAsmt.questionsArr || null,
                provider: pushingAsmt.provider || 'Internal',
                duration: pushingAsmt.duration || '20 Mins'
            }, selectedUsers, assessmentScheduledDate);

            const successMessage = `Assigned "${pushingAsmt.title}" to ${assignmentSummary.assignedCount} selected learners. Skipped ${assignmentSummary.skippedCount} existing assignments.`;
            alert(successMessage);

            if (updateAdminNotifications) {
                await updateAdminNotifications({
                    id: Date.now(),
                    title: 'Assessment Assigned',
                    text: successMessage,
                    time: 'Just now',
                    type: 'success'
                });
            }

            setSelectedUsers([]);
            setAssessmentUserSearchTerm('');
            setAssessmentScheduledDate(new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]);
            setPushingAsmt(null);
            setShowPushModal(false);
            await loadAssessmentResults({ silent: true });
            window.setTimeout(() => {
                window.dispatchEvent(new Event(LMS_AUDIT_REFRESH_EVENT));
            }, 500);
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : 'Assessment assignment failed.';
            console.error('ASSIGNMENT FAILED:', error);
            alert(errorMessage);
        } finally {
            setIsAssigningAssessment(false);
        }
    };

    const combinedAssessments = [...defaultAssessments, ...adminAssessments];

    const filteredResults = assessmentResults.filter((res: IAssessmentTrackerItem) =>
        (res.assessment || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
        (res.learner || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
        (res.learnerEmail || '').toLowerCase().includes(searchTerm.toLowerCase())
    );

    if (creating) {
        return (
            <div className="fade-in">
                <header className="view-header" style={{ marginBottom: '1.5rem' }}>
                    <div>
                        <h1 className="view-title">Assessment Builder</h1>
                        <p style={{ color: 'var(--text-muted)', fontWeight: 600 }}>Design and configure a new certification assessment.</p>
                    </div>
                </header>

                <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: '2rem', maxWidth: 'min(850px, 100%)' }}>
                    <div style={{ background: 'white', padding: '2.5rem', borderRadius: '24px', border: '1px solid #e2e8f0', boxShadow: '0 10px 25px -5px rgba(0,0,0,0.05)' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '16px', marginBottom: '2.5rem', paddingBottom: '1.5rem', borderBottom: '1px solid #f1f5f9' }}>
                            <div style={{ width: '56px', height: '56px', borderRadius: '16px', background: 'var(--primary-light)', color: 'var(--primary)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                <FileQuestion size={28} style={{ opacity: 0.9 }} />
                            </div>
                            <div>
                                <h2 style={{ fontSize: '1.4rem', fontWeight: 800, color: '#0f172a', margin: 0, letterSpacing: '-0.02em' }}>Core Configuration</h2>
                                <p style={{ fontSize: '0.9rem', color: '#64748b', margin: '4px 0 0 0', fontWeight: 600 }}>Define the assessment tracking parameters.</p>
                            </div>
                        </div>

                        <div style={{ display: 'grid', gap: '1.5rem' }}>
                            <div>
                                <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#475569', marginBottom: '0.5rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Assessment Title *</label>
                                <input type="text" className="input-field" placeholder="e.g. Identity & Access Management Final" value={newTitle} onChange={e => setNewTitle(e.target.value)} style={{ padding: '0.9rem 1.2rem', fontSize: '1.05rem', background: '#f8fafc', border: '2px solid #e2e8f0', borderRadius: '12px', width: '100%', outline: 'none', transition: 'border-color 0.2s' }} />
                                <div style={{ fontSize: '0.75rem', color: '#94a3b8', marginTop: '0.5rem', fontWeight: 600 }}>The public name users will see.</div>
                            </div>
                            <div>
                                <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#475569', marginBottom: '0.5rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Associated Path / Course Code *</label>
                                <input type="text" className="input-field" placeholder="e.g. SC-300" value={newPath} onChange={e => setNewPath(e.target.value)} style={{ padding: '0.9rem 1.2rem', fontSize: '1.05rem', background: '#f8fafc', border: '2px solid #e2e8f0', borderRadius: '12px', width: '100%', outline: 'none', transition: 'border-color 0.2s' }} />
                            </div>

                            <div className="responsive-two-column-grid" style={{ gap: '1.5rem', marginTop: '0.5rem' }}>
                                <div>
                                    <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#475569', marginBottom: '0.5rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Passing Threshold (%) *</label>
                                    <input type="number" className="input-field" value={newThreshold} onChange={e => setNewThreshold(parseInt(e.target.value))} min={10} max={100} style={{ padding: '0.9rem 1.2rem', fontSize: '1rem', background: '#f8fafc', border: '2px solid #e2e8f0', borderRadius: '12px', width: '100%', outline: 'none' }} />
                                </div>
                                <div>
                                    <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#475569', marginBottom: '0.5rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Allowed Retakes</label>
                                    <input type="number" className="input-field" defaultValue={3} min={1} max={10} style={{ padding: '0.9rem 1.2rem', fontSize: '1rem', background: '#f8fafc', border: '2px solid #e2e8f0', borderRadius: '12px', width: '100%', outline: 'none' }} />
                                    <div style={{ fontSize: '0.7rem', color: '#94a3b8', marginTop: '0.5rem', fontWeight: 600 }}>Set 0 for unlimited attempts.</div>
                                </div>
                            </div>
                        </div>

                        <div style={{ marginTop: '3.5rem', paddingTop: '2.5rem', borderTop: '2px dashed #e2e8f0' }}>
                            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '1.5rem' }}>
                                <div>
                                    <h3 style={{ fontSize: '1.2rem', fontWeight: 800, color: '#0f172a', margin: 0 }}>Question Bank</h3>
                                    <p style={{ fontSize: '0.85rem', color: '#64748b', margin: '4px 0 0 0', fontWeight: 600 }}>Upload or manually add assessment questions.</p>
                                </div>
                                <div style={{ display: 'flex', gap: '0.8rem', alignItems: 'center' }}>
                                    <div style={{ display: 'flex', alignItems: 'center', background: '#f1f5f9', padding: '0.4rem 0.8rem', borderRadius: '10px', gap: '8px' }}>
                                        <span style={{ fontSize: '0.75rem', fontWeight: 800, color: '#64748b' }}>QTY:</span>
                                        <input
                                            type="number"
                                            min="1"
                                            max="20"
                                            value={numToGenerate}
                                            onChange={(e) => setNumToGenerate(parseInt(e.target.value))}
                                            style={{ width: '40px', border: 'none', background: 'transparent', fontWeight: 800, fontSize: '0.85rem', outline: 'none' }}
                                        />
                                    </div>
                                    <button
                                        onClick={handleAIAutoGenerate}
                                        disabled={isGenerating}
                                        style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '0.6rem 1.2rem', background: 'linear-gradient(135deg, #6366f1 0%, #a855f7 100%)', border: 'none', color: 'white', borderRadius: '10px', fontSize: '0.85rem', fontWeight: 700, cursor: 'pointer', opacity: isGenerating ? 0.7 : 1 }}
                                    >
                                        <Sparkles size={16} /> {isGenerating ? 'Analyzing...' : 'Auto-Generate (AI)'}
                                    </button>
                                    <label className="btn-secondary" style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '0.6rem 1rem', background: '#f1f5f9', border: 'none', color: '#475569', borderRadius: '10px', fontSize: '0.85rem', fontWeight: 700, cursor: 'pointer' }}>
                                        <Upload size={16} /> Bulk Upload CSV
                                        <input type="file" accept=".csv" style={{ display: 'none' }} onChange={handleBulkUpload} />
                                    </label>
                                </div>
                            </div>

                            {questions.length === 0 ? (
                                <div
                                    style={{ background: '#f8fafc', border: '2px dashed #cbd5e1', borderRadius: '20px', padding: '3.5rem 2rem', textAlign: 'center', transition: 'all 0.2s', cursor: 'pointer' }}
                                    onClick={() => setQuestions([{ id: 1, q: '', options: ['', '', '', ''], correct: 0 }])}
                                >
                                    <div style={{ width: '72px', height: '72px', background: 'white', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 1.5rem auto', color: '#94a3b8', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.05)' }}>
                                        <FileQuestion size={36} />
                                    </div>
                                    <h4 style={{ fontSize: '1.15rem', fontWeight: 800, color: '#334155', margin: '0 0 0.5rem 0' }}>No questions added yet</h4>
                                    <p style={{ fontSize: '0.95rem', color: '#64748b', margin: '0 0 2rem 0', fontWeight: 500 }}>Drop test questions here or add them manually row by row.</p>
                                    <button className="btn-primary" style={{ display: 'inline-flex', padding: '0.75rem 2rem', borderRadius: '12px', fontSize: '0.95rem', fontWeight: 800 }}>+ Add First Question</button>
                                </div>
                            ) : (
                                <div style={{ display: 'grid', gap: '1.5rem' }}>
                                    {questions.map((q, qIdx) => (
                                        <div key={q.id} style={{ background: 'white', border: '1px solid #e2e8f0', borderRadius: '16px', padding: '1.5rem', position: 'relative' }}>
                                            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '1rem' }}>
                                                <span style={{ fontSize: '0.75rem', fontWeight: 900, color: 'var(--primary)', textTransform: 'uppercase' }}>Question {qIdx + 1}</span>
                                                <button onClick={() => setQuestions(questions.filter(item => item.id !== q.id))} style={{ background: 'none', border: 'none', color: '#ef4444', cursor: 'pointer' }}><Trash2 size={16} /></button>
                                            </div>
                                            <input
                                                type="text"
                                                className="input-field"
                                                placeholder="Enter question text..."
                                                value={q.q}
                                                onChange={(e) => {
                                                    const up = [...questions];
                                                    up[qIdx].q = e.target.value;
                                                    setQuestions(up);
                                                }}
                                                style={{ marginBottom: '1.5rem', fontWeight: 600 }}
                                            />
                                            <div className="responsive-two-column-grid">
                                                {q.options.map((opt: string, oIdx: number) => (
                                                    <div key={oIdx} style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                                        <input
                                                            type="radio"
                                                            name={`correct_${q.id}`}
                                                            checked={q.correct === oIdx}
                                                            onChange={() => {
                                                                const up = [...questions];
                                                                up[qIdx].correct = oIdx;
                                                                setQuestions(up);
                                                            }}
                                                        />
                                                        <input
                                                            type="text"
                                                            className="input-field"
                                                            placeholder={`Option ${oIdx + 1}`}
                                                            value={opt}
                                                            onChange={(e) => {
                                                                const up = [...questions];
                                                                up[qIdx].options[oIdx] = e.target.value;
                                                                setQuestions(up);
                                                            }}
                                                            style={{ fontSize: '0.85rem', padding: '0.6rem' }}
                                                        />
                                                    </div>
                                                ))}
                                            </div>
                                        </div>
                                    ))}
                                    <button
                                        className="btn-secondary"
                                        onClick={() => setQuestions([...questions, { id: Date.now(), q: '', options: ['', '', '', ''], correct: 0 }])}
                                        style={{ padding: '0.75rem', borderRadius: '12px', background: '#f8fafc', border: '2px dashed #e2e8f0', color: '#64748b', fontWeight: 700 }}
                                    >
                                        + Add Another Question
                                    </button>
                                </div>
                            )}
                        </div>

                        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '1rem', marginTop: '3.5rem', paddingTop: '2rem', borderTop: '1px solid #f1f5f9' }}>
                            <button className="btn-secondary" onClick={() => setCreating(false)} style={{ padding: '0.85rem 1.8rem', borderRadius: '14px', fontSize: '1rem', border: '2px solid #e2e8f0', color: '#475569', fontWeight: 700, background: 'transparent' }}>Discard changes</button>
                            <button className="btn-primary" onClick={handleCreateAssessment} style={{ padding: '0.85rem 2.2rem', borderRadius: '14px', fontSize: '1rem', boxShadow: '0 8px 16px rgba(14, 165, 233, 0.25)' }}>Publish Assessment</button>
                        </div>
                    </div>
                </div>
            </div>
        );
    }

    return (
        <div className="fade-in">
            <header className="view-header" style={{ marginBottom: '1.5rem' }}>
                <div>
                    <h1 className="view-title">Assessment Central</h1>
                    <p style={{ color: 'var(--text-muted)', fontWeight: 600 }}>Create evaluations and track learner attempts.</p>
                </div>
                {subTab === 'management' && <button className="btn-primary" onClick={() => setCreating(true)}>+ Create Assessment</button>}
            </header>

            <div style={{ display: 'flex', gap: '0.5rem', marginBottom: '2rem', background: '#f1f5f9', padding: '0.5rem', borderRadius: '16px', width: 'fit-content' }}>
                <button
                    onClick={() => setSubTab('management')}
                    style={{
                        padding: '0.75rem 1.5rem', borderRadius: '12px', border: 'none',
                        background: subTab === 'management' ? 'white' : 'transparent',
                        color: subTab === 'management' ? 'var(--primary)' : '#64748b',
                        fontWeight: 800, cursor: 'pointer',
                        boxShadow: subTab === 'management' ? '0 4px 6px rgba(0,0,0,0.05)' : 'none'
                    }}
                >
                    Assessment Builder
                </button>
                <button
                    onClick={() => setSubTab('results')}
                    style={{
                        padding: '0.75rem 1.5rem', borderRadius: '12px', border: 'none',
                        background: subTab === 'results' ? 'white' : 'transparent',
                        color: subTab === 'results' ? 'var(--primary)' : '#64748b',
                        fontWeight: 800, cursor: 'pointer',
                        boxShadow: subTab === 'results' ? '0 4px 6px rgba(0,0,0,0.05)' : 'none'
                    }}
                >
                    Learner Results Tracker
                </button>
            </div>

            {subTab === 'management' ? (
                <div className="table-container">
                    <table className="admin-table">
                        <thead><tr><th>Assessment Title</th><th>Associated Path</th><th>Pass Threshold</th><th>Status</th><th>Actions</th></tr></thead>
                        <tbody>
                            {combinedAssessments.map(a => (
                                <tr key={a.id}>
                                    <td style={{ fontWeight: 800, color: '#0f172a' }}>{a.title}</td>
                                    <td><span className="pill activity">{a.path}</span></td>
                                    <td style={{ fontWeight: 700 }}>{a.threshold}%</td>
                                    <td>{a.isPublished ? <StatusBadge status="completed" label="Published" /> : <StatusBadge status="scheduled" label="Draft" />}</td>
                                    <td>
                                        <div className="action-btns">
                                            <button className="btn-icon" title="Assign to Selected" style={{ color: 'var(--primary)' }} onClick={() => handleOpenPushModal(a)}><Send size={16} /></button>
                                            <button className="btn-icon" onClick={() => setCreating(true)}><Edit size={16} /></button>
                                            {String(a.id).startsWith('admin_') && (
                                                <button className="btn-icon" style={{ color: '#ef4444' }} onClick={() => handleDeleteAssessment(a.id)}><Trash2 size={16} /></button>
                                            )}
                                        </div>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            ) : (
                <>
                    <div className="search-box-unified" style={{ marginBottom: '2rem', width: '100%', height: '56px' }}>
                        <Search size={20} />
                        <input
                            type="text"
                            placeholder="Search by learner or assessment..."
                            value={searchTerm}
                            onChange={(e) => setSearchTerm(e.target.value)}
                        />
                    </div>
                    <div className="table-container">
                        <table className="admin-table">
                            <thead>
                                <tr>
                                    <th>Learner</th>
                                    <th>Assessment</th>
                                    <th>Created</th>
                                    <th style={{ textAlign: 'right' }}>Actions</th>
                                </tr>
                            </thead>
                            <tbody>
                                {assessmentResultsLoading ? (
                                    <tr>
                                        <td colSpan={4} style={{ textAlign: 'center', padding: '4rem', color: '#64748b', fontWeight: 700 }}>
                                            Loading assessment records from SharePoint...
                                        </td>
                                    </tr>
                                ) : assessmentResultsError && filteredResults.length === 0 ? (
                                    <tr>
                                        <td colSpan={4} style={{ textAlign: 'center', padding: '4rem', color: '#ef4444', fontWeight: 700 }}>
                                            {assessmentResultsError}
                                        </td>
                                    </tr>
                                ) : filteredResults.map((res: IAssessmentTrackerItem) => (
                                    <tr key={res.id}>
                                        <td>
                                            <div style={{ fontWeight: 850, color: '#1e293b' }}>{res.learner || 'Not Available'}</div>
                                            <div style={{ fontSize: '0.75rem', color: '#94a3b8', fontWeight: 700 }}>{res.learnerEmail || 'No email'}</div>
                                        </td>
                                        <td style={{ fontWeight: 850, color: '#1e293b' }}>{res.assessment || 'Assessment'}</td>
                                        <td style={{ color: '#64748b', fontSize: '0.85rem', fontWeight: 700 }}>
                                            {res.created ? new Date(res.created).toLocaleString() : 'Not Available'}
                                        </td>
                                        <td style={{ textAlign: 'right' }}>
                                            <button className="btn-icon" style={{ color: '#ef4444' }} onClick={() => handleDeleteResult(res.id)} title="Delete assessment">
                                                <Trash2 size={16} />
                                            </button>
                                        </td>
                                    </tr>
                                ))}
                                {!assessmentResultsLoading && !assessmentResultsError && filteredResults.length === 0 && (
                                    <tr>
                                        <td colSpan={4} style={{ textAlign: 'center', padding: '4rem', color: '#94a3b8', fontWeight: 700 }}>
                                            No assessment assignments found in the SharePoint assignment list.
                                        </td>
                                    </tr>
                                )}
                            </tbody>
                        </table>
                    </div>
                </>
            )}

            {/* Push Assignment Modal */}
            {showPushModal && (
                <div className="modal-overlay" style={{
                    position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
                    background: 'rgba(15, 23, 42, 0.4)', backdropFilter: 'blur(8px)',
                    display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999
                }}>
                    <div className="fade-in direct-learner-modal">
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '2rem' }}>
                            <div>
                                <h2 style={{ fontSize: '1.85rem', fontWeight: 950, color: '#1e293b', margin: 0, letterSpacing: '-0.04em' }}>Direct Learner Assignment</h2>
                                <p style={{ color: 'var(--text-muted)', fontSize: '0.95rem', fontWeight: 600, marginTop: '0.4rem' }}>
                                    Push <b>{pushingAsmt?.title}</b> only to the selected learners in SharePoint groups.
                                </p>
                            </div>
                            <button onClick={handleClosePushModal} className="btn-icon" disabled={isAssigningAssessment}><X size={24} /></button>
                        </div>
                        <div className="direct-learner-modal-body">
                            <div className="form-group">
                                <label style={{ fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase' }}>Search Learners</label>
                                <div className="direct-learner-search">
                                    <Search size={18} />
                                    <input
                                        className="input-field direct-learner-search-input"
                                        value={assessmentUserSearchTerm}
                                        onChange={e => setAssessmentUserSearchTerm(e.target.value)}
                                        placeholder="Search by name, email, designation, or SharePoint group"
                                    />
                                </div>
                                <div className="direct-learner-search-meta">
                                    <span>{filteredAssessmentLearners.length} result{filteredAssessmentLearners.length === 1 ? '' : 's'}</span>
                                    {assessmentUserSearchTerm && (
                                        <button
                                            type="button"
                                            className="direct-learner-clear"
                                            onClick={() => setAssessmentUserSearchTerm('')}
                                        >
                                            Clear search
                                        </button>
                                    )}
                                </div>
                                {assessmentLearnerState.loading && (
                                    <div className="direct-learner-loading">
                                        <Loader2 size={18} className="direct-learner-spinner" />
                                        <span>Loading SharePoint learners...</span>
                                    </div>
                                )}
                            </div>
                            <div className="form-group">
                                <label style={{ fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase', display: 'block' }}>
                                    Scheduled Date
                                </label>
                                <input
                                    type="date"
                                    className="input-field"
                                    value={assessmentScheduledDate}
                                    min={new Date().toISOString().split('T')[0]}
                                    onChange={(event) => setAssessmentScheduledDate(event.target.value)}
                                />
                            </div>
                            <div>
                                <div className="direct-learner-section-header">
                                    <label style={{ fontSize: '0.85rem', fontWeight: 800, color: '#64748b', marginBottom: '0.75rem', textTransform: 'uppercase', display: 'block' }}>
                                        Select Learners
                                    </label>
                                    <div className="selected-count">
                                        {selectedUsers.length} learner{selectedUsers.length === 1 ? '' : 's'} selected
                                    </div>
                                </div>
                                <div className="learner-list">
                                    {assessmentLearnerState.loading && assessmentLearnerState.users.length === 0 ? (
                                        <div className="direct-learner-empty">
                                            Loading learners...
                                        </div>
                                    ) : filteredAssessmentLearners.length === 0 ? (
                                        <div className="direct-learner-empty">
                                            No matching learners found
                                        </div>
                                    ) : filteredAssessmentLearners.map((user: any, index: number) => {
                                        const learnerEmail = (user?.email || user?.Email || '').toString().trim().toLowerCase();
                                        const isSelected = selectedAssessmentUserEmails.has(learnerEmail);
                                        const learnerKey = learnerEmail || `assessment-learner-${index}`;

                                        return (
                                            <div
                                                key={learnerKey}
                                                className={`learner-card ${isSelected ? 'selected' : ''}`}
                                                onClick={() => toggleSelectedUser(user)}
                                                onKeyDown={(event) => {
                                                    if (event.key === 'Enter' || event.key === ' ') {
                                                        event.preventDefault();
                                                        toggleSelectedUser(user);
                                                    }
                                                }}
                                                role="button"
                                                tabIndex={0}
                                                aria-pressed={isSelected}
                                            >
                                                <div className="learner-card-header">
                                                    <input
                                                        type="checkbox"
                                                        checked={isSelected}
                                                        onChange={() => toggleSelectedUser(user)}
                                                        onClick={(event) => event.stopPropagation()}
                                                        className="learner-card-checkbox"
                                                    />
                                                    <div className="learner-card-info">
                                                        <strong>{user.name}</strong>
                                                        <p>{user.email}</p>
                                                        <span className="learner-card-designation">{user.jobTitle || 'No designation'}</span>
                                                        <span className="learner-card-meta">{user.employeeId || 'No Employee ID'} | {user.department || 'Not Available'}</span>
                                                    </div>
                                                    {isSelected && (
                                                        <div className="learner-card-checkmark">
                                                            <CheckCircle2 size={18} />
                                                        </div>
                                                    )}
                                                </div>
                                            </div>
                                        );
                                    })}
                                </div>
                                {selectedUsers.length > 0 && (
                                    <div className="selected-learner-chips">
                                        {selectedUsers.map((user: any) => (
                                            <span key={(user.email || '').toString().toLowerCase()} className="selected-learner-chip">
                                                {user.name}
                                            </span>
                                        ))}
                                    </div>
                                )}
                            </div>
                            <div style={{ borderRadius: '16px', border: '1px solid #dbeafe', background: '#eff6ff', padding: '1rem 1.25rem', color: '#1d4ed8' }}>
                                <div style={{ fontSize: '0.75rem', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: '0.45rem' }}>SharePoint Source</div>
                                <div style={{ fontSize: '0.95rem', fontWeight: 700 }}>Users are resolved from `Members (3)`, `Visitors (5)`, and `Owners (4)` via SharePoint REST.</div>
                                <div style={{ fontSize: '0.82rem', fontWeight: 600, marginTop: '0.5rem', color: '#334155' }}>
                                    Current learner directory count: {assessmentLearnerState.users.length}
                                </div>
                            </div>
                            <button
                                className="btn-primary push-cert-btn"
                                style={{ marginTop: '1rem' }}
                                onClick={() => { void handlePushAssignment(); }}
                                disabled={selectedUsers.length === 0 || assessmentLearnerState.loading || isAssigningAssessment}
                            >
                                {isAssigningAssessment ? `Assigning Selected (${selectedUsers.length})...` : `Assign to Selected (${selectedUsers.length})`} <Send size={20} />
                            </button>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}

function AuditView() {
    const [auditLogs, setAuditLogs] = useState<any[]>([]);
    const [loading, setLoading] = useState(false);
    const [loadError, setLoadError] = useState<string | null>(null);

    const loadData = React.useCallback(async () => {
        setLoading(true);
        try {
            const logsResult = await SharePointService.getAuditLogs()
                .then((value) => ({ value, error: null as Error | null }))
                .catch((error) => ({ value: [] as any[], error }));

            if (logsResult.error) {
                console.error('Failed to load audit logs', logsResult.error);
                setLoadError('Unable to load audit activity right now.');
            } else {
                setLoadError(null);
            }

            setAuditLogs(logsResult.value);
        } finally {
            setLoading(false);
        }
    }, []);

    useEffect(() => {
        void loadData();

        const handleAuditRefresh = () => {
            void loadData();
        };

        window.addEventListener(LMS_AUDIT_REFRESH_EVENT, handleAuditRefresh);
        return () => window.removeEventListener(LMS_AUDIT_REFRESH_EVENT, handleAuditRefresh);
    }, [loadData]);

    return (
        <div className="fade-in">
            <header className="view-header">
                <div>
                    <h1 className="view-title">Audit Logs</h1>
                    <p style={{ color: 'var(--text-muted)', fontWeight: 600 }}>Administrative oversight for SharePoint-backed LMS activity.</p>
                </div>
                <div style={{ display: 'flex', gap: '0.75rem', alignItems: 'center' }}>
                    <button className="btn-secondary" onClick={() => alert("Generating CSV Export...")}><BarChart3 size={18} /> Export CSV</button>
                </div>
            </header>

            <div className="table-container">
                <table className="admin-table">
                    <thead>
                        <tr>
                            <th>Operation Action</th>
                            <th>Administrative Actor</th>
                            <th>Target Entity</th>
                            <th>Event Timestamp</th>
                        </tr>
                    </thead>
                    <tbody>
                        {loading ? (
                            <tr><td colSpan={4} style={{ textAlign: 'center', padding: '4rem', color: '#64748b', fontWeight: 700 }}>Loading audit logs...</td></tr>
                        ) : loadError && auditLogs.length === 0 ? (
                            <tr><td colSpan={4} style={{ textAlign: 'center', padding: '4rem', color: '#ef4444', fontWeight: 700 }}>{loadError}</td></tr>
                        ) : auditLogs.length > 0 ? auditLogs.map(log => (
                            <tr key={log.id}>
                                <td>
                                    <span className="pill activity" style={{
                                        backgroundColor: log.action.toUpperCase().indexOf('REVOKE') !== -1 || log.action.toUpperCase().indexOf('DELETE') !== -1 ? '#fef2f2' : log.action.toUpperCase().indexOf('ASSIGN') !== -1 ? '#ecfdf5' : '#f1f5f9',
                                        color: log.action.toUpperCase().indexOf('REVOKE') !== -1 || log.action.toUpperCase().indexOf('DELETE') !== -1 ? '#ef4444' : log.action.toUpperCase().indexOf('ASSIGN') !== -1 ? '#10b981' : '#475569',
                                        fontWeight: 900
                                    }}>
                                        {log.action}
                                    </span>
                                </td>
                                <td style={{ fontWeight: 800 }}>{log.learnerName || 'SharePoint LMS'}</td>
                                <td style={{ fontWeight: 600 }}>{[log.assignmentName, log.learnerEmail].filter(Boolean).join(' ï¿½ ') || log.title}</td>
                                <td style={{ color: '#64748b', fontSize: '0.85rem', fontWeight: 700 }}>{new Date(log.assignmentDate || log.created).toLocaleString()}</td>
                            </tr>
                        )) : (
                            <tr><td colSpan={4} style={{ textAlign: 'center', padding: '4rem', color: '#94a3b8', fontWeight: 700 }}>No audit logs available</td></tr>
                        )}
                    </tbody>
                </table>
            </div>
        </div>
    );
}

function ReportsView({ realEnrollments, allUsers }: any) {
    type IUpcomingRenewalReportRow = IUpcomingRenewalRecord & { department: string };
    const [selectedDepartment, setSelectedDepartment] = React.useState('');
    const [sortMode, setSortMode] = React.useState('progress-desc');
    const [assessmentAssignments, setAssessmentAssignments] = React.useState<IAssessmentAssignmentRecord[]>([]);
    const [assessmentAssignmentsLoading, setAssessmentAssignmentsLoading] = React.useState(false);
    const [assessmentAssignmentsError, setAssessmentAssignmentsError] = React.useState<string | null>(null);
    const [reportEnrollments, setReportEnrollments] = React.useState<any[]>([]);
    const [reportEnrollmentsLoaded, setReportEnrollmentsLoaded] = React.useState(false);
    const [selectedLearner, setSelectedLearner] = React.useState<IDepartmentProgressLearner | null>(null);
    const [upcomingRenewals, setUpcomingRenewals] = React.useState<IUpcomingRenewalRecord[]>([]);
    const [upcomingRenewalsLoading, setUpcomingRenewalsLoading] = React.useState(false);
    const [upcomingRenewalsError, setUpcomingRenewalsError] = React.useState<string | null>(null);

    React.useEffect(() => {
        let isMounted = true;

        const loadAssessmentAssignments = async () => {
            setAssessmentAssignmentsLoading(true);

            try {
                const items = await SharePointService.getAllAssessmentAssignments();
                if (!isMounted) {
                    return;
                }

                setAssessmentAssignments((prev) =>
                    JSON.stringify(prev) === JSON.stringify(items) ? prev : items
                );
                setAssessmentAssignmentsError(null);
            } catch (error) {
                console.error('[Reports] Failed to load assessment assignments', error);
                if (isMounted) {
                    setAssessmentAssignments([]);
                    setAssessmentAssignmentsError('Unable to load learner assessment activity right now.');
                }
            } finally {
                if (isMounted) {
                    setAssessmentAssignmentsLoading(false);
                }
            }
        };

        void loadAssessmentAssignments();

        return () => {
            isMounted = false;
        };
    }, []);

    React.useEffect(() => {
        let isMounted = true;

        const loadReportEnrollments = async () => {
            try {
                const items = await SharePointService.getEnrollments('', '', {
                    excludeStatuses: ['Not Started']
                });
                if (!isMounted) {
                    return;
                }

                setReportEnrollments(items);
            } catch (error) {
                console.error('[Reports] Failed to load filtered enrollments', error);
                if (isMounted) {
                    setReportEnrollments([]);
                }
            } finally {
                if (isMounted) {
                    setReportEnrollmentsLoaded(true);
                }
            }
        };

        void loadReportEnrollments();

        return () => {
            isMounted = false;
        };
    }, [realEnrollments]);

    React.useEffect(() => {
        let isMounted = true;

        const loadUpcomingRenewals = async () => {
            setUpcomingRenewalsLoading(true);

            try {
                const items = await SharePointService.getUpcomingRenewalRecords(30);
                if (!isMounted) {
                    return;
                }

                setUpcomingRenewals(items);
                setUpcomingRenewalsError(null);
            } catch (error) {
                console.error('[Reports] Failed to load upcoming renewals', error);
                if (isMounted) {
                    setUpcomingRenewals([]);
                    setUpcomingRenewalsError('Unable to load upcoming renewals right now.');
                }
            } finally {
                if (isMounted) {
                    setUpcomingRenewalsLoading(false);
                }
            }
        };

        void loadUpcomingRenewals();

        return () => {
            isMounted = false;
        };
    }, [realEnrollments]);

    const formatReportDate = React.useCallback((value?: string): string => {
        if (!value) {
            return 'Not Available';
        }

        const parsedValue = new Date(value);
        if (Number.isNaN(parsedValue.getTime())) {
            return value;
        }

        return parsedValue.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    }, []);

    const getRenewalTone = React.useCallback((daysUntilRenewal: number): { background: string; color: string; border: string; } => {
        if (daysUntilRenewal <= 7) {
            return {
                background: '#fef2f2',
                color: '#dc2626',
                border: '#fecaca'
            };
        }

        return {
            background: '#fff7ed',
            color: '#ea580c',
            border: '#fed7aa'
        };
    }, []);

    const getEnrollmentProgressState = React.useCallback((enrollment: any): 'completed' | 'in-progress' | 'not-started' => {
        const progress = Number(enrollment?.progress ?? enrollment?.Progress ?? 0) || 0;
        const status = (enrollment?.status || enrollment?.Status || '').toString().trim().toLowerCase();

        if (status === 'completed' || progress >= 100) {
            return 'completed';
        }

        if (progress > 0 || status === 'scheduled' || status === 'rescheduled' || status === 'assigned' || status === 'in progress' || status === 'in-progress') {
            return 'in-progress';
        }

        return 'not-started';
    }, []);

    const learnerDeptMap = React.useMemo(() => {
        return (allUsers || []).reduce((acc: Record<string, string>, user: any) => {
            const email = (user?.email || user?.Email || '').toString().trim().toLowerCase();
            if (!email) {
                return acc;
            }

            acc[email] = (user?.department || '').toString().trim() || 'Not Available';
            return acc;
        }, {});
    }, [allUsers]);

    const directoryLearnersByDepartment = React.useMemo(() => {
        const groupedDepartments = new Map<string, Map<string, { learner: string; learnerEmail: string; department: string }>>();

        (allUsers || []).forEach((user: any, index: number) => {
            const learnerEmail = (user?.email || user?.Email || user?.login || '').toString().trim();
            const learnerName = (user?.name || user?.Title || learnerEmail || `Learner ${index + 1}`).toString().trim();
            const department = ((user?.department || '').toString().trim()) || 'Not Available';
            const learnerKey = (learnerEmail || learnerName || `learner-${index}`).toString().trim().toLowerCase();

            if (!learnerKey) {
                return;
            }

            if (!groupedDepartments.has(department)) {
                groupedDepartments.set(department, new Map<string, { learner: string; learnerEmail: string; department: string }>());
            }

            const departmentLearners = groupedDepartments.get(department)!;
            if (!departmentLearners.has(learnerKey)) {
                departmentLearners.set(learnerKey, {
                    learner: learnerName || 'Not Available',
                    learnerEmail,
                    department
                });
            }
        });

        return groupedDepartments;
    }, [allUsers]);

    const globalLearnerCount = React.useMemo(() => {
        const uniqueLearners = new Set<string>();

        directoryLearnersByDepartment.forEach((departmentLearners) => {
            departmentLearners.forEach((learner, learnerKey) => {
                const normalizedKey = (learner.learnerEmail || learnerKey || '').toString().trim().toLowerCase();
                if (normalizedKey) {
                    uniqueLearners.add(normalizedKey);
                }
            });
        });

        return uniqueLearners.size;
    }, [directoryLearnersByDepartment]);

    const upcomingRenewalRows = React.useMemo<IUpcomingRenewalReportRow[]>(() => {
        return (upcomingRenewals || []).map((item) => ({
            ...item,
            department: learnerDeptMap[(item.learnerEmail || '').toLowerCase()] || 'Not Available'
        }));
    }, [learnerDeptMap, upcomingRenewals]);

    const visibleUpcomingRenewals = React.useMemo<IUpcomingRenewalReportRow[]>(() => {
        if (!selectedDepartment) {
            return upcomingRenewalRows;
        }

        return upcomingRenewalRows.filter((item) => item.department === selectedDepartment);
    }, [selectedDepartment, upcomingRenewalRows]);

    const upcomingRenewalSummaryByLearner = React.useMemo(() => {
        const summary = new Map<string, { count: number; nextRenewal: IUpcomingRenewalReportRow | null }>();

        upcomingRenewalRows.forEach((item) => {
            const learnerEmail = (item.learnerEmail || '').toString().trim().toLowerCase();
            if (!learnerEmail) {
                return;
            }

            const existing = summary.get(learnerEmail);
            if (!existing) {
                summary.set(learnerEmail, {
                    count: 1,
                    nextRenewal: item
                });
                return;
            }

            const existingRenewalDate = new Date(existing.nextRenewal?.renewalDate || 0).getTime();
            const incomingRenewalDate = new Date(item.renewalDate || 0).getTime();
            summary.set(learnerEmail, {
                count: existing.count + 1,
                nextRenewal:
                    Number.isNaN(existingRenewalDate) || (!Number.isNaN(incomingRenewalDate) && incomingRenewalDate < existingRenewalDate)
                        ? item
                        : existing.nextRenewal
            });
        });

        return summary;
    }, [upcomingRenewalRows]);

    const assessmentCountsByLearner = React.useMemo(() => {
        const counts = new Map<string, number>();

        (assessmentAssignments || []).forEach((assignment) => {
            const normalizedEmail = (assignment?.userEmail || '').toString().trim().toLowerCase();
            if (!normalizedEmail) {
                return;
            }

            counts.set(normalizedEmail, Number(counts.get(normalizedEmail) || 0) + 1);
        });

        return counts;
    }, [assessmentAssignments]);

    const reportEnrollmentSource = React.useMemo(() => {
        if (reportEnrollmentsLoaded) {
            return reportEnrollments;
        }

        return realEnrollments || [];
    }, [realEnrollments, reportEnrollments, reportEnrollmentsLoaded]);

    const mergedData = React.useMemo<IDepartmentProgressLearner[]>(() => {
        return (reportEnrollmentSource || []).map((item: any) => {
            const learnerEmail = (item?.LearnerEmail || item?.learnerEmail || item?.userEmail || item?.email || '').toString().trim();
            const normalizedEmail = learnerEmail.toLowerCase();
            const progress = Number(item?.Progress ?? item?.progress ?? 0) || 0;

            return {
                learner: (item?.userName || item?.Name || item?.learner || learnerEmail || 'Not Available').toString(),
                learnerEmail,
                path: (item?.Title || item?.title || item?.name || item?.certName || item?.certificateName || item?.certCode || 'Not Available').toString(),
                progress,
                department: learnerDeptMap[normalizedEmail] || 'Not Available',
                status: getEnrollmentProgressState(item),
                assessmentCount: assessmentCountsByLearner.get(normalizedEmail) || 0
            } as IDepartmentProgressLearner;
        });
    }, [assessmentCountsByLearner, getEnrollmentProgressState, learnerDeptMap, reportEnrollmentSource]);

    const departmentDashboard = React.useMemo(() => {
        type LearnerAggregate = {
            learner: string;
            learnerEmail: string;
            department: string;
            pathCount: number;
            completedPathCount: number;
            startedPathCount: number;
            progressTotal: number;
            assessmentCount: number;
            pathNames: string[];
        };

        const groupedDepartments = new Map<string, Map<string, LearnerAggregate>>();

        directoryLearnersByDepartment.forEach((departmentLearners, department) => {
            const seededLearners = new Map<string, LearnerAggregate>();

            departmentLearners.forEach((directoryLearner, learnerKey) => {
                seededLearners.set(learnerKey, {
                    learner: directoryLearner.learner,
                    learnerEmail: directoryLearner.learnerEmail,
                    department,
                    pathCount: 0,
                    completedPathCount: 0,
                    startedPathCount: 0,
                    progressTotal: 0,
                    assessmentCount: Number(assessmentCountsByLearner.get((directoryLearner.learnerEmail || '').toLowerCase()) || 0),
                    pathNames: []
                });
            });

            groupedDepartments.set(department, seededLearners);
        });

        mergedData.forEach((learner) => {
            const department = (learner.department || 'Not Available').toString().trim() || 'Not Available';
            const learnerKey = (learner.learnerEmail || learner.learner || '').toString().trim().toLowerCase() || `unknown-${department}`;
            if (!groupedDepartments.has(department)) {
                groupedDepartments.set(department, new Map<string, LearnerAggregate>());
            }

            const departmentLearners = groupedDepartments.get(department)!;
            const existingLearner: LearnerAggregate = departmentLearners.get(learnerKey) || {
                learner: learner.learner,
                learnerEmail: learner.learnerEmail,
                department,
                pathCount: 0,
                completedPathCount: 0,
                startedPathCount: 0,
                progressTotal: 0,
                assessmentCount: Number(learner.assessmentCount || 0),
                pathNames: []
            };

            existingLearner.pathCount += 1;
            existingLearner.progressTotal += Number(learner.progress || 0);
            existingLearner.assessmentCount = Math.max(existingLearner.assessmentCount, Number(learner.assessmentCount || 0));

            if (learner.status === 'completed') {
                existingLearner.completedPathCount += 1;
                existingLearner.startedPathCount += 1;
            } else if (learner.status === 'in-progress') {
                existingLearner.startedPathCount += 1;
            }

            if (learner.path && existingLearner.pathNames.indexOf(learner.path) === -1) {
                existingLearner.pathNames.push(learner.path);
            }

            departmentLearners.set(learnerKey, existingLearner);
        });

        return Array.from(groupedDepartments.entries())
            .map(([department, learnersMap]) => {
                const learners = Array.from(learnersMap.values()).map((learner) => {
                    const averageProgress = learner.pathCount > 0
                        ? Math.round(learner.progressTotal / learner.pathCount)
                        : 0;
                    const status: 'completed' | 'in-progress' | 'not-started' =
                        learner.pathCount > 0 && learner.completedPathCount === learner.pathCount ? 'completed' :
                            learner.startedPathCount > 0 ? 'in-progress' :
                                'not-started';

                    return {
                        learner: learner.learner || 'Not Available',
                        learnerEmail: learner.learnerEmail || '',
                        department,
                        path: learner.pathCount === 0
                            ? 'No certifications assigned'
                            : learner.pathCount === 1
                                ? (learner.pathNames[0] || '1 certification path')
                                : `${learner.pathCount} certification paths`,
                        progress: averageProgress,
                        status,
                        pathCount: learner.pathCount,
                        completedPathCount: learner.completedPathCount,
                        assessmentCount: learner.assessmentCount
                    } as IDepartmentProgressLearner;
                }).sort((left, right) => {
                    const progressDelta = Number(right.progress || 0) - Number(left.progress || 0);
                    if (progressDelta !== 0) {
                        return progressDelta;
                    }

                    return (left.learner || '').localeCompare(right.learner || '');
                });

                const totalLearners = learners.length;
                const enrolledCount = learners.filter((learner) => learner.status !== 'not-started').length;
                const completedCount = learners.filter((learner) => learner.status === 'completed').length;
                const inProgressCount = learners.filter((learner) => learner.status === 'in-progress').length;
                const notStartedCount = learners.filter((learner) => learner.status === 'not-started').length;
                const enrolledPercent = totalLearners > 0 ? Math.round((enrolledCount / totalLearners) * 100) : 0;
                const completedPercent = totalLearners > 0 ? Math.round((completedCount / totalLearners) * 100) : 0;

                return {
                    department,
                    totalLearners,
                    enrolledCount,
                    completedCount,
                    inProgressCount,
                    notStartedCount,
                    enrolledPercent,
                    completedPercent,
                    learners
                } as IDepartmentProgressSummary;
            })
            .sort((left, right) => left.department.localeCompare(right.department));
    }, [assessmentCountsByLearner, directoryLearnersByDepartment, mergedData]);

    const filteredDepartmentDashboard = React.useMemo<IDepartmentProgressSummary[]>(() => {
        return (departmentDashboard || [])
            .map((department) => {
                const learners = (department.learners || []).filter((learner: IDepartmentProgressLearner) => {
                    const normalizedStatus = (learner.status || '').toString().trim().toLowerCase();
                    return normalizedStatus === 'completed' || normalizedStatus === 'in-progress';
                });

                if (learners.length === 0) {
                    return null;
                }

                const completedCount = learners.filter((learner) => learner.status === 'completed').length;
                const inProgressCount = learners.filter((learner) => learner.status === 'in-progress').length;
                const totalLearners = learners.length;
                const completedPercent = totalLearners > 0 ? Math.round((completedCount / totalLearners) * 100) : 0;

                return {
                    ...department,
                    learners,
                    totalLearners,
                    enrolledCount: totalLearners,
                    completedCount,
                    inProgressCount,
                    notStartedCount: 0,
                    enrolledPercent: totalLearners > 0 ? 100 : 0,
                    completedPercent
                } as IDepartmentProgressSummary;
            })
            .filter((department): department is IDepartmentProgressSummary => !!department);
    }, [departmentDashboard]);

    const dashboardTotals = React.useMemo(() => {
        return departmentDashboard.reduce((acc, department) => {
            acc.totalLearners += Number(department.totalLearners || 0);
            acc.enrolledCount += Number(department.enrolledCount || 0);
            acc.completedCount += Number(department.completedCount || 0);
            acc.inProgressCount += Number(department.inProgressCount || 0);
            acc.notStartedCount += Number(department.notStartedCount || 0);
            return acc;
        }, {
            totalLearners: 0,
            enrolledCount: 0,
            completedCount: 0,
            inProgressCount: 0,
            notStartedCount: 0
        });
    }, [departmentDashboard]);

    const visibleDepartments = React.useMemo(() => {
        if (!selectedDepartment) {
            return departmentDashboard;
        }

        return departmentDashboard.filter((item) => item.department === selectedDepartment);
    }, [departmentDashboard, selectedDepartment]);

    const visibleFilteredDepartments = React.useMemo(() => {
        if (!selectedDepartment) {
            return filteredDepartmentDashboard;
        }

        return filteredDepartmentDashboard.filter((item) => item.department === selectedDepartment);
    }, [filteredDepartmentDashboard, selectedDepartment]);

    const drillDownLearners = React.useMemo(() => {
        const rows = visibleFilteredDepartments.reduce((acc: IDepartmentProgressLearner[], department: IDepartmentProgressSummary) => {
            const departmentLearners = department.learners.map((learner: IDepartmentProgressLearner) => ({
                ...learner,
                department: department.department
            }));
            return acc.concat(departmentLearners);
        }, []);

        return rows.sort((left: IDepartmentProgressLearner, right: IDepartmentProgressLearner) => {
            switch (sortMode) {
                case 'progress-asc':
                    return Number(left.progress || 0) - Number(right.progress || 0);
                case 'name-asc':
                    return (left.learner || '').localeCompare(right.learner || '');
                case 'department-asc':
                    return (left.department || '').localeCompare(right.department || '');
                case 'progress-desc':
                default:
                    return Number(right.progress || 0) - Number(left.progress || 0);
            }
        });
    }, [visibleFilteredDepartments, sortMode]);

    const getLearnerProgressState = (learner: IDepartmentProgressLearner): 'completed' | 'in-progress' | 'not-started' => {
        const normalizedStatus = (learner.status || '').toString().trim().toLowerCase();
        if (normalizedStatus === 'completed') {
            return 'completed';
        }

        if (normalizedStatus === 'in-progress') {
            return 'in-progress';
        }

        return 'not-started';
    };

    const selectedLearnerCertifications = React.useMemo<IDepartmentProgressLearner[]>(() => {
        if (!selectedLearner?.learnerEmail) {
            return [];
        }

        const normalizedEmail = selectedLearner.learnerEmail.toLowerCase();
        return (mergedData || [])
            .filter((item) => (item.learnerEmail || '').toString().trim().toLowerCase() === normalizedEmail)
            .filter((item) => {
                const normalizedStatus = (item.status || '').toString().trim().toLowerCase();
                return normalizedStatus === 'completed' || normalizedStatus === 'in-progress';
            })
            .sort((left, right) => Number(right.progress || 0) - Number(left.progress || 0));
    }, [mergedData, selectedLearner]);

    const selectedLearnerAssessments = React.useMemo<IAssessmentAssignmentRecord[]>(() => {
        if (!selectedLearner?.learnerEmail) {
            return [];
        }

        const normalizedEmail = selectedLearner.learnerEmail.toLowerCase();
        return (assessmentAssignments || [])
            .filter((assignment) => (assignment.userEmail || '').toString().trim().toLowerCase() === normalizedEmail)
            .sort((left, right) => new Date(right.created || 0).getTime() - new Date(left.created || 0).getTime());
    }, [assessmentAssignments, selectedLearner]);

    const reportCategories = [
        { title: 'Departments', icon: <Briefcase size={20} />, value: departmentDashboard.length, trend: 'Learner departments in directory' },
        { title: 'Total Learners', icon: <Users size={20} />, value: globalLearnerCount, trend: 'Unique users across all departments' },
        { title: 'Enrolled Learners', icon: <TrendingUp size={20} />, value: dashboardTotals.enrolledCount, trend: `${dashboardTotals.inProgressCount} in progress` },
        { title: 'Completed Learners', icon: <CheckCircle2 size={20} />, value: dashboardTotals.completedCount, trend: 'Finished all assigned paths' },
        { title: 'Upcoming Renewals', icon: <Clock size={20} />, value: visibleUpcomingRenewals.length, trend: 'Due within the next 30 days' }
    ];

    return (
        <div className="fade-in">
            <header className="view-header">
                <div>
                    <h1 className="view-title">Department Progress Dashboard</h1>
                    <p style={{ color: 'var(--text-muted)', fontWeight: 600 }}>Department-wise learner analytics built from SharePoint Enrollments, the synced Learner Management directory, and assessment activity.</p>
                </div>
                <div style={{ display: 'flex', gap: '0.75rem', flexWrap: 'wrap' }}>
                    <button className="btn-secondary" onClick={() => setSelectedDepartment('')}>
                        Show All Departments
                    </button>
                </div>
            </header>

            <div className="stats-grid" style={{ marginBottom: '2rem' }}>
                {reportCategories.map((category, index) => (
                    <div key={index} className="stat-card">
                        <div className="stat-icon" style={{ background: 'var(--primary-light)', color: 'var(--primary)' }}>{category.icon}</div>
                        <div className="stat-content">
                            <span className="stat-label">{category.title}</span>
                            <div className="stat-value">{category.value}</div>
                            <span className="stat-trend">{category.trend}</span>
                        </div>
                    </div>
                ))}
            </div>

            <div style={{ display: 'flex', gap: '1rem', marginBottom: '1.5rem', flexWrap: 'wrap' }}>
                <div className="search-box-unified" style={{ flex: '1 1 320px', maxWidth: '420px', height: '56px' }}>
                    <Filter size={18} />
                    <select
                        value={selectedDepartment}
                        onChange={(event) => setSelectedDepartment(event.target.value)}
                        style={{ border: 'none', background: 'transparent', width: '100%', fontWeight: 700, color: '#1e293b', outline: 'none' }}
                    >
                        <option value="">All Departments</option>
                        {departmentDashboard.map((department) => (
                            <option key={department.department} value={department.department}>{department.department}</option>
                        ))}
                    </select>
                </div>
                <div className="search-box-unified" style={{ flex: '1 1 280px', maxWidth: '320px', height: '56px' }}>
                    <TrendingUp size={18} />
                    <select
                        value={sortMode}
                        onChange={(event) => setSortMode(event.target.value)}
                        style={{ border: 'none', background: 'transparent', width: '100%', fontWeight: 700, color: '#1e293b', outline: 'none' }}
                    >
                        <option value="progress-desc">Sort by Progress: High to Low</option>
                        <option value="progress-asc">Sort by Progress: Low to High</option>
                        <option value="name-asc">Sort by Learner Name</option>
                        <option value="department-asc">Sort by Department</option>
                    </select>
                </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: '1.25rem', marginBottom: '2rem' }}>
                {visibleDepartments.map((department) => (
                    <button
                        key={department.department}
                        type="button"
                        onClick={() => setSelectedDepartment((current) => current === department.department ? '' : department.department)}
                        style={{
                            textAlign: 'left',
                            border: selectedDepartment === department.department ? '2px solid var(--primary)' : '1px solid #e2e8f0',
                            background: 'white',
                            borderRadius: '24px',
                            padding: '1.5rem',
                            boxShadow: '0 10px 25px -15px rgba(15, 23, 42, 0.2)',
                            cursor: 'pointer'
                        }}
                    >
                        <div style={{ fontSize: '1rem', fontWeight: 900, color: '#0f172a', marginBottom: '1rem' }}>{department.department}</div>
                        <div style={{ display: 'grid', gap: '0.5rem' }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', color: '#475569', fontWeight: 700 }}>
                                <span>Total Learners</span>
                                <span style={{ color: '#0f172a' }}>{department.totalLearners}</span>
                            </div>
                            <div style={{ display: 'flex', justifyContent: 'space-between', color: '#475569', fontWeight: 700 }}>
                                <span>Enrolled</span>
                                <span style={{ color: '#2563eb' }}>{department.enrolledCount} ({department.enrolledPercent}%)</span>
                            </div>
                            <div style={{ display: 'flex', justifyContent: 'space-between', color: '#475569', fontWeight: 700 }}>
                                <span>Completed</span>
                                <span style={{ color: '#059669' }}>{department.completedCount} ({department.completedPercent}%)</span>
                            </div>
                            <div style={{ display: 'flex', justifyContent: 'space-between', color: '#475569', fontWeight: 700 }}>
                                <span>In Progress</span>
                                <span style={{ color: '#2563eb' }}>{department.inProgressCount}</span>
                            </div>
                        </div>
                    </button>
                ))}
                {visibleDepartments.length === 0 && (
                    <div className="chart-container" style={{ padding: '2rem', color: '#94a3b8', fontWeight: 700 }}>
                        No department progress records found.
                    </div>
                )}
            </div>

            <div className="table-container">
                <table className="admin-table">
                    <thead>
                        <tr>
                            <th>Department</th>
                            <th>Learner</th>
                            <th>Learning Activity</th>
                            <th>Progress</th>
                            <th>Status</th>
                            <th>Upcoming Renewals</th>
                        </tr>
                    </thead>
                    <tbody>
                        {drillDownLearners.length > 0 ? (
                            drillDownLearners.map((learner: IDepartmentProgressLearner, index: number) => {
                                const learnerState = getLearnerProgressState(learner);
                                const progress = Number(learner.progress || 0);
                                const learnerRenewalSummary = upcomingRenewalSummaryByLearner.get((learner.learnerEmail || '').toString().trim().toLowerCase());
                                const renewalTone = learnerRenewalSummary?.nextRenewal
                                    ? getRenewalTone(Number(learnerRenewalSummary.nextRenewal.daysUntilRenewal || 0))
                                    : null;

                                return (
                                    <tr key={`${learner.learnerEmail}-${learner.path}-${index}`}>
                                        <td style={{ fontWeight: 800, color: '#1e293b' }}>{learner.department || 'Not Available'}</td>
                                        <td>
                                            <button
                                                type="button"
                                                onClick={() => setSelectedLearner(learner)}
                                                style={{
                                                    border: 'none',
                                                    background: 'transparent',
                                                    padding: 0,
                                                    margin: 0,
                                                    fontWeight: 850,
                                                    color: 'var(--primary)',
                                                    cursor: 'pointer',
                                                    textAlign: 'left'
                                                }}
                                            >
                                                {learner.learner || 'Not Available'}
                                            </button>
                                            <div style={{ fontSize: '0.75rem', color: '#94a3b8', fontWeight: 700 }}>{learner.learnerEmail || 'No email'}</div>
                                        </td>
                                        <td style={{ fontWeight: 700, color: '#334155' }}>
                                            <div>{learner.path || 'No certifications assigned'}</div>
                                            <div style={{ fontSize: '0.75rem', color: '#64748b', fontWeight: 700, marginTop: '0.35rem' }}>
                                                {Number(learner.pathCount || 0)} certification path(s) • {Number(learner.assessmentCount || 0)} assessment(s)
                                            </div>
                                        </td>
                                        <td style={{ width: '180px' }}>
                                            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                                                <div style={{ flex: 1, height: '6px', background: '#f1f5f9', borderRadius: '10px', overflow: 'hidden' }}>
                                                    <div
                                                        style={{
                                                            width: `${Math.max(0, Math.min(progress, 100))}%`,
                                                            height: '100%',
                                                            background: learnerState === 'completed' ? '#10b981' : learnerState === 'in-progress' ? '#2563eb' : '#f59e0b'
                                                        }}
                                                    />
                                                </div>
                                                <span style={{ fontSize: '0.75rem', fontWeight: 900, color: '#1e293b' }}>{progress}%</span>
                                            </div>
                                        </td>
                                        <td>
                                                <span
                                                    className="pill"
                                                    style={{
                                                        backgroundColor: learnerState === 'completed' ? '#ecfdf5' : '#eff6ff',
                                                        color: learnerState === 'completed' ? '#059669' : '#2563eb',
                                                        fontWeight: 900
                                                    }}
                                                >
                                                    {learnerState === 'completed' ? 'COMPLETED' : 'IN PROGRESS'}
                                                </span>
                                        </td>
                                        <td>
                                            {learnerRenewalSummary?.nextRenewal ? (
                                                <div style={{ display: 'grid', gap: '0.4rem' }}>
                                                    <span
                                                        className="pill"
                                                        style={{
                                                            backgroundColor: renewalTone?.background,
                                                            color: renewalTone?.color,
                                                            fontWeight: 900,
                                                            border: `1px solid ${renewalTone?.border}`
                                                        }}
                                                    >
                                                        {learnerRenewalSummary.count} renewal{learnerRenewalSummary.count === 1 ? '' : 's'} due
                                                    </span>
                                                    <div style={{ fontSize: '0.75rem', color: '#475569', fontWeight: 700 }}>
                                                        {formatReportDate(learnerRenewalSummary.nextRenewal.renewalDate)}
                                                    </div>
                                                    <div style={{ fontSize: '0.72rem', color: renewalTone?.color, fontWeight: 800 }}>
                                                        {Number(learnerRenewalSummary.nextRenewal.daysUntilRenewal || 0) <= 0
                                                            ? 'Due today'
                                                            : `${learnerRenewalSummary.nextRenewal.daysUntilRenewal} day(s) left`}
                                                    </div>
                                                </div>
                                            ) : (
                                                <span style={{ color: '#94a3b8', fontWeight: 700 }}>No renewals due</span>
                                            )}
                                        </td>
                                    </tr>
                                );
                            })
                        ) : (
                            <tr>
                                <td colSpan={6} style={{ textAlign: 'center', padding: '4rem', color: '#94a3b8', fontWeight: 700 }}>
                                    No learners match the current department view.
                                </td>
                            </tr>
                        )}
                    </tbody>
                </table>
            </div>

            <div className="chart-container" style={{ padding: '1.5rem', marginTop: '2rem' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', gap: '1rem', alignItems: 'center', marginBottom: '1rem', flexWrap: 'wrap' }}>
                    <div>
                        <div style={{ fontWeight: 900, color: '#0f172a' }}>Upcoming Renewals</div>
                        <div style={{ color: '#64748b', fontWeight: 700, fontSize: '0.82rem' }}>
                            Completed certifications with renewal dates due in the next 30 days.
                        </div>
                    </div>
                    <span style={{ fontSize: '0.8rem', fontWeight: 900, color: '#ea580c', background: '#fff7ed', border: '1px solid #fed7aa', borderRadius: '999px', padding: '0.45rem 0.85rem' }}>
                        {visibleUpcomingRenewals.length} renewal due soon
                    </span>
                </div>

                {upcomingRenewalsLoading ? (
                    <div style={{ color: '#64748b', fontWeight: 700 }}>Loading upcoming renewals...</div>
                ) : upcomingRenewalsError ? (
                    <div style={{ color: '#ef4444', fontWeight: 700 }}>{upcomingRenewalsError}</div>
                ) : visibleUpcomingRenewals.length > 0 ? (
                    <div style={{ display: 'grid', gap: '0.9rem' }}>
                        {visibleUpcomingRenewals.map((item) => {
                            const renewalTone = getRenewalTone(Number(item.daysUntilRenewal || 0));

                            return (
                                <div
                                    key={`renewal-${item.id}-${item.learnerEmail}`}
                                    style={{
                                        border: `1px solid ${renewalTone.border}`,
                                        borderRadius: '18px',
                                        padding: '1rem 1.1rem',
                                        background: '#fff'
                                    }}
                                >
                                    <div style={{ display: 'flex', justifyContent: 'space-between', gap: '1rem', alignItems: 'flex-start', flexWrap: 'wrap' }}>
                                        <div>
                                            <div style={{ fontWeight: 900, color: '#0f172a' }}>{item.learnerName || item.learnerEmail || 'Learner'}</div>
                                            <div style={{ color: '#64748b', fontWeight: 700, fontSize: '0.8rem', marginTop: '0.2rem' }}>
                                                {item.department} • {item.learnerEmail || 'No email'}
                                            </div>
                                        </div>
                                        <span
                                            className="pill"
                                            style={{
                                                backgroundColor: renewalTone.background,
                                                color: renewalTone.color,
                                                border: `1px solid ${renewalTone.border}`,
                                                fontWeight: 900
                                            }}
                                        >
                                            Renewal Due Soon
                                        </span>
                                    </div>

                                    <div style={{ marginTop: '0.9rem', fontWeight: 850, color: '#1e293b' }}>
                                        {item.title || item.certId || item.examCode || 'Certification'}
                                    </div>

                                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(170px, 1fr))', gap: '0.85rem', marginTop: '0.85rem' }}>
                                        <div>
                                            <div style={{ color: '#94a3b8', fontSize: '0.72rem', fontWeight: 800, textTransform: 'uppercase' }}>Exam Date</div>
                                            <div style={{ color: '#0f172a', fontWeight: 800 }}>{formatReportDate(item.examDate)}</div>
                                        </div>
                                        <div>
                                            <div style={{ color: '#94a3b8', fontSize: '0.72rem', fontWeight: 800, textTransform: 'uppercase' }}>Renewal Date</div>
                                            <div style={{ color: '#0f172a', fontWeight: 800 }}>{formatReportDate(item.renewalDate)}</div>
                                        </div>
                                        <div>
                                            <div style={{ color: '#94a3b8', fontSize: '0.72rem', fontWeight: 800, textTransform: 'uppercase' }}>Exam Code</div>
                                            <div style={{ color: '#0f172a', fontWeight: 800 }}>{item.examCode || item.certId || 'Not Available'}</div>
                                        </div>
                                        <div>
                                            <div style={{ color: '#94a3b8', fontSize: '0.72rem', fontWeight: 800, textTransform: 'uppercase' }}>Window</div>
                                            <div style={{ color: renewalTone.color, fontWeight: 900 }}>
                                                {Number(item.daysUntilRenewal || 0) <= 0 ? 'Due today' : `${item.daysUntilRenewal} day(s) remaining`}
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                ) : (
                    <div style={{ color: '#94a3b8', fontWeight: 700 }}>
                        No completed certifications have renewal dates due within the next 30 days.
                    </div>
                )}
            </div>

            {selectedLearner && (
                <div
                    className="modal-overlay"
                    style={{ backgroundColor: 'rgba(15, 23, 42, 0.4)', backdropFilter: 'blur(10px)', zIndex: 1000 }}
                    onClick={() => setSelectedLearner(null)}
                >
                    <div
                        className="chart-container modal-card"
                        style={{ maxWidth: '900px', padding: '2rem', borderRadius: '28px' }}
                        onClick={(event) => event.stopPropagation()}
                    >
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: '1rem', marginBottom: '1.5rem' }}>
                            <div>
                                <h2 style={{ margin: 0, fontSize: '1.75rem', fontWeight: 900, color: '#0f172a' }}>{selectedLearner.learner || 'Learner Activity'}</h2>
                                <p style={{ margin: '0.4rem 0 0 0', color: '#64748b', fontWeight: 700 }}>
                                    {selectedLearner.department || 'Not Available'} • {selectedLearner.learnerEmail || 'No email'}
                                </p>
                            </div>
                            <button className="btn-icon" onClick={() => setSelectedLearner(null)} title="Close activity view">
                                <X size={18} />
                            </button>
                        </div>

                        <div className="stats-grid" style={{ marginBottom: '1.5rem' }}>
                            <div className="stat-card">
                                <div className="stat-content">
                                    <span className="stat-label">Certification Paths</span>
                                    <div className="stat-value">{selectedLearnerCertifications.length}</div>
                                    <span className="stat-trend">Assigned in Enrollments</span>
                                </div>
                            </div>
                            <div className="stat-card">
                                <div className="stat-content">
                                    <span className="stat-label">Assessment Attempts</span>
                                    <div className="stat-value">{selectedLearnerAssessments.length}</div>
                                    <span className="stat-trend">Assigned assessments</span>
                                </div>
                            </div>
                            <div className="stat-card">
                                <div className="stat-content">
                                    <span className="stat-label">Completed Paths</span>
                                    <div className="stat-value">{Number(selectedLearner.completedPathCount || 0)}</div>
                                    <span className="stat-trend">Out of {Number(selectedLearner.pathCount || 0)}</span>
                                </div>
                            </div>
                            <div className="stat-card">
                                <div className="stat-content">
                                    <span className="stat-label">Average Progress</span>
                                    <div className="stat-value">{Number(selectedLearner.progress || 0)}%</div>
                                    <span className="stat-trend">{getLearnerProgressState(selectedLearner).replace('-', ' ')}</span>
                                </div>
                            </div>
                        </div>

                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '1rem' }}>
                            <div className="chart-container" style={{ padding: '1.5rem' }}>
                                <div style={{ fontWeight: 900, color: '#0f172a', marginBottom: '1rem' }}>Certification Paths</div>
                                {selectedLearnerCertifications.length > 0 ? (
                                    <div style={{ display: 'grid', gap: '0.85rem' }}>
                                        {selectedLearnerCertifications.map((item, index) => {
                                            const status = getLearnerProgressState(item);
                                            return (
                                                <div key={`${item.learnerEmail}-${item.path}-${index}`} style={{ border: '1px solid #e2e8f0', borderRadius: '18px', padding: '1rem' }}>
                                                    <div style={{ display: 'flex', justifyContent: 'space-between', gap: '0.75rem', alignItems: 'center' }}>
                                                        <div style={{ fontWeight: 800, color: '#1e293b' }}>{item.path || 'Not Available'}</div>
                                                        <span
                                                            className="pill"
                                                            style={{
                                                                backgroundColor: status === 'completed' ? '#ecfdf5' : '#eff6ff',
                                                                color: status === 'completed' ? '#059669' : '#2563eb',
                                                                fontWeight: 900
                                                            }}
                                                        >
                                                            {status === 'completed' ? 'COMPLETED' : 'IN PROGRESS'}
                                                        </span>
                                                    </div>
                                                    <div style={{ marginTop: '0.75rem', display: 'flex', alignItems: 'center', gap: '10px' }}>
                                                        <div style={{ flex: 1, height: '6px', background: '#f1f5f9', borderRadius: '999px', overflow: 'hidden' }}>
                                                            <div
                                                                style={{
                                                                    width: `${Math.max(0, Math.min(Number(item.progress || 0), 100))}%`,
                                                                    height: '100%',
                                                                    background: status === 'completed' ? '#10b981' : '#2563eb'
                                                                }}
                                                            />
                                                        </div>
                                                        <span style={{ fontSize: '0.75rem', fontWeight: 900, color: '#1e293b' }}>{Number(item.progress || 0)}%</span>
                                                    </div>
                                                </div>
                                            );
                                        })}
                                    </div>
                                ) : (
                                    <div style={{ color: '#94a3b8', fontWeight: 700 }}>No certification activity found.</div>
                                )}
                            </div>

                            <div className="chart-container" style={{ padding: '1.5rem' }}>
                                <div style={{ fontWeight: 900, color: '#0f172a', marginBottom: '1rem' }}>Assessment Activity</div>
                                {assessmentAssignmentsLoading ? (
                                    <div style={{ color: '#64748b', fontWeight: 700 }}>Loading assessment activity...</div>
                                ) : assessmentAssignmentsError ? (
                                    <div style={{ color: '#ef4444', fontWeight: 700 }}>{assessmentAssignmentsError}</div>
                                ) : selectedLearnerAssessments.length > 0 ? (
                                    <div style={{ display: 'grid', gap: '0.85rem' }}>
                                        {selectedLearnerAssessments.map((assignment) => (
                                            <div key={assignment.id} style={{ border: '1px solid #e2e8f0', borderRadius: '18px', padding: '1rem' }}>
                                                <div style={{ display: 'flex', justifyContent: 'space-between', gap: '0.75rem', alignItems: 'center' }}>
                                                    <div style={{ fontWeight: 800, color: '#1e293b' }}>
                                                        {assignment.title || assignment.assessmentName || 'Assessment'}
                                                    </div>
                                                    <span className="pill" style={{ backgroundColor: '#eff6ff', color: '#2563eb', fontWeight: 900 }}>
                                                        ASSIGNED
                                                    </span>
                                                </div>
                                                <div style={{ marginTop: '0.5rem', color: '#64748b', fontWeight: 700, fontSize: '0.82rem' }}>
                                                    Scheduled: {assignment.scheduledDate ? new Date(assignment.scheduledDate).toLocaleDateString() : 'Not scheduled'}
                                                </div>
                                                <div style={{ marginTop: '0.25rem', color: '#64748b', fontWeight: 700, fontSize: '0.82rem' }}>
                                                    Created: {assignment.created ? new Date(assignment.created).toLocaleString() : 'Not Available'}
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                ) : (
                                    <div style={{ color: '#94a3b8', fontWeight: 700 }}>No assessment activity found.</div>
                                )}
                            </div>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}

function UsersView({ allUsers, setAllUsers, taxonomyData, setShowAddUserModal, userEmail, updateAdminNotifications, seatManagedCerts, context, realEnrollments, onEnrollmentsChanged, onCertificationCountsChanged, directorySyncState }: any) {
    const [searchTerm, setSearchTerm] = React.useState('');
    const [filterDept, setFilterDept] = React.useState('');
    const [editingUser, setEditingUser] = React.useState<any>(null);
    const [showEditModal, setShowEditModal] = React.useState(false);
    const [assigningUser, setAssigningUser] = React.useState<any>(null);
    const [showAssignModal, setShowAssignModal] = React.useState(false);
    const [selectedPath, setSelectedPath] = React.useState('');
    const [assignedStartDate, setAssignedStartDate] = React.useState(new Date().toISOString().split('T')[0]);
    const [assignedEndDate, setAssignedEndDate] = React.useState(new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]);
    const [totalLearnerCount, setTotalLearnerCount] = React.useState<number | null>(null);
    const adminDisplayName =
        SharePointService.getCurrentContextUserName() ||
        context?.pageContext?.user?.displayName ||
        'Admin';
    const adminUserId = SharePointService.getCurrentContextUserId() ||
        Number(context?.pageContext?.legacyPageContext?.userId || 0) ||
        undefined;

    React.useEffect(() => {
        let isMounted = true;

        const loadTotalLearnerCount = async () => {
            try {
                const count = await SharePointService.getLearnerDirectoryCount();
                if (isMounted) {
                    setTotalLearnerCount(count);
                }
            } catch (error) {
                console.error('[Learners] Failed to load learner directory count', error);
                if (isMounted) {
                    setTotalLearnerCount((allUsers || []).length);
                }
            }
        };

        void loadTotalLearnerCount();

        return () => {
            isMounted = false;
        };
    }, [allUsers.length]);

    const resolvedTotalLearnerCount = totalLearnerCount ?? (allUsers || []).length;

    const filtered = (allUsers || []).filter((u: any) =>
        (u.name.toLowerCase().includes(searchTerm.toLowerCase()) || (u.employeeId && u.employeeId.toLowerCase().includes(searchTerm.toLowerCase()))) &&
        (filterDept === '' || u.department === filterDept) &&
        (u.role && (
            u.role.toLowerCase().includes('member') ||
            u.role.toLowerCase().includes('owner') ||
            u.role.toLowerCase().includes('visitor')
        ))
    );

    const availablePathOptions = React.useMemo(() => {
        const seen = new Set<string>();
        return (seatManagedCerts || []).filter((cert: any) => {
            const key = (cert?.name || '').toString().trim().toLowerCase();
            if (!key || seen.has(key)) {
                return false;
            }

            seen.add(key);
            return true;
        });
    }, [seatManagedCerts]);

    // Note: toggleStatus and handleDeleteUser removed â€” activate/deactivate/delete not shown in UI

    const handleUpdateUser = () => {
        const updated = allUsers.map((u: any) => u.id === editingUser.id ? editingUser : u);
        setAllUsers(updated);

        const audit = JSON.parse(localStorage.getItem('lmsAuditLogs') || '[]');
        audit.unshift({ id: Date.now(), user: userEmail || 'Admin', action: 'EDIT', detail: `Updated profile for ${editingUser.email}`, timestamp: new Date().toISOString() });
        localStorage.setItem('lmsAuditLogs', JSON.stringify(audit.slice(0, 50)));

        setShowEditModal(false);
        setEditingUser(null);
    };

    const handleDirectAssign = async () => {
        if (!selectedPath) return;
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const examDate = new Date(assignedEndDate);

        if (!assignedEndDate || Number.isNaN(examDate.getTime()) || examDate < today) {
            alert('Please select a future exam date.');
            return;
        }

        const certObj = (availablePathOptions || []).find((c: any) => c.name === selectedPath) || { name: selectedPath, code: 'CERT-CUST', pathId: '' };
        const assignmentDetails = await resolveCertificationAssignmentDetails(certObj);

        try {
            await SharePointService.createEnrollmentForCertificationAssignment({
                userEmail: assigningUser.email,
                userName: assigningUser.name,
                certCode: assignmentDetails.certCode,
                certName: assignmentDetails.certName,
                pathId: assignmentDetails.pathId,
                assignedByName: adminDisplayName,
                assignedById: adminUserId,
                examScheduledDate: assignedEndDate
            });
        } catch (error) {
            const errorMessage = getEnrollmentAssignmentErrorMessage(error);
            console.error("Failed to sync assignment with SharePoint:", error);
            alert(errorMessage);
            return;
        }

        if (updateAdminNotifications) {
            await updateAdminNotifications({
                id: Date.now() + 1,
                title: 'Direct Assignment',
                text: `Assigned ${assignmentDetails.certName} to ${assigningUser.name}`,
                time: 'Just now',
                type: 'info',
                targetEmail: 'Admin'
            });
        }

        if (onEnrollmentsChanged) {
            await onEnrollmentsChanged();
        }

        if (onCertificationCountsChanged) {
            await onCertificationCountsChanged(true);
        }

        dispatchEnrollmentRefreshSignal();

        window.setTimeout(() => {
            window.dispatchEvent(new Event(LMS_AUDIT_REFRESH_EVENT));
        }, 500);

        alert("Assignment pushed successfully!");
        setShowAssignModal(false);
        setAssigningUser(null);
        setSelectedPath('');
        setAssignedStartDate(new Date().toISOString().split('T')[0]);
        setAssignedEndDate(new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]);
    };



    return (
        <div className="fade-in">
            <div className="view-header">
                <div>
                    <h1 className="view-title">Learner Management</h1>
                    <p style={{ color: 'var(--text-muted)', fontWeight: 600 }}>Manage employees, learning paths, and organizational roles.</p>
                    <p style={{ color: '#64748b', fontWeight: 700, fontSize: '0.8rem', marginTop: '0.5rem' }}>
                        Source: SharePoint Learners directory, with SharePoint PeopleManager enrichment for designation and department.
                    </p>
                    <p style={{ color: '#64748b', fontWeight: 700, fontSize: '0.8rem', marginTop: '0.35rem' }}>
                        Any SharePoint Owner or Member with admin access can assign certifications from this tab.
                    </p>
                </div>
                <div className="header-actions" style={{ display: 'flex', gap: '1rem' }}>
                    <span style={{ fontSize: '0.85rem', fontWeight: 800, color: '#0f172a', background: '#f8fafc', padding: '0.5rem 1rem', borderRadius: '12px', display: 'flex', alignItems: 'center', gap: '8px', border: '1px solid #e2e8f0' }}>
                        <Users size={16} /> Total Learners: {resolvedTotalLearnerCount}
                    </span>
                    <span style={{ fontSize: '0.85rem', fontWeight: 800, color: 'var(--primary)', background: 'var(--primary-light)', padding: '0.5rem 1rem', borderRadius: '12px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <ShieldCheck size={16} /> REAL-TIME CONNECTED
                    </span>
                </div>
            </div>

            {directorySyncState.error && (
                <div style={{ marginBottom: '1rem', padding: '1rem 1.25rem', borderRadius: '16px', border: '1px solid #fecaca', background: '#fef2f2', color: '#b91c1c', fontWeight: 700 }}>
                    {directorySyncState.error}
                </div>
            )}

            {directorySyncState.loading && allUsers.length > 0 && (
                <div style={{ marginBottom: '1rem', padding: '0.9rem 1rem', borderRadius: '14px', border: '1px solid #bfdbfe', background: '#eff6ff', color: '#1d4ed8', fontWeight: 700 }}>
                    Refreshing SharePoint users from the default site groups without clearing the current table.
                </div>
            )}

            <div className="search-box-unified" style={{ marginBottom: '2rem', width: '100%', height: '56px' }}>
                <Search size={20} />
                <input
                    type="text"
                    placeholder="Search by name or employee ID..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                />
                <div style={{ width: '1.5px', height: '24px', background: '#e2e8f0', margin: '0 0.5rem' }} />
                <Filter size={18} style={{ color: '#94a3b8' }} />
                <select
                    style={{ border: 'none', background: 'transparent', fontWeight: 700, outline: 'none', padding: '0 1rem', color: '#1e293b' }}
                    value={filterDept}
                    onChange={(e) => setFilterDept(e.target.value)}
                >
                    <option value="">All Departments</option>
                    {taxonomyData.departments.map((d: string) => <option key={d} value={d}>{d}</option>)}
                </select>
            </div>

            <div className="table-container">
                <table className="admin-table">
                    <thead>
                        <tr>
                            <th>Employee Info</th>
                            <th>Corporate Identity</th>
                            <th>Designation</th>
                            <th>Department</th>
                            <th>Progress</th>
                            <th style={{ textAlign: 'right' }}>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
                        {directorySyncState.loading && filtered.length === 0 && (
                            <tr>
                                <td colSpan={6} style={{ padding: '2rem', textAlign: 'center', color: '#475569', fontWeight: 700 }}>
                                    Loading users from the SharePoint Learners directory source...
                                </td>
                            </tr>
                        )}
                        {!directorySyncState.loading && filtered.length === 0 && (
                            <tr>
                                <td colSpan={6} style={{ padding: '2rem', textAlign: 'center', color: '#64748b', fontWeight: 700 }}>
                                    {directorySyncState.error || ((searchTerm || filterDept) && allUsers.length > 0
                                        ? 'No users match the current filters.'
                                        : 'No users found in the SharePoint Learners directory source.')}
                                </td>
                            </tr>
                        )}
                        {filtered.map((user: any) => (
                            <tr key={user.id}>
                                <td>
                                    <div className="user-info-cell" style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
                                        <div className="user-avatar-small" style={{ width: '40px', height: '40px', borderRadius: '12px', background: 'var(--primary-light)', color: 'var(--primary)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 900 }}>
                                            {user.name.charAt(0)}
                                        </div>
                                        <div style={{ display: 'flex', flexDirection: 'column' }}>
                                            <span style={{ fontWeight: 850, color: '#1e293b' }}>{user.name}</span>
                                            <span style={{ fontSize: '0.7rem', color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.05em', fontWeight: 800 }}>{user.employeeId}</span>
                                        </div>
                                    </div>
                                </td>
                                <td style={{ color: '#64748b', fontWeight: 600 }}>{user.email}</td>
                                <td style={{ color: '#1e293b', fontWeight: 700 }}>{user.jobTitle || 'Not Available'}</td>
                                <td style={{ color: '#1e293b', fontWeight: 700 }}>{user.department || 'Not Available'}</td>
                                <td style={{ width: '180px' }}>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                                        <div style={{ flex: 1, height: '6px', background: '#f1f5f9', borderRadius: '10px', overflow: 'hidden' }}>
                                            <div style={{ width: `${user.progress}%`, height: '100%', background: 'var(--gradient-primary)' }} />
                                        </div>
                                        <span style={{ fontSize: '0.75rem', fontWeight: 900, color: '#1e293b' }}>{user.progress}%</span>
                                    </div>
                                </td>
                                <td style={{ textAlign: 'right' }}>
                                    <div className="action-btns">
                                        <button className="btn-icon" title="Edit" onClick={() => { setEditingUser(user); setShowEditModal(true); }}><Edit size={16} /></button>
                                        <button
                                            className="btn-icon"
                                            title="Assign Certification"
                                            style={{ color: 'var(--primary)' }}
                                            onClick={() => {
                                                setAssigningUser(user);
                                                setAssignedStartDate(new Date().toISOString().split('T')[0]);
                                                setAssignedEndDate(new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]);
                                                setShowAssignModal(true);
                                            }}
                                        ><Award size={16} /></button>
                                    </div>
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>

            {/* Edit User Modal */}
            {showEditModal && editingUser && (
                <div className="modal-overlay" style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(15, 23, 42, 0.4)', backdropFilter: 'blur(10px)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                    <div className="modal-card fade-in" style={{ backgroundColor: 'white', padding: '2.5rem', borderRadius: '32px', width: '100%', maxWidth: '500px', boxShadow: '0 25px 50px -12px rgba(0,0,0,0.25)' }}>
                        <h2 style={{ fontSize: '1.8rem', fontWeight: 950, color: '#1e293b', marginBottom: '1.5rem' }}>Edit Learner <span style={{ color: 'var(--primary)' }}>Profile</span></h2>
                        <div style={{ display: 'grid', gap: '1.2rem' }}>
                            <div>
                                <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Full Name</label>
                                <input className="admin-input" value={editingUser.name} onChange={(e) => setEditingUser({ ...editingUser, name: e.target.value })} />
                            </div>
                            <div className="responsive-two-column-grid">
                                <div>
                                    <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Department</label>
                                    <select className="admin-input" value={editingUser.department} onChange={(e) => setEditingUser({ ...editingUser, department: e.target.value })}>
                                        {taxonomyData.departments.map((d: any) => <option key={d} value={d}>{d}</option>)}
                                    </select>
                                </div>
                                <div>
                                    <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Role</label>
                                    <select className="admin-input" value={editingUser.role} onChange={(e) => setEditingUser({ ...editingUser, role: e.target.value })}>
                                        {taxonomyData.roles.map((r: any) => <option key={r} value={r}>{r}</option>)}
                                    </select>
                                </div>
                            </div>
                            <div>
                                <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Corporate Email</label>
                                <input className="admin-input" value={editingUser.email} disabled style={{ backgroundColor: '#f8fafc', color: '#94a3b8' }} />
                            </div>
                            <div className="responsive-two-column-grid">
                                <div>
                                    <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Employee ID</label>
                                    <input className="admin-input" value={editingUser.employeeId || ''} disabled style={{ backgroundColor: '#f8fafc', color: '#94a3b8' }} />
                                </div>
                                <div>
                                    <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Site Group</label>
                                    <input className="admin-input" value={editingUser.siteGroup || ''} disabled style={{ backgroundColor: '#f8fafc', color: '#94a3b8' }} />
                                </div>
                            </div>
                        </div>
                        <div style={{ display: 'flex', gap: '1rem', marginTop: '2.5rem' }}>
                            <button className="btn-secondary" style={{ flex: 1 }} onClick={() => setShowEditModal(false)}>Cancel</button>
                            <button className="btn-primary" style={{ flex: 2 }} onClick={handleUpdateUser}>Save Changes</button>
                        </div>
                    </div>
                </div>
            )}

            {/* Assign Cert Modal */}
            {showAssignModal && assigningUser && (
                <div className="modal-overlay" style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(15, 23, 42, 0.4)', backdropFilter: 'blur(10px)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                    <div className="modal-card fade-in" style={{ backgroundColor: 'white', padding: '2.5rem', borderRadius: '32px', width: '100%', maxWidth: '480px', boxShadow: '0 25px 50px -12px rgba(0,0,0,0.25)' }}>
                        <h2 style={{ fontSize: '1.8rem', fontWeight: 950, color: '#1e293b', marginBottom: '0.5rem' }}>Assign <span style={{ color: 'var(--primary)' }}>Certification</span></h2>
                        <p style={{ color: '#64748b', fontWeight: 600, fontSize: '0.95rem', marginBottom: '1.5rem' }}>Directly enroll <b>{assigningUser.name}</b> in a path.</p>

                        <div>
                            <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Select Path</label>
                            <select className="admin-input" value={selectedPath} onChange={(e) => setSelectedPath(e.target.value)}>
                                <option value="">Select Path...</option>
                                <optgroup label="Core Microsoft Paths">
                                    {availablePathOptions.map((c: any) => (
                                        <option key={c.pathId || c.id || c.code || c.name} value={c.name}>{c.name}</option>
                                    ))}
                                </optgroup>
                            </select>
                        </div>

                        <div className="responsive-two-column-grid" style={{ gap: '1rem', marginTop: '1.2rem' }}>
                            <div>
                                <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Assigned Date</label>
                                <input type="date" className="admin-input" value={assignedStartDate} onChange={e => setAssignedStartDate(e.target.value)} disabled />
                            </div>
                            <div>
                                <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', textTransform: 'uppercase' }}>Exam Scheduled Date</label>
                                <input type="date" className="admin-input" value={assignedEndDate} min={assignedStartDate} onChange={e => setAssignedEndDate(e.target.value)} />
                            </div>
                        </div>
                        <div style={{ marginTop: '1rem', padding: '0.9rem 1rem', background: '#f8fafc', borderRadius: '16px', border: '1px solid #e2e8f0', color: '#475569', fontSize: '0.85rem', fontWeight: 700 }}>
                            Employee ID: {assigningUser.employeeId || 'Not available from directory'}<br />
                            Department: {assigningUser.department || 'Not available'}<br />
                            Site Group: {assigningUser.siteGroup || 'Not available'}
                        </div>
                        <div style={{ display: 'flex', gap: '1rem', marginTop: '2.5rem' }}>
                            <button className="btn-secondary" style={{ flex: 1 }} onClick={() => { setShowAssignModal(false); setAssigningUser(null); }}>Cancel</button>
                            <button className="btn-primary" style={{ flex: 2 }} onClick={handleDirectAssign}>Assign with Exam Date</button>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}

function TaxonomyView({ taxonomyData, setTaxonomyData, activeTab, setActiveTab, setShowTaxonomyModal }: any) {
    const tabs = [
        { id: 'departments', label: 'Departments', icon: <Briefcase size={18} /> },
        { id: 'businessUnits', label: 'Units', icon: <Hash size={18} /> },
        { id: 'roles', label: 'Roles', icon: <Users size={18} /> },
        { id: 'locations', label: 'Locations', icon: <MapPin size={18} /> },
        { id: 'groups', label: 'Groups', icon: <Tag size={18} /> }
    ];

    const currentItems = taxonomyData[activeTab] || [];

    const deleteItem = async (item: string) => {
        if (!window.confirm(`Are you sure you want to remove "${item}"?`)) return;
        const updatedItems = currentItems.filter((i: string) => i !== item);
        const updated = {
            ...taxonomyData,
            [activeTab]: updatedItems
        };
        setTaxonomyData(updated);
        localStorage.setItem('lmsTaxonomyData', JSON.stringify(updated));

        // Sync to SharePoint
        try {
            await SharePointService.updateTaxonomy(activeTab as string, updatedItems);
        } catch (e) {
            console.warn("Taxonomy deletion sync failed", e);
        }
    };

    return (
        <div className="fade-in">
            <div className="view-header">
                <div>
                    <h1 className="view-title">Taxonomy & Meta</h1>
                    <p style={{ color: 'var(--text-muted)', fontWeight: 600 }}>Manage the organizational structural data.</p>
                </div>
                <button className="btn-primary" onClick={() => setShowTaxonomyModal(true)}><PlusCircle size={20} /> Add Item</button>
            </div>

            <div style={{ display: 'flex', gap: '0.4rem', marginBottom: '2.5rem', background: '#f1f5f9', padding: '0.4rem', borderRadius: '18px', width: 'fit-content' }}>
                {tabs.map(tab => (
                    <button
                        key={tab.id}
                        onClick={() => setActiveTab(tab.id)}
                        style={{
                            display: 'flex',
                            alignItems: 'center',
                            gap: '8px',
                            padding: '0.65rem 1.1rem',
                            borderRadius: '14px',
                            border: 'none',
                            background: activeTab === tab.id ? 'white' : 'transparent',
                            color: activeTab === tab.id ? 'var(--primary)' : '#64748b',
                            fontWeight: 800,
                            fontSize: '0.85rem',
                            cursor: 'pointer',
                            boxShadow: activeTab === tab.id ? '0 4px 12px rgba(0,0,0,0.05)' : 'none',
                            transition: 'all 0.2s'
                        }}
                    >
                        {tab.icon}
                        {tab.label}
                    </button>
                ))}
            </div>

            <div className="table-container">
                <table className="admin-table">
                    <thead>
                        <tr>
                            <th>Item Name</th>
                            <th>Metadata Label</th>
                            <th style={{ textAlign: 'right' }}>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
                        {currentItems.map((item: string, idx: number) => (
                            <tr key={idx}>
                                <td style={{ fontWeight: 850, color: '#1e293b', fontSize: '1rem' }}>{item}</td>
                                <td><span className="pill approved" style={{ fontSize: '0.65rem', padding: '0.2rem 0.6rem' }}>{activeTab.toUpperCase()}</span></td>
                                <td style={{ textAlign: 'right' }}>
                                    <div className="action-btns">
                                        <button className="btn-icon"><Edit size={16} /></button>
                                        <button className="btn-icon" style={{ color: '#ef4444' }} onClick={() => deleteItem(item)}><Trash2 size={16} /></button>
                                    </div>
                                </td>
                            </tr>
                        ))}
                        {currentItems.length === 0 && (
                            <tr>
                                <td colSpan={3} style={{ textAlign: 'center', padding: '4rem', color: '#94a3b8', fontWeight: 700 }}>No items defined in this category.</td>
                            </tr>
                        )}
                    </tbody>
                </table>
            </div>
        </div>
    );
}

function StatusBadge({ status, label }: { status: string, label?: string }) {
    return <span className={`status-badge ${status?.toLowerCase() || 'open'}`}>{label || status || 'OPEN'}</span>;
}

function ContentLibraryView({ showToast, userEmail, context, updateAdminNotifications }: any) {
    const contentFolders = ['sales', 'presales', 'engineering'];
    const normalizedCurrentUserEmail = (userEmail || SharePointService.getCurrentContextUserEmail() || '').toString().trim().toLowerCase();

    const [assets, setAssets] = useState<any[]>([]);
    const [showAddModal, setShowAddModal] = useState(false);
    const [editingAsset, setEditingAsset] = useState<any>(null);
    const [searchTerm, setSearchTerm] = useState('');
    const [filterType, setFilterType] = useState('ALL');
    const [isSavingAsset, setIsSavingAsset] = useState(false);
    const [newAsset, setNewAsset] = useState<any>({ name: '', type: 'VIDEO', owner: normalizedCurrentUserEmail || 'system.admin@skysecure.com', uploadedBy: normalizedCurrentUserEmail || 'system.admin@skysecure.com', status: 'Uploaded', description: '', size: '', path: '', fileObject: null, folderName: 'sales', newFolderName: '' });
    const fileInputRef = useRef<any>(null);
    const editFileInputRef = useRef<any>(null);

    const dispatchContentLibraryRefresh = (): void => {
        window.dispatchEvent(new Event(LMS_CONTENT_LIBRARY_REFRESH_EVENT));
        try {
            localStorage.setItem('lmsContentLibrary:lastRefresh', new Date().toISOString());
        } catch (error) {
            console.warn('[ContentLibrary] Failed to persist the refresh token.', error);
        }
    };

    const loadContentAssets = async (): Promise<any[]> => {
        try {
            const spAssets = await SharePointService.getContentAssets();
            const normalizedAssets = spAssets || [];
            setAssets(normalizedAssets);

            try {
                if (normalizedAssets.length > 0) {
                    localStorage.setItem('lmsContentLibrary', JSON.stringify(normalizedAssets));
                } else {
                    localStorage.removeItem('lmsContentLibrary');
                }
            } catch (storageError) {
                console.warn('[ContentLibrary] Failed to update cached asset snapshot.', storageError);
            }

            return normalizedAssets;
        } catch (error) {
            console.error("Error loading SharePoint assets:", error);
            return [];
        }
    };

    const handleFileSelect = (e: any, setAssetData: any, currentData: any) => {
        const file = e.target.files[0];
        if (file) {
            const bytes = file.size;
            const sizeStr = bytes < 1024 * 1024 ? (bytes / 1024).toFixed(1) + ' KB' : (bytes / (1024 * 1024)).toFixed(1) + ' MB';
            setAssetData({ ...currentData, fileObject: file, path: file.name, size: sizeStr });
        }
    };

    useEffect(() => {
        void loadContentAssets();

        const storageListener = (event: StorageEvent) => {
            if (!event.key || event.key === 'lmsContentLibrary:lastRefresh') {
                void loadContentAssets();
            }
        };
        const refreshListener = () => {
            void loadContentAssets();
        };
        const intervalId = window.setInterval(() => {
            void loadContentAssets();
        }, 5000);
        window.addEventListener('storage', storageListener);
        window.addEventListener(LMS_CONTENT_LIBRARY_REFRESH_EVENT, refreshListener);
        return () => {
            window.clearInterval(intervalId);
            window.removeEventListener('storage', storageListener);
            window.removeEventListener(LMS_CONTENT_LIBRARY_REFRESH_EVENT, refreshListener);
        };
    }, []);

    const handleSave = async () => {
        if (isSavingAsset) {
            return;
        }

        if (!newAsset.name) {
            alert("Name is required.");
            return;
        }

        const finalFolderName = (
            newAsset.folderName === '__new__'
                ? newAsset.newFolderName
                : newAsset.folderName
        ).toString().trim();

        if (newAsset.type !== 'LINK' && !newAsset.fileObject && !newAsset.path) {
            alert("Please attach a file using the 'Click to Choose File' area.");
            return;
        }

        if (newAsset.fileObject && !finalFolderName) {
            alert("Folder name is required.");
            return;
        }

        try {
            setIsSavingAsset(true);

            let assetUrl = '';
            let uploadedAssetDetails: any = null;
            if (newAsset.path && newAsset.path.startsWith('http')) {
                assetUrl = newAsset.path;
            }

            if (newAsset.fileObject && (!context?.spHttpClient || !context?.pageContext?.web?.absoluteUrl)) {
                if (showToast) {
                    showToast('SharePoint site context is unavailable. Open this admin web part on a SharePoint site page and try again.', 'error');
                }
                return;
            }

            // Upload to SharePoint if a file is selected
            if (newAsset.fileObject) {
                try {
                    if (showToast) showToast('Connecting to SharePoint Documents1 storage...', 'info');

                    const uploaded = await SharePointService.uploadFileToDocuments1(newAsset.fileObject, finalFolderName);
                    assetUrl = uploaded.url;
                    uploadedAssetDetails = {
                        name: uploaded.name,
                        path: uploaded.serverRelativeUrl,
                        url: uploaded.url,
                        folderName: uploaded.folderName || finalFolderName
                    };

                    const refreshedDocuments1Files = await SharePointService.getFilesFromDocuments1();
                    const refreshedUploadedFile = refreshedDocuments1Files.find((file: any) =>
                        (file.path || '').toLowerCase() === uploaded.serverRelativeUrl.toLowerCase() ||
                        (file.name || '').toLowerCase() === uploaded.name.toLowerCase()
                    );

                    if (refreshedUploadedFile) {
                        uploadedAssetDetails = refreshedUploadedFile;
                        assetUrl = refreshedUploadedFile.url || assetUrl;
                    }

                    if (showToast) showToast(`Upload successful! File stored in Documents1/${uploadedAssetDetails?.folderName || finalFolderName}.`, 'success');
                } catch (err: any) {
                    console.error('Error during upload to Documents1', err);
                    if (showToast) {
                        showToast(`Error uploading to Documents1: ${err.message}`, 'error');
                    }
                    return; // Abort saving if upload failed
                }
            }

            const finalAsset = {
                ...newAsset,
                name: uploadedAssetDetails?.name || newAsset.name,
                type: uploadedAssetDetails?.type || newAsset.type,
                url: assetUrl,
                id: Date.now(),
                dateAdded: uploadedAssetDetails?.dateAdded || new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }),
                size: newAsset.type === 'LINK' ? 'N/A' : (uploadedAssetDetails?.size || newAsset.size || (Math.random() * 50 + 1).toFixed(1) + ' MB'),
                path: uploadedAssetDetails?.path || newAsset.path,
                folderName: uploadedAssetDetails?.folderName || finalFolderName,
                uploadedBy: normalizedCurrentUserEmail || (newAsset.owner || '').toString().trim().toLowerCase(),
                owner: normalizedCurrentUserEmail || (newAsset.owner || '').toString().trim().toLowerCase(),
                status: 'Uploaded'
            };
            delete finalAsset.fileObject; // Don't persist File object to localStorage

            // Sync metadata to SharePoint ContentLibrary list
            try {
                await SharePointService.addContentAsset(finalAsset);
            } catch (error) {
                console.error("Error syncing asset metadata to ContentLibrary:", error);
                if (uploadedAssetDetails) {
                    try {
                        await SharePointService.deleteContentAsset(uploadedAssetDetails);
                    } catch (rollbackError) {
                        console.error("Error rolling back uploaded file after metadata failure:", rollbackError);
                    }
                }
                if (showToast) {
                    showToast(`Upload failed while saving ContentLibrary metadata: ${(error as Error)?.message || 'Unknown error'}`, 'error');
                }
                return;
            }

            await loadContentAssets();

            // Add to Admin Notifications
            const newAdminNotif = {
                id: Date.now(),
                title: 'New Content Deployed',
                text: `Asset "${finalAsset.name}" is now available in the content library.`,
                time: 'Just now',
                type: 'success'
            };
            if (updateAdminNotifications) updateAdminNotifications(newAdminNotif);
            dispatchContentLibraryRefresh();

            setShowAddModal(false);
            setNewAsset({ name: '', type: 'PDF', owner: normalizedCurrentUserEmail || 'system.admin@skysecure.com', uploadedBy: normalizedCurrentUserEmail || 'system.admin@skysecure.com', status: 'Uploaded', description: '', size: '', path: '', fileObject: null, folderName: 'sales', newFolderName: '' } as any);
            if (showToast) showToast(`Asset "${finalAsset.name}" successfully deployed and synced.`);
        } catch (refreshError) {
            console.error('Error refreshing content assets after upload', refreshError);
            if (showToast) {
                showToast(`Asset upload completed, but refresh failed. ${(refreshError as Error)?.message || ''}`.trim(), 'error');
            }
        } finally {
            setIsSavingAsset(false);
        }
    };

    const handleUpdate = async () => {
        if (!editingAsset?.name) {
            return;
        }

        const updatedAsset = {
            ...editingAsset,
            uploadedBy: (editingAsset.uploadedBy || normalizedCurrentUserEmail || editingAsset.owner || '').toString().trim().toLowerCase(),
            owner: (editingAsset.uploadedBy || normalizedCurrentUserEmail || editingAsset.owner || '').toString().trim().toLowerCase(),
            status: 'Updated'
        };

        try {
            await SharePointService.addContentAsset(updatedAsset);
            await loadContentAssets();
            dispatchContentLibraryRefresh();
            setEditingAsset(null);
            if (showToast) showToast(`Asset "${updatedAsset.name}" updated successfully.`, "info");
        } catch (error) {
            console.error("ContentLibrary update failed", error);
            if (showToast) {
                showToast(`Failed to update "${updatedAsset.name}".`, "error");
            }
        }
    };

    const handleDelete = async (asset: any) => {
        if (!window.confirm("Permanently archive this asset?")) return;

        try {
            await SharePointService.deleteContentAsset(asset);
            await loadContentAssets();
            dispatchContentLibraryRefresh();
            if (showToast) showToast("Asset archived successfully.", "info");
        } catch (e) {
            console.warn("Archive sync failed", e);
            if (showToast) {
                showToast(`Failed to archive "${asset?.name || 'asset'}". ${(e as Error)?.message || ''}`.trim(), "error");
            }
        }
    };

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
        let finalUrl = asset.url;

        // Handle failed uploads or legacy relative paths
        if (finalUrl && !finalUrl.startsWith('http') && !finalUrl.startsWith('/')) {
            finalUrl = ''; // Treat as missing if it's just a file name
        }

        // Fallback for legacy items before the SP upload feature
        if (!finalUrl && asset.path) {
            if (asset.path.startsWith('http') || asset.path.startsWith('/')) {
                finalUrl = asset.path;
            } else {
                alert("Cannot open document. This is an unsupported legacy record. Wait for the admin to re-upload the valid source URL.");
                return;
            }
        }

        if (!finalUrl) {
            alert("This document cannot be previewed because no valid source URL was found.");
            return;
        }

        const targetUrl = buildAssetActionUrl(finalUrl, openInSharePointViewer);
        if (!targetUrl) {
            alert("This document cannot be previewed because no valid source URL was found.");
            return;
        }

        window.open(targetUrl, '_blank', 'noopener,noreferrer');
    };

    const getIcon = (type: string) => {
        switch (type) {
            case 'VIDEO': return <Video size={18} />;
            case 'PDF': return <FileText size={18} />;
            case 'EXCEL': return <FileSpreadsheet size={18} />;
            case 'DOC': return <FileCode size={18} />;
            case 'PPT': return <Presentation size={18} />;
            case 'SCORM': return <FileArchive size={18} />;
            case 'LINK': return <Link2 size={18} />;
            default: return <Folder size={18} />;
        }
    };

    const filteredAssets = assets.filter(a => {
        const matchesSearch = (a.name || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
            (a.description || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
            (a.folderName || '').toLowerCase().includes(searchTerm.toLowerCase());
        const matchesType = filterType === 'ALL' || a.type === filterType;
        return matchesSearch && matchesType;
    });

    return (
        <div className="fade-in">
            <header className="view-header">
                <div>
                    <h1 className="view-title">Content Intelligence</h1>
                    <p style={{ color: 'var(--text-muted)', fontWeight: 600 }}>Centralized repository for organizational learning assets.</p>
                </div>
                <div style={{ display: 'flex', gap: '1rem' }}>
                    <button className="btn-secondary" onClick={() => { void loadContentAssets().then(() => { if (showToast) showToast('Content library synced from SharePoint.', 'success'); }); }}><Activity size={18} /> Sync SP</button>
                    <button className="btn-primary" onClick={() => setShowAddModal(true)} style={{ background: 'linear-gradient(135deg, #2563eb, #7c3aed)', border: 'none' }}>
                        <Upload size={18} /> Upload from PC / Laptop
                    </button>
                </div>
            </header>

            <div className="search-box-unified" style={{ marginTop: '2rem', height: '60px', width: '100%', maxWidth: 'none', display: 'flex', alignItems: 'center', padding: '0 1.5rem', background: 'white', borderRadius: '20px', border: '1.5px solid var(--border)', boxShadow: '0 4px 20px -5px rgba(0,0,0,0.05)' }}>
                <Search size={22} color="#94a3b8" />
                <input
                    type="text"
                    placeholder="Search by asset name, metadata tags, or description..."
                    value={searchTerm}
                    onChange={e => setSearchTerm(e.target.value)}
                    style={{ background: 'transparent', border: 'none', flex: 1, padding: '0 1rem', fontSize: '1rem', fontWeight: 700, outline: 'none' }}
                />
                <div style={{ width: '2px', height: '30px', background: '#e2e8f0', margin: '0 1.5rem' }} />
                <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                    {['ALL', 'VIDEO', 'PDF', 'EXCEL', 'DOC', 'PPT', 'SCORM', 'LINK'].map(t => (
                        <button
                            key={t}
                            onClick={() => setFilterType(t)}
                            style={{
                                padding: '0.5rem 1rem', borderRadius: '12px', fontSize: '0.75rem', fontWeight: 900,
                                background: filterType === t ? 'var(--primary)' : 'white',
                                color: filterType === t ? 'white' : '#64748b',
                                border: '1.5px solid', borderColor: filterType === t ? 'var(--primary)' : '#e2e8f0',
                                cursor: 'pointer', transition: 'all 0.2s'
                            }}
                        >
                            {t}
                        </button>
                    ))}
                </div>
            </div>

            <div className="table-container" style={{ marginTop: '2rem', background: 'white', borderRadius: '32px', padding: '1.5rem', border: '1.5px solid var(--border)' }}>
                <table className="admin-table">
                    <thead>
                        <tr>
                            <th>ASSET IDENTITY</th>
                            <th>TYPE & SIZE</th>
                            <th>OWNER / DATE</th>
                            <th>STATUS</th>
                            <th style={{ textAlign: 'right' }}>OPERATIONS</th>
                        </tr>
                    </thead>
                    <tbody>
                        {filteredAssets.map((asset) => (
                            <tr key={asset.id} style={{ transition: 'all 0.2s' }}>
                                <td>
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                        <div style={{ fontWeight: 950, color: '#1e293b', fontSize: '1.05rem', display: 'flex', alignItems: 'center', gap: '8px' }}>
                                            {asset.name}
                                            {asset.path && <span style={{ fontSize: '0.6rem', background: '#f1f5f9', color: '#64748b', padding: '2px 6px', borderRadius: '4px', fontWeight: 700 }}>{asset.path}</span>}
                                            {asset.folderName && <span style={{ fontSize: '0.65rem', background: '#eff6ff', color: '#2563eb', padding: '2px 8px', borderRadius: '999px', fontWeight: 800, textTransform: 'capitalize' }}>{asset.folderName}</span>}
                                        </div>
                                        <div style={{ fontSize: '0.8rem', color: '#64748b', fontWeight: 600, maxWidth: '300px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{asset.description || 'No description provided.'}</div>
                                    </div>
                                </td>
                                <td>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                                        <div style={{ width: '40px', height: '40px', background: '#f8fafc', borderRadius: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--primary)' }}>
                                            {getIcon(asset.type)}
                                        </div>
                                        <div>
                                            <div style={{ fontWeight: 900, fontSize: '0.85rem' }}>{asset.type}</div>
                                            <div style={{ fontSize: '0.75rem', color: '#94a3b8', fontWeight: 700 }}>{asset.size} {asset.duration ? `â€¢ ${asset.duration}` : ''}</div>
                                        </div>
                                    </div>
                                </td>
                                <td>
                                    <div style={{ display: 'flex', flexDirection: 'column' }}>
                                        <div style={{ fontWeight: 800, color: '#475569', fontSize: '0.85rem' }}>{asset.owner}</div>
                                        <div style={{ display: 'flex', alignItems: 'center', gap: '4px', fontSize: '0.75rem', color: '#94a3b8', fontWeight: 600 }}>
                                            <Clock size={12} /> {asset.dateAdded}
                                        </div>
                                    </div>
                                </td>
                                <td>
                                    <StatusBadge status={asset.status === 'Updated' ? 'completed' : 'active'} label={(asset.status || 'Uploaded').toUpperCase()} />
                                </td>
                                <td style={{ textAlign: 'right' }}>
                                    <div className="action-btns" style={{ justifyContent: 'flex-end' }}>
                                        <button className="btn-icon" onClick={() => setEditingAsset(asset)} title="Edit Properties"><Edit size={16} /></button>
                                        <button className="btn-icon" onClick={() => handleOpenAsset(asset, true)} title="Preview Content"><Eye size={16} /></button>
                                        <button className="btn-icon" onClick={() => handleOpenAsset(asset, false)} title="Open Source"><Download size={16} /></button>
                                        <button className="btn-icon delete-btn" onClick={() => handleDelete(asset)} style={{ color: '#ef4444' }}><Trash2 size={16} /></button>
                                    </div>
                                </td>
                            </tr>
                        ))}
                        {filteredAssets.length === 0 && (
                            <tr>
                                <td colSpan={5} style={{ textAlign: 'center', padding: '6rem 0', color: '#94a3b8' }}>
                                    <div style={{ marginBottom: '1.5rem', opacity: 0.3 }}><FileQuestion size={64} /></div>
                                    <div style={{ fontWeight: 950, fontSize: '1.2rem', color: '#1e293b' }}>No Learning Assets Found</div>
                                    <div style={{ fontSize: '0.95rem', fontWeight: 600, marginTop: '0.5rem' }}>Adjust your filters or upload a new asset to the repository.</div>
                                </td>
                            </tr>
                        )}
                    </tbody>
                </table>
            </div>

            {/* Content Library Modals (Add / Edit) */}
            {showAddModal && (
                <div className="modal-overlay" style={{
                    position: 'fixed', inset: 0, backgroundColor: 'rgba(15, 23, 42, 0.4)', backdropFilter: 'blur(12px)', zIndex: 2000,
                    display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '1.5rem'
                }}>
                    <div className="modal-card fade-in" style={{
                        backgroundColor: 'white', padding: '3rem', borderRadius: '36px', width: '100%', maxWidth: '560px',
                        boxShadow: '0 25px 90px -15px rgba(0,0,0,0.3)', border: '1.5px solid var(--border)'
                    }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '2.5rem' }}>
                            <div>
                                <h2 style={{ fontSize: '2rem', fontWeight: 950, color: '#1e293b', margin: 0, letterSpacing: '-0.04em' }}>Deploy <span style={{ color: 'var(--primary)' }}>New Asset</span></h2>
                                <p style={{ color: '#64748b', fontSize: '1rem', fontWeight: 600, marginTop: '0.6rem' }}>Configure metadata for the enterprise library.</p>
                            </div>
                            <button onClick={() => setShowAddModal(false)} className="btn-icon"><X size={26} /></button>
                        </div>

                        <div style={{ display: 'grid', gap: '1.75rem' }}>
                            <div className="form-group" style={{ padding: '1.5rem', background: 'rgba(59, 130, 246, 0.03)', borderRadius: '24px', border: '2px dashed var(--primary)' }}>
                                <label style={{ fontSize: '0.85rem', fontWeight: 950, color: 'var(--primary)', marginBottom: '1rem', display: 'flex', alignItems: 'center', gap: '8px', textTransform: 'uppercase' }}>
                                    <Cloud size={18} /> STEP 1: Upload from PC / Laptop
                                </label>
                                <input
                                    type="file"
                                    ref={fileInputRef}
                                    style={{ display: 'none' }}
                                    onChange={(e) => handleFileSelect(e, setNewAsset, newAsset)}
                                />
                                <div
                                    onClick={() => fileInputRef.current.click()}
                                    style={{
                                        borderRadius: '20px',
                                        padding: '1.5rem',
                                        textAlign: 'center',
                                        cursor: 'pointer',
                                        background: 'white',
                                        transition: 'all 0.2s',
                                        display: 'flex',
                                        flexDirection: 'column',
                                        alignItems: 'center',
                                        gap: '12px',
                                        boxShadow: '0 4px 12px rgba(0,0,0,0.05)'
                                    }}
                                >
                                    <Upload size={32} style={{ color: 'var(--primary)' }} />
                                    <div style={{ fontWeight: 950, color: '#1e293b', fontSize: '1.1rem' }}>Click to Choose File</div>
                                    <div style={{ fontSize: '0.8rem', color: '#64748b', fontWeight: 600 }}>Videos, PDFs, or Documents</div>
                                </div>
                                {newAsset.path && (
                                    <div style={{ marginTop: '1rem', padding: '0.9rem', background: '#ecfdf5', borderRadius: '14px', border: '1px solid #10b981', display: 'flex', alignItems: 'center', gap: '8px', color: '#065f46', fontWeight: 900, fontSize: '0.85rem' }}>
                                        <CheckCircle2 size={18} /> SELECTED: {newAsset.path}
                                    </div>
                                )}
                            </div>

                            <div className="form-group">
                                <label style={{ fontSize: '0.85rem', fontWeight: 900, color: '#475569', marginBottom: '0.8rem', textTransform: 'uppercase' }}>STEP 2: Asset Details</label>
                                <input
                                    type="text"
                                    className="input-field"
                                    placeholder="Asset Display Name (e.g. MS-900 Study Guide)"
                                    value={newAsset.name}
                                    onChange={e => setNewAsset({ ...newAsset, name: e.target.value })}
                                    style={{ width: '100%', padding: '1.1rem', borderRadius: '16px', border: '2px solid #e2e8f0', outline: 'none', background: '#f8fafc', fontWeight: 700 }}
                                />
                                <textarea
                                    className="input-field"
                                    placeholder="Brief description of asset purpose..."
                                    value={newAsset.description}
                                    onChange={e => setNewAsset({ ...newAsset, description: e.target.value })}
                                    style={{ width: '100%', padding: '1.1rem', borderRadius: '16px', border: '2px solid #e2e8f0', outline: 'none', background: '#f8fafc', fontWeight: 700, height: '80px', marginTop: '1rem', resize: 'none' }}
                                />
                                <div style={{ marginTop: '1rem' }}>
                                    <label style={{ fontSize: '0.8rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', display: 'block', textTransform: 'uppercase' }}>Destination Folder</label>
                                    <select
                                        value={newAsset.folderName}
                                        onChange={e => setNewAsset({ ...newAsset, folderName: e.target.value })}
                                        style={{ width: '100%', padding: '1.1rem', borderRadius: '16px', border: '2px solid #e2e8f0', outline: 'none', background: '#f8fafc', fontWeight: 700, textTransform: 'capitalize' }}
                                    >
                                        {contentFolders.map(folder => (
                                            <option key={folder} value={folder}>
                                                {folder}
                                            </option>
                                        ))}
                                        <option value="__new__">+ Create New Folder</option>
                                    </select>
                                    {newAsset.folderName === '__new__' && (
                                        <input
                                            type="text"
                                            className="input-field"
                                            placeholder="Enter new folder name"
                                            value={newAsset.newFolderName}
                                            onChange={e => setNewAsset({ ...newAsset, newFolderName: e.target.value })}
                                            style={{ width: '100%', padding: '1.1rem', borderRadius: '16px', border: '2px solid #e2e8f0', outline: 'none', background: '#f8fafc', fontWeight: 700, marginTop: '0.75rem' }}
                                        />
                                    )}
                                </div>
                                <div style={{ marginTop: '1rem' }}>
                                    <label style={{ fontSize: '0.8rem', fontWeight: 800, color: '#64748b', marginBottom: '0.5rem', display: 'block', textTransform: 'uppercase' }}>Manual Source Path (Optional)</label>
                                    <input
                                        type="text"
                                        className="input-field"
                                        placeholder="e.g. /shared/docs/guide.pdf"
                                        value={newAsset.path}
                                        onChange={e => setNewAsset({ ...newAsset, path: e.target.value })}
                                        style={{ width: '100%', padding: '1.1rem', borderRadius: '16px', border: '2px solid #e2e8f0', outline: 'none', background: '#f8fafc', fontWeight: 700 }}
                                    />
                                </div>
                            </div>

                            <div className="form-group">
                                <label style={{ fontSize: '0.85rem', fontWeight: 900, color: '#475569', marginBottom: '0.8rem', textTransform: 'uppercase' }}>Classification</label>
                                <div className="responsive-asset-type-grid">
                                    {['VIDEO', 'PDF', 'EXCEL', 'DOC', 'PPT', 'SCORM', 'LINK'].map(type => (
                                        <button
                                            key={type}
                                            onClick={() => setNewAsset({ ...newAsset, type })}
                                            style={{
                                                padding: '1rem', borderRadius: '20px', border: '2px solid',
                                                borderColor: newAsset.type === type ? 'var(--primary)' : '#e2e8f0',
                                                background: newAsset.type === type ? 'rgba(59, 130, 246, 0.05)' : 'white',
                                                color: newAsset.type === type ? 'var(--primary)' : '#64748b',
                                                fontWeight: 900, fontSize: '0.9rem', cursor: 'pointer', transition: '0.2s', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px'
                                            }}
                                        >
                                            {getIcon(type)} {type}
                                        </button>
                                    ))}
                                </div>
                            </div>

                                <div className="form-group">
                                    <label style={{ fontSize: '0.85rem', fontWeight: 900, color: '#475569', marginBottom: '0.8rem', textTransform: 'uppercase' }}>Visibility</label>
                                    <div style={{ display: 'flex', gap: '1.5rem' }}>
                                        {['Uploaded', 'Updated'].map(s => (
                                            <label key={s} style={{
                                                display: 'flex', alignItems: 'center', gap: '12px', cursor: 'pointer', fontWeight: 800, fontSize: '1rem', flex: 1, padding: '1.1rem', background: '#f8fafc', borderRadius: '20px', border: '2px solid',
                                                borderColor: newAsset.status === s ? 'var(--primary)' : '#e2e8f0'
                                        }}>
                                            <input
                                                type="radio"
                                                name="add-status"
                                                checked={newAsset.status === s}
                                                onChange={() => setNewAsset({ ...newAsset, status: s })}
                                                style={{ scale: '1.3' }}
                                            />
                                            {s}
                                        </label>
                                    ))}
                                </div>
                            </div>

                            <button className="btn-primary" onClick={handleSave} disabled={isSavingAsset} style={{ width: '100%', padding: '1.3rem', justifyContent: 'center', fontSize: '1.1rem', borderRadius: '20px', boxShadow: '0 10px 25px -5px rgba(59, 130, 246, 0.3)', opacity: isSavingAsset ? 0.7 : 1, cursor: isSavingAsset ? 'not-allowed' : 'pointer' }}>
                                <Upload size={22} /> {isSavingAsset ? 'Uploading...' : 'Deploy Asset to Library'}
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {/* Edit Asset Modal */}
            {
                editingAsset && (
                    <div className="modal-overlay" style={{
                        position: 'fixed', inset: 0, backgroundColor: 'rgba(15, 23, 42, 0.5)', backdropFilter: 'blur(16px)', zIndex: 2000,
                        display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '1.5rem'
                    }}>
                        <div className="modal-card fade-in" style={{
                            backgroundColor: 'white', padding: '3rem', borderRadius: '40px', width: '100%', maxWidth: '600px',
                            boxShadow: '0 30px 100px -20px rgba(0,0,0,0.4)', border: '1.5px solid var(--border)'
                        }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '3rem' }}>
                                <div>
                                    <h2 style={{ fontSize: '2.1rem', fontWeight: 950, color: '#1e293b', margin: 0, letterSpacing: '-0.04em' }}>Update <span style={{ color: 'var(--primary)' }}>Identity</span></h2>
                                    <p style={{ color: '#64748b', fontSize: '1.05rem', fontWeight: 600, marginTop: '0.6rem' }}>Refine asset metadata and architectural positioning.</p>
                                </div>
                                <button onClick={() => setEditingAsset(null)} className="btn-icon"><X size={28} /></button>
                            </div>

                            <div style={{ display: 'grid', gap: '2.5rem' }}>
                                <div className="form-group">
                                    <label style={{ fontSize: '0.9rem', fontWeight: 950, color: '#334155', marginBottom: '1rem', textTransform: 'uppercase' }}>Asset Properties</label>
                                    <div style={{ display: 'grid', gap: '1.25rem' }}>
                                        <input
                                            type="text"
                                            className="input-field"
                                            value={editingAsset.name}
                                            onChange={e => setEditingAsset({ ...editingAsset, name: e.target.value })}
                                            style={{ width: '100%', padding: '1.2rem', borderRadius: '18px', border: '2px solid #e2e8f0', background: '#f8fafc', fontWeight: 800, fontSize: '1.05rem' }}
                                        />
                                        <textarea
                                            className="input-field"
                                            value={editingAsset.description}
                                            onChange={e => setEditingAsset({ ...editingAsset, description: e.target.value })}
                                            style={{ width: '100%', padding: '1.2rem', borderRadius: '18px', border: '2px solid #e2e8f0', background: '#f8fafc', fontWeight: 700, height: '100px', resize: 'none' }}
                                        />
                                        <div style={{ marginTop: '0.5rem' }}>
                                            <label style={{ fontSize: '0.8rem', fontWeight: 850, color: '#64748b', marginBottom: '0.6rem', display: 'block' }}>SOURCE PATH / FILENAME</label>
                                            <input
                                                type="text"
                                                className="input-field"
                                                value={editingAsset.path || ''}
                                                onChange={e => setEditingAsset({ ...editingAsset, path: e.target.value })}
                                                style={{ width: '100%', padding: '1rem', borderRadius: '16px', border: '2px solid #e2e8f0', background: '#f8fafc', fontWeight: 800 }}
                                            />
                                        </div>
                                    </div>
                                </div>

                                <div className="form-group">
                                    <label style={{ fontSize: '0.9rem', fontWeight: 950, color: '#334155', marginBottom: '1rem', textTransform: 'uppercase' }}>Replacement Media</label>
                                    <input
                                        type="file"
                                        ref={editFileInputRef}
                                        style={{ display: 'none' }}
                                        onChange={(e) => handleFileSelect(e, setEditingAsset, editingAsset)}
                                    />
                                    <div
                                        onClick={() => editFileInputRef.current.click()}
                                        style={{
                                            border: '2px dashed #cbd5e1',
                                            borderRadius: '24px',
                                            padding: '2rem',
                                            textAlign: 'center',
                                            cursor: 'pointer',
                                            background: '#f8fafc',
                                            transition: 'all 0.2s',
                                            display: 'flex',
                                            flexDirection: 'column',
                                            alignItems: 'center',
                                            gap: '12px'
                                        }}
                                    >
                                        <Cloud size={32} style={{ color: '#94a3b8' }} />
                                        <div style={{ fontWeight: 850, color: '#475569' }}>Click to replace current asset</div>
                                    </div>
                                </div>

                                <div className="responsive-two-column-grid" style={{ gap: '2rem' }}>
                                    <div className="form-group">
                                        <label style={{ fontSize: '0.85rem', fontWeight: 900, color: '#64748b', marginBottom: '0.8rem' }}>OWNERSHIP</label>
                                        <input type="text" className="input-field" value={editingAsset.owner} disabled style={{ width: '100%', padding: '1rem', borderRadius: '16px', background: '#f1f5f9', opacity: 0.8, fontWeight: 800 }} />
                                    </div>
                                    <div className="form-group">
                                        <label style={{ fontSize: '0.85rem', fontWeight: 900, color: '#64748b', marginBottom: '0.8rem' }}>DEPLOYMENT DATE</label>
                                        <input type="text" className="input-field" value={editingAsset.dateAdded} disabled style={{ width: '100%', padding: '1rem', borderRadius: '16px', background: '#f1f5f9', opacity: 0.8, fontWeight: 800 }} />
                                    </div>
                                </div>

                                <div className="form-group">
                                    <label style={{ fontSize: '0.9rem', fontWeight: 950, color: '#334155', marginBottom: '1rem', textTransform: 'uppercase' }}>Access Control</label>
                                    <div style={{ display: 'flex', gap: '1.5rem' }}>
                                        {['Uploaded', 'Updated'].map(s => (
                                            <label key={s} style={{
                                                display: 'flex', alignItems: 'center', gap: '15px', cursor: 'pointer', fontWeight: 900, fontSize: '1.1rem', flex: 1, padding: '1.25rem', borderRadius: '24px', border: '2.5px solid',
                                                borderColor: editingAsset.status === s ? 'var(--primary)' : 'transparent',
                                                background: editingAsset.status === s ? 'rgba(59, 130, 246, 0.05)' : '#f8fafc'
                                            }}>
                                                <input
                                                    type="radio"
                                                    name="edit-status"
                                                    checked={editingAsset.status === s}
                                                    onChange={() => setEditingAsset({ ...editingAsset, status: s })}
                                                    style={{ scale: '1.5' }}
                                                />
                                                {s}
                                            </label>
                                        ))}
                                    </div>
                                </div>

                                <div style={{ display: 'flex', gap: '1.5rem', marginTop: '1rem' }}>
                                    <button className="btn-secondary" onClick={() => setEditingAsset(null)} style={{ flex: 1, padding: '1.3rem', borderRadius: '20px', fontWeight: 850, fontSize: '1rem' }}>Discard Changes</button>
                                    <button className="btn-primary" onClick={handleUpdate} style={{ flex: 1.5, padding: '1.3rem', borderRadius: '20px', fontWeight: 900, fontSize: '1rem', boxShadow: '0 10px 30px rgba(59, 130, 246, 0.4)' }}>Apply Strategic Updates</button>
                                </div>
                            </div>
                        </div>
                    </div>
                )
            }
        </div >
    );
}

function AssignmentsView({ taxonomyData, allUsers, seatManagedCerts, userEmail, updateAdminNotifications, realEnrollments, onEnrollmentsChanged, onCertificationCountsChanged, context }: any) {
    const [selectedUser, setSelectedUser] = useState('');
    const [selectedPath, setSelectedPath] = useState('');
    const [showAssignModal, setShowAssignModal] = useState(false);
    const [assignmentExamDate, setAssignmentExamDate] = useState(new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]);
    const adminDisplayName = context?.pageContext?.user?.displayName || 'Admin';
    const availablePathOptions = useMemo(() => {
        const seen = new Set<string>();
        return (seatManagedCerts || []).filter((cert: any) => {
            const key = (cert?.name || '').toString().trim().toLowerCase();
            if (!key || seen.has(key)) {
                return false;
            }

            seen.add(key);
            return true;
        });
    }, [seatManagedCerts]);

    const handleAssign = () => {
        if (!selectedUser || !selectedPath) { alert("Select both user and path."); return; }
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const examDate = new Date(assignmentExamDate);
        if (!assignmentExamDate || Number.isNaN(examDate.getTime()) || examDate < today) {
            alert('Select a future exam date.');
            return;
        }
        void (async () => {
            const userObj = allUsers.find((u: any) => u.email === selectedUser) || { name: selectedUser, email: selectedUser };
            const certObj = (availablePathOptions || []).find((c: any) => c.name === selectedPath) || { name: selectedPath, code: 'CERT-CUST', pathId: '' };
            const assignmentDetails = await resolveCertificationAssignmentDetails(certObj);

            try {
                await SharePointService.createEnrollmentForCertificationAssignment({
                    userEmail: userObj.email,
                    userName: userObj.name,
                    certCode: assignmentDetails.certCode,
                    certName: assignmentDetails.certName,
                    pathId: assignmentDetails.pathId,
                    assignedByName: adminDisplayName,
                    examScheduledDate: assignmentExamDate
                });
            } catch (error) {
                const errorMessage = getEnrollmentAssignmentErrorMessage(error);
                console.error('Assignment save failed:', error);
                alert(errorMessage);
                return;
            }

            const newAdminNotif = {
                id: Date.now(),
                title: 'Assignment Pushed',
                text: `Assigned ${assignmentDetails.certName} to ${userObj.name}.`,
                time: 'Just now',
                type: 'info'
            };
            if (updateAdminNotifications) {
                await updateAdminNotifications(newAdminNotif);
            }

            window.setTimeout(() => {
                window.dispatchEvent(new Event(LMS_AUDIT_REFRESH_EVENT));
            }, 500);

            let audit = [];
            try { audit = JSON.parse(localStorage.getItem('lmsAuditLogs') || '[]'); } catch (e) { }
            audit.unshift({ id: Date.now(), user: userEmail || 'Admin', action: 'ASSIGN_CERT', detail: `Assigned ${assignmentDetails.certCode} to ${userObj.email}`, timestamp: new Date().toISOString() });
            localStorage.setItem('lmsAuditLogs', JSON.stringify(audit.slice(0, 50)));

            if (onEnrollmentsChanged) {
                await onEnrollmentsChanged();
            }

            if (onCertificationCountsChanged) {
                await onCertificationCountsChanged(true);
            }

            dispatchEnrollmentRefreshSignal();

            alert("Certification path successfully pushed to " + userObj.name + ". A notification has been sent to their portal.");
            setShowAssignModal(false);
            setSelectedUser('');
            setSelectedPath('');
            setAssignmentExamDate(new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]);
        })();
    };

    return (
        <div className="fade-in">
            <header className="view-header">
                <div>
                    <h1 className="view-title">Assignment Engine</h1>
                    <p style={{ color: 'var(--text-muted)', fontWeight: 600 }}>Create dynamic assignment rules or push manual assignments.</p>
                </div>
                <button className="btn-primary" onClick={() => { setAssignmentExamDate(new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]); setShowAssignModal(true); }}><PlusCircle size={18} /> Push Assignment</button>
            </header>

            {showAssignModal && (
                <div style={{
                    position: 'fixed', inset: 0, backgroundColor: 'rgba(15, 23, 42, 0.4)', backdropFilter: 'blur(12px)', zIndex: 2000,
                    display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '1.5rem'
                }}>
                    <div className="fade-in" style={{
                        backgroundColor: 'white', padding: '2.5rem', borderRadius: '32px', width: '100%', maxWidth: '520px',
                        boxShadow: '0 25px 70px -12px rgba(0,0,0,0.3)', border: '1.5px solid var(--border)'
                    }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '2rem' }}>
                            <div>
                                <h2 style={{ fontSize: '1.85rem', fontWeight: 950, color: '#1e293b', margin: 0, letterSpacing: '-0.04em' }}>Path <span style={{ color: 'var(--primary)' }}>Assignment</span></h2>
                                <p style={{ color: 'var(--text-muted)', fontSize: '0.95rem', fontWeight: 600, marginTop: '0.4rem' }}>{selectedPath ? `Assigning: ${selectedPath}` : "Select a learner and a certification path."}</p>
                            </div>
                            <button onClick={() => setShowAssignModal(false)} className="btn-icon"><X size={24} /></button>
                        </div>

                        <div style={{ display: 'grid', gap: '1.5rem' }}>
                            <div>
                                <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#475569', marginBottom: '0.5rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Target Learner</label>
                                <select className="input-field" value={selectedUser} onChange={e => setSelectedUser(e.target.value)} style={{ width: '100%', padding: '0.8rem' }}>
                                    <option value="">Select a user...</option>
                                    {allUsers
                                        .filter((u: any) => u.role && (u.role.toLowerCase().includes('member') || u.role.toLowerCase().includes('owner')))
                                        .map((u: any) => <option key={u.id} value={u.email}>{u.name} ({u.email})</option>)}
                                </select>
                            </div>

                            <div>
                                <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#475569', marginBottom: '0.5rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Certification Path</label>
                                <select className="input-field" value={selectedPath} onChange={e => setSelectedPath(e.target.value)} style={{ width: '100%', padding: '0.8rem', fontWeight: 750, color: '#1e293b' }}>
                                    <option value="">Select Certification Path...</option>
                                    <optgroup label="Core Microsoft Paths">
                                        {availablePathOptions.map((c: any) => (
                                            <option key={c.pathId || c.id || c.code || c.name} value={c.name}>{c.name}</option>
                                        ))}
                                    </optgroup>
                                </select>
                            </div>

                            <div>
                                <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: '#475569', marginBottom: '0.5rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Exam Scheduled Date</label>
                                <input
                                    type="date"
                                    className="input-field"
                                    value={assignmentExamDate}
                                    min={new Date().toISOString().split('T')[0]}
                                    onChange={e => setAssignmentExamDate(e.target.value)}
                                    style={{ width: '100%', padding: '0.8rem', fontWeight: 750, color: '#1e293b' }}
                                />
                            </div>

                            <button className="btn-primary" onClick={handleAssign} style={{ marginTop: '1rem', width: '100%', justifyContent: 'center', padding: '1.1rem' }}>Assign & Notify</button>
                        </div>
                    </div>
                </div>
            )}

            <div className="stats-grid" style={{ marginBottom: '2rem' }}>
                <div className="stat-card">
                    <div className="stat-content">
                        <span className="stat-label">Active Rules</span>
                        <div className="stat-value">5</div>
                    </div>
                </div>
                <div className="stat-card">
                    <div className="stat-content">
                        <span className="stat-label">Manual Pushes Today</span>
                        <div className="stat-value">12</div>
                    </div>
                </div>
            </div>
            <div className="table-container">
                <table className="admin-table">
                    <thead><tr><th>Rule / Assignment Title</th><th>Target Audience</th><th>Triggers</th><th>Status</th></tr></thead>
                    <tbody>
                        <tr>
                            <td style={{ fontWeight: 800 }}>Onboarding Bootcamp v2</td>
                            <td>All New Hires</td>
                            <td>Role = Trainee</td>
                            <td><StatusBadge status="active" label="Active" /></td>
                        </tr>
                        <tr>
                            <td style={{ fontWeight: 800 }}>Annual Security Compliance '26</td>
                            <td>Global Org</td>
                            <td>Manual Push (Due 30 Nov)</td>
                            <td><StatusBadge status="active" label="Active" /></td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    );
}




