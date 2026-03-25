import { SPHttpClient } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';

export interface IContentAsset {
    id?: number;
    name: string;
    type: string;
    owner: string;
    status: string;
    dateAdded: string;
    size: string;
    description: string;
    url?: string;
    path?: string;
    folderName?: string;
    uploadedBy?: string;
    assignedTo?: string;
}

export interface IEnrollment {
    id?: number;
    userEmail: string;
    userName: string;
    certificationId?: number;
    certCode: string;
    certName: string;
    provider?: string;
    startDate: string;
    endDate: string;
    status: string;
    progress: number;
    certificateName?: string;
    assignedByAdmin?: boolean;
    assignedDate?: string;
    assignedById?: number;
    assignedByName?: string;
    examScheduledDate?: string;
    rescheduledDate?: string;
    expiryDate?: string;
    completionDate?: string;
    examCode?: string;
    listStatus?: string;
    pathId?: string;
    userId?: number;
    category?: string;
    level?: string;
}

export interface INotification {
    id?: number;
    title: string;
    text: string;
    targetEmail: string; // "Admin" or specific user email
    type: string; // "info", "warning", "success", "priority"
    time: string;
    read: boolean;
    status?: string;
    assignedDate?: string;
    sourceList?: string;
}

export interface IAuditLogRecord {
    id: number;
    title: string;
    learnerEmail: string;
    learnerName?: string;
    action: string;
    assignmentName?: string;
    assignmentDate: string;
    assignedById?: number;
    status?: string;
    created: string;
    pathId?: string;
    userId?: number;
    timestamp?: string;
}

export interface IAdminCert {
    id?: number;
    name: string;
    code: string;
    description: string;
    provider: string;
    modules: string;
    targetAudience: string;
    category?: string;
    level?: string;
    pathId?: string;
    maxSeats?: number;
    enrolledCount?: number;
    assignedLearnerCount?: number;
    isSharePointManaged?: boolean;
}

export interface ISiteMembershipSnapshot {
    owners: ILearnerDirectoryUser[];
    members: ILearnerDirectoryUser[];
    visitors: ILearnerDirectoryUser[];
    learners: ILearnerDirectoryUser[];
}

export interface IAdminPortalAccessState extends ISiteMembershipSnapshot {
    currentUserRole: 'Owner' | 'Member' | 'Visitor' | 'Learner' | 'Unknown';
    canAccessAdmin: boolean;
    accessCheckFailed: boolean;
}

export interface ISharePointGroupUser {
    Id: number | string;
    Title: string;
    Email: string;
    LoginName: string;
    id: number | string;
    userId: number | string;
    name: string;
    email: string;
    login: string;
    group: string;
    siteGroup: string;
    role: 'Owner' | 'Member' | 'Visitor';
}

export interface ILearnerDirectoryUser extends ISharePointGroupUser {
    jobTitle?: string;
    department?: string;
}

export interface IAssessmentAssignmentDefinition {
    id?: string | number;
    title: string;
    assessmentName: string;
    certCode?: string;
    threshold?: number;
    questions?: number;
    questionsArr?: any[] | null;
    provider?: string;
    duration?: string;
}

export interface IAssessmentAssignmentRecord {
    id: number;
    title: string;
    userEmail: string;
    userName: string;
    assessmentName: string;
    orderIndex: number;
    assignedGroup: 'Owners' | 'Members' | 'Visitors';
    scheduledDate?: string;
    created?: string;
    assessmentPayload?: IAssessmentAssignmentDefinition | null;
}

export interface IAssessmentTrackerItem {
    id: number;
    learner: string;
    learnerEmail: string;
    assessment: string;
    created: string;
}

export interface IDepartmentProgressLearner {
    learner: string;
    learnerEmail: string;
    path: string;
    progress: number;
    department: string;
    status: string;
    pathCount?: number;
    completedPathCount?: number;
    assessmentCount?: number;
}

export interface IDepartmentProgressSummary {
    department: string;
    totalLearners: number;
    enrolledCount: number;
    completedCount: number;
    inProgressCount: number;
    notStartedCount: number;
    enrolledPercent: number;
    completedPercent: number;
    learners: IDepartmentProgressLearner[];
}

export interface ICertificationAssignmentRecord {
    id: number;
    title: string;
    userEmail: string;
    userName: string;
    certificationName: string;
    certCode?: string;
    certificationId?: number;
    assignedDate: string;
    issuedDate: string;
    expiryDate?: string;
    status?: string;
    orderIndex: number;
    assignedGroup: 'Owners' | 'Members' | 'Visitors';
    created?: string;
}

export interface ICertificationMaxSeatsItem {
    id: number;
    title: string;
    code: string;
    maxSeats: number;
    enrolledCount?: number;
    assignedLearnerCount?: number;
    usedSeats?: number;
    category?: string;
    provider?: string;
    level?: string;
    link?: string;
}

export interface ICertificationImportRow {
    title: string;
    code: string;
    maxSeats?: number;
    provider?: string;
}

export interface ICertificationCompletionRecord {
    id: number;
    title: string;
    certId: string;
    examDate: string;
    renewalDate: string;
    examCode?: string;
    created?: string;
    modified?: string;
    authorEmail?: string;
    authorName?: string;
}

export interface IUpcomingRenewalRecord extends ICertificationCompletionRecord {
    learnerEmail: string;
    learnerName: string;
    daysUntilRenewal: number;
    urgency: 'urgent' | 'soon';
}

interface ICertificationCatalogSchema {
    codeField: string | null;
    maxSeatsField: string | null;
    categoryField: string | null;
    providerField: string | null;
    levelField: string | null;
    linkField: string | null;
    fileUrlField: string | null;
    folderNameField: string | null;
    codeFieldType: string;
    maxSeatsFieldType: string;
    categoryFieldType: string;
    providerFieldType: string;
    levelFieldType: string;
    linkFieldType: string;
    fileUrlFieldType: string;
    folderNameFieldType: string;
}

interface ICertificationCompletionListSchema {
    certIdField: string | null;
    examDateField: string | null;
    renewalDateField: string | null;
    examCodeField: string | null;
}

interface ISharePointListFieldInfo {
    Title?: string;
    InternalName?: string;
    StaticName?: string;
    Hidden?: boolean;
    ReadOnlyField?: boolean;
    TypeAsString?: string;
    FieldTypeKind?: number;
}

interface ICertificationListSchema {
    assignedToField: string;
    userEmailField: string | null;
    userNameField: string | null;
    certificationNameField: string | null;
    scheduledDateField: string;
    assignedDateField: string | null;
    issuedDateField: string | null;
    expiryDateField: string | null;
    statusField: string | null;
    certCodeField: string | null;
    orderIndexField: string | null;
    assignedGroupField: string | null;
}

interface IAssessmentAssignmentListSchema {
    assignedToField: string;
    scheduledDateField: string;
    userEmailField: string | null;
    userNameField: string | null;
    assessmentNameField: string | null;
    orderIndexField: string | null;
    assignedGroupField: string | null;
    assessmentPayloadField: string | null;
}

interface IEnrollmentListSchema {
    assignedToField: string | null;
    assignedByField: string | null;
    certificationLookupField: string | null;
    userEmailField: string | null;
    userNameField: string | null;
    certCodeField: string | null;
    certNameField: string | null;
    startDateField: string | null;
    endDateField: string | null;
    statusField: string | null;
    progressField: string | null;
    certificateNameField: string | null;
    assignedDateField: string | null;
    examScheduledDateField: string | null;
    rescheduledDateField: string | null;
    expiryDateField: string | null;
    completionDateField: string | null;
    examCodeField: string | null;
    assignedByNameField: string | null;
    assignedToEmailField: string | null;
    pathIdField: string | null;
    fieldTypes: Record<string, string>;
}

interface IAuditLogListSchema {
    userField: string | null;
    learnerEmailField: string | null;
    learnerNameField: string | null;
    actionField: string | null;
    assignmentNameField: string | null;
    assignmentDateField: string | null;
    assignedByField: string | null;
    statusField: string | null;
    pathIdField: string | null;
    timestampField: string | null;
}

interface IContentLibraryListSchema {
    fileLinkField: string;
    uploadedByField: string;
    assignedToField: string;
    statusField: string;
    folderNameField: string | null;
    assetTypeField: string | null;
    descriptionField: string | null;
    fileSizeField: string | null;
}

export const LMS_ENROLLMENTS_REFRESH_EVENT = 'lms-enrollments-refresh';
export const LMS_ENROLLMENTS_REFRESH_STORAGE_KEY = 'lms-enrollments-refresh-token';
export const LMS_CONTENT_LIBRARY_REFRESH_EVENT = 'lms-content-library-refresh';

export class SharePointService {
    private static _siteUrl: string;
    private static _spHttpClient: SPHttpClient;
    private static _context: any;
    private static readonly _jsonHeaders: Record<string, string> = {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': ''
    };
    private static readonly _taxonomyListCandidates: string[] = ['LMS Taxonomy', 'LMS_Taxonomy'];
    private static readonly _enrollmentListCandidates: string[] = ['Enrollment', 'Enrollments', 'LMS_Enrollments'];
    private static readonly _userNotificationListCandidates: string[] = ['Notifications', 'LMS_Notifications'];
    private static readonly _assessmentAssignmentListCandidates: string[] = ['Assessment Assignments', 'Asseswment'];
    private static readonly _learningAndSkillsListCandidates: string[] = ['LearningAndSkills', 'Learning and Skills', 'Learning & Skills', 'Learningandks'];
    private static readonly _certificationCompletionListCandidates: string[] = ['uploadlist', 'UploadList', 'Upload List'];
    private static readonly _learnerGroupCandidates: string[] = ['Learners', 'Learner'];
    private static readonly _certificationAssignmentListName = 'Certifications';
    private static readonly _contentLibraryListName = 'ContentLibrary';
    private static readonly _auditLogListName = 'Audit Logs';
    private static readonly _contentLibraryListCandidates: string[] = ['ContentLibrary', 'Content Library'];
    private static readonly _auditLogListCandidates: string[] = ['Audit Logs', 'AuditLogs'];
    private static readonly _defaultThrottleRetryMs = 4000;
    private static readonly _assessmentAssignmentGroups: Array<{
        groupId: number;
        siteGroup: 'Owners' | 'Members' | 'Visitors';
        role: 'Owner' | 'Member' | 'Visitor';
    }> = [
        { groupId: 3, siteGroup: 'Members', role: 'Member' },
        { groupId: 5, siteGroup: 'Visitors', role: 'Visitor' },
        { groupId: 4, siteGroup: 'Owners', role: 'Owner' }
    ];
    private static readonly _listRetryCooldownMs = 60000;
    private static _ensuredLists = new Set<string>();
    private static _pendingListEnsures = new Map<string, Promise<void>>();
    private static _listFailureTimestamps = new Map<string, number>();
    private static _documentLibraryPromise: Promise<any | null> | null = null;
    private static _membershipSnapshotCache: { siteUrl: string; snapshot: ISiteMembershipSnapshot } | null = null;
    private static _membershipSnapshotPromise: Promise<ISiteMembershipSnapshot> | null = null;
    private static _loggedResponseIssues = new Set<string>();
    private static _enrollmentListNameCache: string | null = null;
    private static _userNotificationListNameCache: string | null = null;
    private static _assessmentAssignmentListNameCache: string | null = null;
    private static _contentLibraryListNameCache: string | null = null;
    private static _auditLogListNameCache: string | null = null;
    private static _learningAndSkillsListNameCache: string | null = null;
    private static _certificationCompletionListNameCache: string | null = null;
    private static _listTitlesCache: { siteUrl: string; titles: Set<string>; titleMap: Map<string, string> } | null = null;
    private static _listTitlesPromise: Promise<Set<string>> | null = null;
    private static _formDigestCache: { siteUrl: string; value: string; expiresAt: number } | null = null;
    private static _formDigestPromise: Promise<string> | null = null;
    private static _certificationListSchemaCache: { siteUrl: string; schema: ICertificationListSchema } | null = null;
    private static _certificationListSchemaPromise: Promise<ICertificationListSchema> | null = null;
    private static _certificationCatalogSchemaCache: { siteUrl: string; schema: ICertificationCatalogSchema } | null = null;
    private static _certificationCatalogSchemaPromise: Promise<ICertificationCatalogSchema> | null = null;
    private static _certificationCompletionListSchemaCache: { siteUrl: string; schema: ICertificationCompletionListSchema } | null = null;
    private static _certificationCompletionListSchemaPromise: Promise<ICertificationCompletionListSchema> | null = null;
    private static _assessmentAssignmentListSchemaCache: { siteUrl: string; schema: IAssessmentAssignmentListSchema } | null = null;
    private static _assessmentAssignmentListSchemaPromise: Promise<IAssessmentAssignmentListSchema> | null = null;
    private static _enrollmentListSchemaCache: { siteUrl: string; schema: IEnrollmentListSchema } | null = null;
    private static _enrollmentListSchemaPromise: Promise<IEnrollmentListSchema> | null = null;
    private static _enrollmentEditRoleDefinitionId: number | null | undefined = undefined;
    private static _enrollmentEditRoleDefinitionPromise: Promise<number | null> | null = null;
    private static _enrollmentLearnerEditAccessSyncKeys = new Set<string>();
    private static _contentLibraryListSchemaCache: { siteUrl: string; schema: IContentLibraryListSchema } | null = null;
    private static _contentLibraryListSchemaPromise: Promise<IContentLibraryListSchema> | null = null;
    private static _auditLogListSchemaCache: { siteUrl: string; schema: IAuditLogListSchema } | null = null;
    private static _auditLogListSchemaPromise: Promise<IAuditLogListSchema> | null = null;
    private static readonly _masterCertificationCatalog: Array<{ title: string; code: string; maxSeats: number; provider: string }> = [
        { title: 'Azure Data Scientist Associate', code: 'DP-100', maxSeats: 0, provider: 'Microsoft' },
        { title: 'Azure AI Engineer Associate', code: 'AI-102', maxSeats: 0, provider: 'Microsoft' },
        { title: 'Azure Data Engineer Associate', code: 'DP-203', maxSeats: 0, provider: 'Microsoft' },
        { title: 'Fabric Analytics Engineer Associate', code: 'DP-600', maxSeats: 0, provider: 'Microsoft' },
        { title: 'DevOps Engineer Expert', code: 'AZ-400', maxSeats: 0, provider: 'Microsoft' },
        { title: 'Azure Developer Associate', code: 'AZ-204', maxSeats: 0, provider: 'Microsoft' }
    ];
    private static readonly _documents1SiteServerRelativePath = '/sites/LearningandskillDevelopment';
    private static readonly _documents1LibraryServerRelativePath = '/sites/LearningandskillDevelopment/Documents1';
    private static _extractUserEmail(user: any): string {
        const rawEmail = user?.Email || user?.email || user?.UserPrincipalName || user?.userPrincipalName;
        if (rawEmail) {
            return rawEmail.toString().trim();
        }

        const loginName = user?.LoginName || user?.loginName || '';
        if (loginName.indexOf('|') !== -1) {
            return loginName.split('|').pop()?.trim() || '';
        }

        return '';
    }

    // @ts-ignore - Method kept for future use in alternative user fetching strategies
    private static _normalizeDirectoryUser(user: any, role?: string, group?: string): any | null {
        const email = this._extractUserEmail(user);
        const title = user?.Title || user?.title || user?.DisplayName || user?.displayName || email;
        const loginName = (user?.LoginName || user?.loginName || '').toString();
        const loginNameLower = loginName.toLowerCase();
        const titleLower = (title || '').toString().toLowerCase();
        const emailLower = email.toLowerCase();

        const isSystemAccount = !email ||
            titleLower === 'system account' ||
            titleLower.indexOf('spsearch') !== -1 ||
            loginNameLower.indexOf('app@sharepoint') !== -1 ||
            loginNameLower.indexOf('nt service') !== -1 ||
            loginNameLower.indexOf('nt authority') !== -1 ||
            emailLower.indexOf('app@sharepoint') !== -1;

        if (isSystemAccount) {
            return null;
        }

        const userId = user?.Id || user?.id || email;

        return {
            Id: userId,
            Title: title,
            Email: email,
            LoginName: loginName,
            id: userId,
            userId,
            name: title,
            email,
            login: loginName,
            employeeId: loginName,
            role: role || (user?.IsSiteAdmin ? 'Owner' : 'Member'),
            group: group || (user?.IsSiteAdmin ? 'Site Owners' : 'Site Members')
        };
    }

    public static init(siteUrl: string, spHttpClient: SPHttpClient, context?: any): void {
        const contextSiteUrl = context?.pageContext?.web?.absoluteUrl;
        const resolvedSiteUrl = contextSiteUrl || siteUrl;

        if (!resolvedSiteUrl) {
            throw new Error('SharePointService.init requires a valid siteUrl (context.pageContext.web.absoluteUrl).');
        }

        this._siteUrl = resolvedSiteUrl.replace(/\/$/, '');
        this._spHttpClient = spHttpClient;
        this._context = context;
        this._documentLibraryPromise = null;
        this._membershipSnapshotCache = null;
        this._membershipSnapshotPromise = null;
        this._enrollmentListNameCache = null;
        this._userNotificationListNameCache = null;
        this._assessmentAssignmentListNameCache = null;
        this._contentLibraryListNameCache = null;
        this._auditLogListNameCache = null;
        this._certificationCompletionListNameCache = null;
        this._listTitlesCache = null;
        this._listTitlesPromise = null;
        this._formDigestCache = null;
        this._formDigestPromise = null;
        this._certificationListSchemaCache = null;
        this._certificationListSchemaPromise = null;
        this._certificationCatalogSchemaCache = null;
        this._certificationCatalogSchemaPromise = null;
        this._certificationCompletionListSchemaCache = null;
        this._certificationCompletionListSchemaPromise = null;
        this._assessmentAssignmentListSchemaCache = null;
        this._assessmentAssignmentListSchemaPromise = null;
        this._enrollmentListSchemaCache = null;
        this._enrollmentListSchemaPromise = null;
        this._contentLibraryListSchemaCache = null;
        this._contentLibraryListSchemaPromise = null;
        this._auditLogListSchemaCache = null;
        this._auditLogListSchemaPromise = null;
        this._ensuredLists.clear();
        this._pendingListEnsures.clear();
        this._listFailureTimestamps.clear();
    }

    private static _getSiteUrl(): string {
        const contextSiteUrl = this._context?.pageContext?.web?.absoluteUrl;
        const resolvedSiteUrl = contextSiteUrl || this._siteUrl;

        if (!resolvedSiteUrl) {
            throw new Error('SharePointService not initialized. Call init() first with context.pageContext.web.absoluteUrl.');
        }

        return resolvedSiteUrl.replace(/\/$/, '');
    }

    private static _getHttpClient(): SPHttpClient {
        const contextClient = this._context?.spHttpClient;
        const resolvedClient = contextClient || this._spHttpClient;

        if (!resolvedClient) {
            throw new Error('SharePointService not initialized. Call init() with context.spHttpClient first.');
        }

        return resolvedClient;
    }

    private static async _readErrorBody(response: any): Promise<string> {
        try {
            return await response.text();
        } catch (error) {
            return '';
        }
    }

    private static _getJsonHeaders(additionalHeaders: Record<string, string> = {}): Record<string, string> {
        return {
            ...this._jsonHeaders,
            ...additionalHeaders
        };
    }

    private static _getApiUrl(path: string): string {
        const siteUrl = this._getSiteUrl();
        return `${siteUrl}${path.startsWith('/') ? path : `/${path}`}`;
    }

    private static async _delay(ms: number): Promise<void> {
        if (ms <= 0) {
            return;
        }

        await new Promise((resolve) => setTimeout(resolve, ms));
    }

    private static _validateApiEndpoint(endpoint: string): void {
        const normalizedEndpoint = (endpoint || '').toString().trim().toLowerCase();
        if (!normalizedEndpoint) {
            throw new Error('SharePoint API endpoint is required.');
        }

        if (normalizedEndpoint.indexOf('/_layouts/15/throttle.htm') !== -1 || normalizedEndpoint.indexOf('/people.aspx') !== -1) {
            throw new Error(`Blocked invalid SharePoint endpoint: ${endpoint}`);
        }

        if (normalizedEndpoint.indexOf('/_api/') === -1) {
            throw new Error(`Blocked non-REST SharePoint endpoint: ${endpoint}`);
        }
    }

    private static _getRetryDelayMs(response: any, fallbackMs: number = this._defaultThrottleRetryMs): number {
        const retryAfterValue = response?.headers?.get?.('Retry-After');
        if (!retryAfterValue) {
            return fallbackMs;
        }

        const retryAfterSeconds = Number(retryAfterValue);
        if (!Number.isNaN(retryAfterSeconds) && retryAfterSeconds > 0) {
            return retryAfterSeconds * 1000;
        }

        const retryAfterDate = Date.parse(retryAfterValue);
        if (!Number.isNaN(retryAfterDate)) {
            return Math.max(retryAfterDate - Date.now(), fallbackMs);
        }

        return fallbackMs;
    }

    private static _looksLikeHtmlResponse(responseText: string, contentType: string): boolean {
        const trimmedResponse = (responseText || '').trim().toLowerCase();
        const normalizedContentType = (contentType || '').toLowerCase();

        if (!trimmedResponse) {
            return false;
        }

        return trimmedResponse.indexOf('<!doctype html') === 0 ||
            trimmedResponse.indexOf('<html') === 0 ||
            trimmedResponse.indexOf('<head') === 0 ||
            trimmedResponse.indexOf('<body') === 0 ||
            trimmedResponse.indexOf('throttle.htm') !== -1 ||
            (normalizedContentType.indexOf('json') === -1 && trimmedResponse.indexOf('<') === 0);
    }

    private static _logResponseIssueOnce(key: string, message: string, details?: any): void {
        if (this._loggedResponseIssues.has(key)) {
            return;
        }

        this._loggedResponseIssues.add(key);
        console.error(message, details);
    }

    private static async _executeJsonRequest<T>(
        requestFactory: () => Promise<any>,
        endpoint: string,
        requestLabel: string,
        retryCount: number = 1
    ): Promise<T | null> {
        this._validateApiEndpoint(endpoint);

        for (let attempt = 0; attempt <= retryCount; attempt += 1) {
            const response = await requestFactory();
            const responseText = await response.text().catch(() => '');
            const contentType = response.headers?.get?.('content-type') || '';
            const retryDelayMs = this._getRetryDelayMs(response);
            const looksLikeHtml = this._looksLikeHtmlResponse(responseText, contentType);

            if (response.status === 429 || response.status === 503) {
                if (attempt < retryCount) {
                    this._logResponseIssueOnce(
                        `${requestLabel}:status:${response.status}`,
                        `[SharePoint] ${requestLabel} was throttled. Retrying once after ${retryDelayMs} ms.`,
                        {
                            endpoint,
                            status: response.status,
                            statusText: response.statusText
                        }
                    );
                    await this._delay(retryDelayMs);
                    continue;
                }

                throw new Error(`SharePoint temporarily throttled ${requestLabel}. Wait a few seconds and try again.`);
            }

            if (!response.ok) {
                if (looksLikeHtml && attempt < retryCount) {
                    this._logResponseIssueOnce(
                        `${requestLabel}:html:${attempt}`,
                        `[SharePoint] ${requestLabel} returned HTML instead of JSON. Retrying once after ${retryDelayMs} ms.`,
                        {
                            endpoint,
                            contentType,
                            responsePreview: responseText.substring(0, 500)
                        }
                    );
                    await this._delay(retryDelayMs);
                    continue;
                }

                throw new Error(
                    `Failed to load ${requestLabel} (HTTP ${response.status} ${response.statusText}): ${responseText.substring(0, 200) || 'No error details returned.'}`
                );
            }

            const trimmedResponse = responseText.trim();
            if (!trimmedResponse) {
                return null;
            }

            if (looksLikeHtml || contentType.toLowerCase().indexOf('json') === -1) {
                if (attempt < retryCount) {
                    const retryMs = Math.max(retryDelayMs, this._defaultThrottleRetryMs);
                    this._logResponseIssueOnce(
                        `${requestLabel}:nonjson:${attempt}`,
                        `[SharePoint] ${requestLabel} returned non-JSON content. Retrying once after ${retryMs} ms.`,
                        {
                            endpoint,
                            contentType,
                            responsePreview: trimmedResponse.substring(0, 500)
                        }
                    );
                    await this._delay(retryMs);
                    continue;
                }

                throw new Error(`SharePoint returned HTML instead of JSON for ${requestLabel}. Please try again in a few seconds.`);
            }

            try {
                return JSON.parse(responseText) as T;
            } catch (error) {
                throw new Error(`SharePoint returned invalid JSON for ${requestLabel}.`);
            }
        }

        return null;
    }

    private static async _safeGetJson<T>(endpoint: string, requestLabel: string): Promise<T | null> {
        return this._executeJsonRequest<T>(
            () => this._getHttpClient().get(
                endpoint,
                SPHttpClient.configurations.v1,
                {
                    headers: this._getJsonHeaders({
                        'Cache-Control': 'no-cache, no-store, must-revalidate',
                        'Pragma': 'no-cache',
                        'Expires': '0'
                    })
                }
            ),
            endpoint,
            requestLabel
        );
    }

    private static async _safePostJson<T>(
        endpoint: string,
        options: { headers: Record<string, string>; body?: string | Blob | ArrayBuffer | ArrayBufferView },
        requestLabel: string
    ): Promise<T | null> {
        return this._executeJsonRequest<T>(
            () => this._getHttpClient().post(
                endpoint,
                SPHttpClient.configurations.v1,
                options
            ),
            endpoint,
            requestLabel
        );
    }

    private static _readMembershipSnapshotCache(siteUrl: string): ISiteMembershipSnapshot | null {
        if (this._membershipSnapshotCache?.siteUrl === siteUrl) {
            return this._membershipSnapshotCache.snapshot;
        }

        return null;
    }

    private static _writeMembershipSnapshotCache(siteUrl: string, snapshot: ISiteMembershipSnapshot): void {
        this._membershipSnapshotCache = { siteUrl, snapshot };
    }

    private static _normalizeListTitle(value: string): string {
        return (value || '').toString().trim().toLowerCase();
    }

    private static _normalizeFieldKey(value: string): string {
        const decodedValue = (value || '').toString().replace(/_x([0-9a-fA-F]{4})_/g, (_match, hex) => {
            const codePoint = parseInt(hex, 16);
            return Number.isNaN(codePoint) ? '' : String.fromCharCode(codePoint);
        });

        return decodedValue.trim().toLowerCase().replace(/[\s_]/g, '');
    }

    private static _escapeXmlAttribute(value: string): string {
        return (value || '')
            .toString()
            .replace(/&/g, '&amp;')
            .replace(/"/g, '&quot;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/'/g, '&apos;');
    }

    private static _readFieldValue(item: any, fieldName?: string | null): any {
        if (!item || !fieldName) {
            return undefined;
        }

        return item[fieldName];
    }

    private static _pickListFieldName(
        fields: ISharePointListFieldInfo[],
        candidates: string[],
        allowTitleField: boolean = false
    ): string | null {
        const normalizedCandidates = candidates.map((candidate) => this._normalizeFieldKey(candidate));
        for (const field of fields) {
            if (field?.Hidden || field?.ReadOnlyField) {
                continue;
            }

            const namesToCheck = [field.InternalName, field.StaticName, field.Title]
                .map((value) => this._normalizeFieldKey(value || ''))
                .filter((value) => !!value);

            if (!allowTitleField && namesToCheck.indexOf('title') !== -1) {
                continue;
            }

            if (namesToCheck.some((value) => normalizedCandidates.indexOf(value) !== -1)) {
                return field.InternalName || field.StaticName || field.Title || null;
            }
        }

        return null;
    }

    private static _pickListFieldNameByType(
        fields: ISharePointListFieldInfo[],
        candidates: string[],
        preferredFieldTypes: string[],
        allowTitleField: boolean = false
    ): string | null {
        const normalizedCandidates = candidates.map((candidate) => this._normalizeFieldKey(candidate));
        const matchedFields = fields.filter((field) => {
            if (field?.Hidden || field?.ReadOnlyField) {
                return false;
            }

            const namesToCheck = [field.InternalName, field.StaticName, field.Title]
                .map((value) => this._normalizeFieldKey(value || ''))
                .filter((value) => !!value);

            if (!allowTitleField && namesToCheck.indexOf('title') !== -1) {
                return false;
            }

            return namesToCheck.some((value) => normalizedCandidates.indexOf(value) !== -1);
        });

        if (matchedFields.length === 0) {
            return null;
        }

        const normalizedPreferredTypes = preferredFieldTypes.map((type) => (type || '').toString().trim().toLowerCase());
        const preferredField = matchedFields.find((field) =>
            normalizedPreferredTypes.indexOf((field.TypeAsString || '').toString().trim().toLowerCase()) !== -1
        );

        const resolvedField = preferredField || matchedFields[0];
        return resolvedField.InternalName || resolvedField.StaticName || resolvedField.Title || null;
    }

    private static _getFieldTypeMap(fields: ISharePointListFieldInfo[]): Record<string, string> {
        return fields.reduce((acc: Record<string, string>, field) => {
            const fieldName = field.InternalName || field.StaticName || field.Title || '';
            if (!fieldName) {
                return acc;
            }

            acc[fieldName] = (field.TypeAsString || '').toString().trim();
            return acc;
        }, {});
    }

    private static _isNumericFieldType(fieldType?: string): boolean {
        const normalizedType = (fieldType || '').toString().trim().toLowerCase();
        return normalizedType === 'number' ||
            normalizedType === 'currency' ||
            normalizedType === 'integer';
    }

    private static _isPersonOrLookupFieldType(fieldType?: string): boolean {
        const normalizedType = (fieldType || '').toString().trim().toLowerCase();
        return normalizedType === 'user' ||
            normalizedType === 'usermulti' ||
            normalizedType === 'lookup' ||
            normalizedType === 'lookupmulti';
    }

    private static _isTextFieldType(fieldType?: string): boolean {
        const normalizedType = (fieldType || '').toString().trim().toLowerCase();
        return !normalizedType ||
            normalizedType === 'text' ||
            normalizedType === 'note' ||
            normalizedType === 'choice';
    }

    private static _isUrlFieldType(fieldType?: string): boolean {
        const normalizedType = (fieldType || '').toString().trim().toLowerCase();
        return normalizedType === 'url';
    }

    private static _isLinkCompatibleFieldType(fieldType?: string): boolean {
        return this._isTextFieldType(fieldType) || this._isUrlFieldType(fieldType);
    }

    private static _normalizeCertificationCode(value: any): string {
        return (value || '').toString().trim().toLowerCase();
    }

    private static _normalizeCertificationLink(value: any): string {
        const trimmedValue = (value || '').toString().trim();
        if (!trimmedValue) {
            return '';
        }

        if (/^[a-z][a-z0-9+.-]*:\/\//i.test(trimmedValue)) {
            return trimmedValue;
        }

        if (trimmedValue.indexOf('//') === 0) {
            return `https:${trimmedValue}`;
        }

        return `https://${trimmedValue.replace(/^\/+/, '')}`;
    }

    private static _extractUrlFieldValue(value: any): string {
        if (!value) {
            return '';
        }

        if (typeof value === 'string') {
            return value.trim();
        }

        if (typeof value === 'object') {
            return (value.Url || value.url || '').toString().trim();
        }

        return '';
    }

    private static _buildUrlFieldPayload(url: string, description?: string): { Url: string; Description: string } | null {
        const normalizedUrl = this._normalizeCertificationLink(url);
        if (!normalizedUrl) {
            return null;
        }

        return {
            Url: normalizedUrl,
            Description: (description || '').toString().trim()
        };
    }

    private static _normalizeCertificationProvider(value: any): string {
        const normalizedValue = (value || '').toString().trim();
        const loweredValue = normalizedValue.toLowerCase();

        if (!loweredValue) {
            return '';
        }

        if (loweredValue.indexOf('google') !== -1 || loweredValue.indexOf('gcp') !== -1) {
            return 'google';
        }

        if (loweredValue.indexOf('aws') !== -1 || loweredValue.indexOf('amazon') !== -1) {
            return 'aws';
        }

        if (
            loweredValue.indexOf('microsoft') !== -1 ||
            loweredValue.indexOf('azure') !== -1 ||
            loweredValue.indexOf('teams') !== -1 ||
            loweredValue.indexOf('m365') !== -1 ||
            loweredValue.indexOf('office 365') !== -1 ||
            loweredValue.indexOf('dynamics') !== -1 ||
            loweredValue.indexOf('power platform') !== -1 ||
            loweredValue.indexOf('power bi') !== -1 ||
            loweredValue.indexOf('fabric') !== -1 ||
            loweredValue.indexOf('sharepoint') !== -1 ||
            loweredValue.indexOf('entra') !== -1
        ) {
            return 'microsoft';
        }

        if (loweredValue.indexOf('other') !== -1) {
            return 'other';
        }

        return loweredValue;
    }

    private static async _getCertificationCatalogSchema(): Promise<ICertificationCatalogSchema> {
        const siteUrl = this._ensureProductionSiteUrl();

        if (this._certificationCatalogSchemaCache?.siteUrl === siteUrl) {
            return this._certificationCatalogSchemaCache.schema;
        }

        if (this._certificationCatalogSchemaPromise) {
            return this._certificationCatalogSchemaPromise;
        }

        const listName = this._certificationAssignmentListName;
        const escapedListName = this._escapeODataValue(listName);

        this._certificationCatalogSchemaPromise = (async () => {
            const fieldsData = await this._safeGetJson<any>(
                `${siteUrl}/_api/web/lists/getbytitle('${escapedListName}')/fields?$select=Title,InternalName,StaticName,Hidden,ReadOnlyField,TypeAsString,FieldTypeKind`,
                `${listName} catalog fields`
            );
            const fields = this._toCollection(fieldsData) as ISharePointListFieldInfo[];
            const fieldTypes = this._getFieldTypeMap(fields);
            const codeField = this._pickListFieldNameByType(
                fields,
                ['Code', 'CertCode', 'Cert Code', 'CertificationCode', 'Certification Code'],
                ['Text', 'Note', 'Choice']
            );
            const maxSeatsField = this._pickListFieldNameByType(
                fields,
                ['MaxSeats', 'Max Seats'],
                ['Number', 'Currency', 'Integer']
            );
            const categoryField = this._pickListFieldNameByType(
                fields,
                ['Category'],
                ['Text', 'Note', 'Choice']
            );
            const providerField = this._pickListFieldNameByType(
                fields,
                ['Provider'],
                ['Text', 'Note', 'Choice']
            );
            const levelField = this._pickListFieldNameByType(
                fields,
                ['Level'],
                ['Text', 'Note', 'Choice']
            );
            const linkField = this._pickListFieldNameByType(
                fields,
                ['Link', 'CertificationLink', 'Certification Link', 'OfficialLink', 'Official Link', 'CertificationUrl', 'Certification URL'],
                ['URL', 'Text', 'Note', 'Choice']
            );
            const fileUrlField = this._pickListFieldNameByType(
                fields,
                ['FileUrl', 'File Url', 'URL', 'Url'],
                ['Text', 'Note', 'Choice']
            );
            const folderNameField = this._pickListFieldNameByType(
                fields,
                ['FolderName', 'Folder Name', 'Folder'],
                ['Text', 'Note', 'Choice']
            );
            const schema: ICertificationCatalogSchema = {
                codeField,
                maxSeatsField,
                categoryField,
                providerField,
                levelField,
                linkField,
                fileUrlField,
                folderNameField,
                codeFieldType: codeField ? (fieldTypes[codeField] || '') : '',
                maxSeatsFieldType: maxSeatsField ? (fieldTypes[maxSeatsField] || '') : '',
                categoryFieldType: categoryField ? (fieldTypes[categoryField] || '') : '',
                providerFieldType: providerField ? (fieldTypes[providerField] || '') : '',
                levelFieldType: levelField ? (fieldTypes[levelField] || '') : '',
                linkFieldType: linkField ? (fieldTypes[linkField] || '') : '',
                fileUrlFieldType: fileUrlField ? (fieldTypes[fileUrlField] || '') : '',
                folderNameFieldType: folderNameField ? (fieldTypes[folderNameField] || '') : ''
            };

            console.log('[Certifications] Resolved catalog field schema', schema);

            this._certificationCatalogSchemaCache = { siteUrl, schema };
            this._certificationCatalogSchemaPromise = null;
            return schema;
        })().catch((error) => {
            this._certificationCatalogSchemaPromise = null;
            throw error;
        });

        return this._certificationCatalogSchemaPromise;
    }

    private static async _getCertificationListSchema(): Promise<ICertificationListSchema> {
        const siteUrl = this._ensureProductionSiteUrl();

        if (this._certificationListSchemaCache?.siteUrl === siteUrl) {
            return this._certificationListSchemaCache.schema;
        }

        if (this._certificationListSchemaPromise) {
            return this._certificationListSchemaPromise;
        }

        const listName = this._certificationAssignmentListName;
        const escapedListName = this._escapeODataValue(listName);

        this._certificationListSchemaPromise = (async () => {
            const fieldsData = await this._safeGetJson<any>(
                `${siteUrl}/_api/web/lists/getbytitle('${escapedListName}')/fields?$select=Title,InternalName,StaticName,Hidden,ReadOnlyField`,
                `${listName} fields`
            );
            const fields = this._toCollection(fieldsData) as ISharePointListFieldInfo[];
            const assignedToField = this._pickListFieldName(fields, ['AssignedTo', 'Assigned To']);
            const scheduledDateField = this._pickListFieldName(fields, ['ScheduledDate', 'Scheduled Date', 'AssignedDate', 'Assigned Date', 'IssuedDate']);

            if (!assignedToField) {
                throw new Error("The 'Certifications' list must contain an 'AssignedTo' person field.");
            }

            if (!scheduledDateField) {
                throw new Error("The 'Certifications' list must contain a 'ScheduledDate' date field.");
            }

            const schema: ICertificationListSchema = {
                assignedToField,
                userEmailField: this._pickListFieldName(fields, ['UserEmail', 'AssignedToEmail', 'Email']),
                userNameField: this._pickListFieldName(fields, ['UserName', 'AssignedToName', 'Name']),
                certificationNameField: this._pickListFieldName(fields, ['CertificationName', 'Certification Name', 'CertName', 'CertificateName', 'Certification']),
                scheduledDateField,
                assignedDateField: this._pickListFieldName(fields, ['AssignedDate', 'Assigned Date', 'IssuedDate']),
                issuedDateField: this._pickListFieldName(fields, ['IssuedDate', 'AssignedDate']),
                expiryDateField: this._pickListFieldName(fields, ['ExpiryDate', 'Expiry Date', 'ExpirationDate', 'Expiration Date', 'ValidTill', 'Valid Till', 'EndDate', 'End Date']),
                statusField: this._pickListFieldName(fields, ['Status']),
                certCodeField: this._pickListFieldName(fields, ['Code', 'CertCode', 'Cert Code', 'CertificationCode', 'Certification Code']),
                orderIndexField: this._pickListFieldName(fields, ['OrderIndex']),
                assignedGroupField: this._pickListFieldName(fields, ['AssignedGroup', 'Group'])
            };

            console.log('[Certifications] Resolved SharePoint field schema', {
                listName,
                assignedToField: schema.assignedToField,
                scheduledDateField: schema.scheduledDateField,
                certificationNameField: schema.certificationNameField
            });

            this._certificationListSchemaCache = { siteUrl, schema };
            this._certificationListSchemaPromise = null;
            return schema;
        })().catch((error) => {
            this._certificationListSchemaPromise = null;
            throw error;
        });

        return this._certificationListSchemaPromise;
    }

    private static async _getCertificationCompletionListSchema(): Promise<ICertificationCompletionListSchema> {
        const siteUrl = this._ensureProductionSiteUrl();

        if (this._certificationCompletionListSchemaCache?.siteUrl === siteUrl) {
            return this._certificationCompletionListSchemaCache.schema;
        }

        if (this._certificationCompletionListSchemaPromise) {
            return this._certificationCompletionListSchemaPromise;
        }

        this._certificationCompletionListSchemaPromise = (async () => {
            const listName = await this._ensureCertificationCompletionList();
            const escapedListName = this._escapeODataValue(listName);
            const fieldsData = await this._safeGetJson<any>(
                `${siteUrl}/_api/web/lists/getbytitle('${escapedListName}')/fields?$select=Title,InternalName,StaticName,Hidden,ReadOnlyField`,
                `${listName} completion fields`
            );
            const fields = this._toCollection(fieldsData) as ISharePointListFieldInfo[];
            const schema: ICertificationCompletionListSchema = {
                certIdField: this._pickListFieldName(fields, ['CertID', 'Cert Id', 'CertificationId', 'Certification ID']),
                examDateField: this._pickListFieldName(fields, ['ExamDate', 'Exam Date']),
                renewalDateField: this._pickListFieldName(fields, ['RenewalDate', 'Renewal Date']),
                examCodeField: this._pickListFieldName(fields, ['ExamCode', 'Exam Code'])
            };

            console.log('[uploadlist] Resolved completion field schema', {
                listName,
                certIdField: schema.certIdField,
                examDateField: schema.examDateField,
                renewalDateField: schema.renewalDateField,
                examCodeField: schema.examCodeField
            });

            this._certificationCompletionListSchemaCache = { siteUrl, schema };
            this._certificationCompletionListSchemaPromise = null;
            return schema;
        })().catch((error) => {
            this._certificationCompletionListSchemaPromise = null;
            throw error;
        });

        return this._certificationCompletionListSchemaPromise;
    }

    private static async _getContentLibraryListSchema(): Promise<IContentLibraryListSchema> {
        const siteUrl = this._getSiteUrl();

        if (this._contentLibraryListSchemaCache?.siteUrl === siteUrl) {
            return this._contentLibraryListSchemaCache.schema;
        }

        if (this._contentLibraryListSchemaPromise) {
            return this._contentLibraryListSchemaPromise;
        }

        this._contentLibraryListSchemaPromise = (async () => {
            const listName = await this._ensureContentLibraryList();
            const escapedListName = this._escapeODataValue(listName);
            const fieldsData = await this._safeGetJson<any>(
                `${siteUrl}/_api/web/lists/getbytitle('${escapedListName}')/fields?$select=Title,InternalName,StaticName,Hidden,ReadOnlyField,TypeAsString,FieldTypeKind`,
                `${listName} content library fields`
            );
            const fields = this._toCollection(fieldsData) as ISharePointListFieldInfo[];
            const schema: IContentLibraryListSchema = {
                fileLinkField: this._pickListFieldNameByType(fields, ['FileLink', 'File Link', 'FileUrl', 'File Url', 'Link'], ['Text', 'Note', 'Choice']) || 'FileLink',
                uploadedByField: this._pickListFieldNameByType(fields, ['UploadedBy', 'Uploaded By'], ['Text', 'Note', 'Choice']) || 'UploadedBy',
                assignedToField: this._pickListFieldNameByType(fields, ['AssignedTo', 'Assigned To'], ['Text', 'Note', 'Choice']) || 'AssignedTo',
                statusField: this._pickListFieldNameByType(fields, ['Status'], ['Text', 'Note', 'Choice']) || 'Status',
                folderNameField: this._pickListFieldNameByType(fields, ['FolderName', 'Folder Name', 'Folder'], ['Text', 'Note', 'Choice']),
                assetTypeField: this._pickListFieldNameByType(fields, ['AssetType', 'Asset Type', 'Type'], ['Text', 'Note', 'Choice']),
                descriptionField: this._pickListFieldNameByType(fields, ['Description'], ['Text', 'Note', 'Choice']),
                fileSizeField: this._pickListFieldNameByType(fields, ['FileSize', 'File Size', 'Size'], ['Text', 'Note', 'Choice'])
            };

            this._contentLibraryListSchemaCache = { siteUrl, schema };
            this._contentLibraryListSchemaPromise = null;
            return schema;
        })().catch((error) => {
            this._contentLibraryListSchemaPromise = null;
            throw error;
        });

        return this._contentLibraryListSchemaPromise;
    }

    private static async _getAssessmentAssignmentListSchema(): Promise<IAssessmentAssignmentListSchema> {
        const siteUrl = this._ensureProductionSiteUrl();

        if (this._assessmentAssignmentListSchemaCache?.siteUrl === siteUrl) {
            return this._assessmentAssignmentListSchemaCache.schema;
        }

        if (this._assessmentAssignmentListSchemaPromise) {
            return this._assessmentAssignmentListSchemaPromise;
        }

        const listName = await this._resolveAssessmentAssignmentListName();
        const escapedListName = this._escapeODataValue(listName);

        this._assessmentAssignmentListSchemaPromise = (async () => {
            const fieldsData = await this._safeGetJson<any>(
                `${siteUrl}/_api/web/lists/getbytitle('${escapedListName}')/fields?$select=Title,InternalName,StaticName,Hidden,ReadOnlyField`,
                `${listName} fields`
            );
            const fields = this._toCollection(fieldsData) as ISharePointListFieldInfo[];
            const assignedToField = this._pickListFieldName(fields, ['AssignedTo', 'Assigned To']);
            const scheduledDateField = this._pickListFieldName(fields, ['ScheduledDate', 'Scheduled Date']);

            if (!assignedToField) {
                throw new Error(`The '${listName}' list must contain an 'AssignedTo' person field.`);
            }

            if (!scheduledDateField) {
                throw new Error(`The '${listName}' list must contain a 'ScheduledDate' date field.`);
            }

            const schema: IAssessmentAssignmentListSchema = {
                assignedToField,
                scheduledDateField,
                userEmailField: this._pickListFieldName(fields, ['UserEmail', 'AssignedToEmail', 'Email']),
                userNameField: this._pickListFieldName(fields, ['UserName', 'AssignedToName', 'Name']),
                assessmentNameField: this._pickListFieldName(fields, ['AssessmentName', 'Assessment Name']),
                orderIndexField: this._pickListFieldName(fields, ['OrderIndex']),
                assignedGroupField: this._pickListFieldName(fields, ['AssignedGroup', 'Group']),
                assessmentPayloadField: this._pickListFieldName(fields, ['AssessmentPayload', 'Assessment Payload'])
            };

            console.log('[AssessmentAssignments] Resolved SharePoint field schema', {
                listName,
                assignedToField: schema.assignedToField,
                scheduledDateField: schema.scheduledDateField,
                assessmentNameField: schema.assessmentNameField
            });

            this._assessmentAssignmentListSchemaCache = { siteUrl, schema };
            this._assessmentAssignmentListSchemaPromise = null;
            return schema;
        })().catch((error) => {
            this._assessmentAssignmentListSchemaPromise = null;
            throw error;
        });

        return this._assessmentAssignmentListSchemaPromise;
    }

    public static async fetchCertificationMaxSeats(forceRefresh: boolean = false): Promise<ICertificationMaxSeatsItem[]> {
        const siteUrl = this._ensureProductionSiteUrl();
        const listName = this._certificationAssignmentListName;
        const schema = await this._getCertificationCatalogSchema();
        const selectFields = Array.from(new Set([
            'Id',
            'Title',
            schema.codeField || '',
            schema.maxSeatsField || '',
            schema.categoryField || '',
            schema.providerField || '',
            schema.levelField || '',
            schema.linkField || '',
            schema.fileUrlField || ''
        ].filter((field) => !!field)));
        const endpoint =
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items` +
            `?$select=${selectFields.join(',')}` +
            `&$top=5000` +
            `&_=${Date.now()}`;

        const data = await this._safeGetJson<any>(endpoint, `${listName} MaxSeats`);
        const items = this._toCollection(data).map((item: any) => {
            const fileUrl = ((schema.fileUrlField ? this._readFieldValue(item, schema.fileUrlField) : undefined) || '').toString().trim();
            const rawCategory = ((schema.categoryField ? this._readFieldValue(item, schema.categoryField) : undefined) || item?.Category || '').toString().trim();
            const rawLevel = ((schema.levelField ? this._readFieldValue(item, schema.levelField) : undefined) || item?.Level || '').toString().trim();
            const rawCode = ((schema.codeField ? this._readFieldValue(item, schema.codeField) : undefined) || item?.Code || '').toString().trim();
            const rawProvider = ((schema.providerField ? this._readFieldValue(item, schema.providerField) : undefined) || item?.Provider || '').toString().trim();
            const certificationLink = this._normalizeCertificationLink(
                (schema.linkField ? this._extractUrlFieldValue(this._readFieldValue(item, schema.linkField)) : undefined) ||
                item?.Link ||
                item?.CertificationLink ||
                ''
            );
            const normalizedProvider = this._normalizeCertificationProvider(
                rawProvider || [item?.Title || '', rawCode, rawCategory, rawLevel].join(' ')
            );
            const storedEnrollmentCount = Number(
                (schema.maxSeatsField ? this._readFieldValue(item, schema.maxSeatsField) : undefined) ||
                item?.MaxSeats ||
                0
            );
            return {
                id: Number(item?.Id || item?.id || 0),
                title: (item?.Title || '').toString().trim(),
                code: rawCode,
                maxSeats: storedEnrollmentCount,
                enrolledCount: storedEnrollmentCount,
                assignedLearnerCount: storedEnrollmentCount,
                category: rawCategory,
                provider: normalizedProvider,
                level: rawLevel,
                link: certificationLink,
                fileUrl
            };
        })
            .filter((item: ICertificationMaxSeatsItem & { fileUrl?: string }) => !!item.title && (!item.fileUrl || !!item.code))
            .sort((a: ICertificationMaxSeatsItem, b: ICertificationMaxSeatsItem) => a.title.localeCompare(b.title));

        console.log('[Certifications] MaxSeats API response', {
            endpoint,
            count: items.length,
            items
        });
        console.log('Total Certifications:', items.length);

        return items;
    }

    public static createCertificationMaxSeatsMap(certifications: ICertificationMaxSeatsItem[]): Map<string, number> {
        const map = new Map<string, number>();

        (certifications || []).forEach((item) => {
            const titleKey = (item?.title || '').toString().trim().toLowerCase();
            const codeKey = this._normalizeCertificationCode(item?.code);
            const assignedLearnerCount = this._getStoredCertificationCount(item);

            if (!titleKey) {
                if (!codeKey) {
                    return;
                }
            }

            if (titleKey) {
                map.set(titleKey, assignedLearnerCount);
            }

            if (codeKey) {
                map.set(codeKey, assignedLearnerCount);
            }
        });

        return map;
    }

    public static async getCertificationDetailsByTitle(certName: string, forceRefresh: boolean = false): Promise<ICertificationMaxSeatsItem | null> {
        const normalizedCertName = (certName || '').toString().trim();
        if (!normalizedCertName) {
            return null;
        }

        const siteUrl = this._ensureProductionSiteUrl();
        const endpoint =
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(this._certificationAssignmentListName)}')/items` +
            `?$select=Id,Title` +
            `&$filter=Title eq '${this._escapeODataValue(normalizedCertName)}'` +
            `&$top=1`;

        try {
            const certifications = await this.fetchCertificationMaxSeats(forceRefresh);
            return certifications.find((item) => (item.title || '').toString().trim().toLowerCase() === normalizedCertName.toLowerCase()) || null;
        } catch (error) {
            console.warn('[Certifications] Direct certification lookup failed. Falling back to cached list read.', {
                certification: normalizedCertName,
                endpoint,
                error
            });

            const certifications = await this.fetchCertificationMaxSeats(forceRefresh);
            return certifications.find((item) => (item.title || '').toString().trim().toLowerCase() === normalizedCertName.toLowerCase()) || null;
        }
    }

    public static async getCertificationDetailsByCodeOrTitle(certCode: string, certName: string, forceRefresh: boolean = false): Promise<ICertificationMaxSeatsItem | null> {
        const normalizedCode = this._normalizeCertificationCode(certCode);
        const normalizedTitle = (certName || '').toString().trim().toLowerCase();
        const certifications = await this.fetchCertificationMaxSeats(forceRefresh);

        return certifications.find((item) =>
            (!!normalizedCode && this._normalizeCertificationCode(item.code) === normalizedCode) ||
            (!!normalizedTitle && (item.title || '').toString().trim().toLowerCase() === normalizedTitle)
        ) || null;
    }

    public static async getCertificationDetailsById(certificationId: number, forceRefresh: boolean = false): Promise<ICertificationMaxSeatsItem | null> {
        const normalizedCertificationId = Number(certificationId || 0);
        if (normalizedCertificationId <= 0) {
            return null;
        }

        const certifications = await this.fetchCertificationMaxSeats(forceRefresh);
        return certifications.find((item) => Number(item.id || 0) === normalizedCertificationId) || null;
    }

    public static async getCertificationMaxSeatsMap(forceRefresh: boolean = false): Promise<Map<string, number>> {
        const certifications = await this.fetchCertificationMaxSeats(true);
        return this.createCertificationMaxSeatsMap(certifications);
    }

    public static async syncDefaultCertifications(forceRefresh: boolean = false): Promise<ICertificationMaxSeatsItem[]> {
        const existingItems = await this.fetchCertificationMaxSeats(forceRefresh);
        const existingKeys = new Set<string>();

        existingItems.forEach((item) => {
            const normalizedCode = this._normalizeCertificationCode(item.code);
            const normalizedTitle = (item.title || '').toString().trim().toLowerCase();
            if (normalizedCode) {
                existingKeys.add(`code:${normalizedCode}`);
            }
            if (normalizedTitle) {
                existingKeys.add(`title:${normalizedTitle}`);
            }
        });

        const missingDefaults = this._masterCertificationCatalog.filter((item) => {
            const normalizedCode = this._normalizeCertificationCode(item.code);
            const normalizedTitle = (item.title || '').toString().trim().toLowerCase();
            if (normalizedCode && existingKeys.has(`code:${normalizedCode}`)) {
                return false;
            }
            if (!normalizedCode && normalizedTitle && existingKeys.has(`title:${normalizedTitle}`)) {
                return false;
            }
            return true;
        });

        if (missingDefaults.length === 0) {
            console.log('[Certifications] Default certification sync is up to date.', {
                existingCount: existingItems.length
            });
            return existingItems;
        }

        console.log('[Certifications] Creating missing default certifications.', {
            missingDefaults
        });

        for (const item of missingDefaults) {
            await this.createCertificationItem(item.title, item.maxSeats, item.code, { skipDuplicateCheck: true, provider: item.provider });
            await this._delay(100);
        }

        return this.fetchCertificationMaxSeats(true);
    }

    public static async seedPredefinedCertifications(): Promise<{ createdCount: number; updatedCount: number; skippedCount: number; totalProcessed: number; }> {
        return this.bulkUpsertCertificationItems(
            this._masterCertificationCatalog.map((item) => ({
                title: item.title,
                code: item.code,
                maxSeats: 0,
                provider: item.provider
            }))
        );
    }

    private static _getStoredCertificationCount(item: Partial<ICertificationMaxSeatsItem> | null | undefined): number {
        const parsedValue = Number(
            item?.assignedLearnerCount ??
            item?.enrolledCount ??
            item?.maxSeats ??
            0
        );

        return Number.isFinite(parsedValue) && parsedValue >= 0 ? parsedValue : 0;
    }

    private static _invalidateCertificationCountCache(): void {
        // MaxSeats is read fresh from SharePoint on every certification load.
    }

    public static async updateCertificationEnrolledCount(id: number, enrolledCount: number): Promise<void> {
        const parsedItemId = Number(id);
        if (!Number.isFinite(parsedItemId) || parsedItemId <= 0) {
            throw new Error('selectedCertification.Id is invalid for the Certifications update.');
        }

        const parsedEnrolledCount = Number(enrolledCount);
        const normalizedEnrolledCount = Number.isFinite(parsedEnrolledCount) && parsedEnrolledCount >= 0 ? parsedEnrolledCount : 0;

        const siteUrl = this._ensureProductionSiteUrl();
        const schema = await this._getCertificationCatalogSchema();
        if (!schema.maxSeatsField) {
            throw new Error("The 'Certifications' list must contain a 'MaxSeats' number column.");
        }

        if (!this._isNumericFieldType(schema.maxSeatsFieldType)) {
            throw new Error(`The '${schema.maxSeatsField}' column on the Certifications list must be a Number column.`);
        }

        const digest = await this._getFormDigestValue();
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(this._certificationAssignmentListName)}')/items(${parsedItemId})`;
        const payload = {
            [schema.maxSeatsField]: Number(normalizedEnrolledCount)
        };

        console.log('[Certifications] Assigned learner count MERGE request', {
            endpoint,
            payload
        });

        const response = await this._getHttpClient().post(
            endpoint,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'MERGE',
                    'X-RequestDigest': digest
                }),
                body: JSON.stringify(payload)
            }
        );

        if (!response.ok) {
            const errorText = await this._readErrorBody(response);
            console.error('[Certifications] Assigned learner count update failed', {
                id: parsedItemId,
                enrolledCount: normalizedEnrolledCount,
                endpoint,
                status: response.status,
                statusText: response.statusText,
                responseText: errorText
            });
            throw new Error(`Failed to update Certifications.MaxSeats assigned count (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 400) || 'No error details returned.'}`);
        }

        this._invalidateCertificationCountCache();
    }

    public static async updateCertificationMaxSeats(id: number, maxSeats: number): Promise<void> {
        await this.updateCertificationEnrolledCount(id, maxSeats);
    }

    public static async updateCertificationItem(
        id: number,
        title: string,
        maxSeats: number,
        code: string,
        options: { provider?: string; link?: string } = {}
    ): Promise<void> {
        const parsedItemId = Number(id);
        if (!Number.isFinite(parsedItemId) || parsedItemId <= 0) {
            throw new Error('selectedCertification.Id is invalid for the Certifications update.');
        }

        const normalizedTitle = (title || '').toString().trim();
        const normalizedCode = (code || '').toString().trim().toUpperCase();
        const parsedMaxSeats = Number(maxSeats);
        const normalizedMaxSeats = Number.isFinite(parsedMaxSeats) && parsedMaxSeats >= 0 ? parsedMaxSeats : 0;
        const normalizedProvider = this._normalizeCertificationProvider(options.provider);
        const normalizedLink = this._normalizeCertificationLink(options.link);

        if (!normalizedTitle) {
            throw new Error('Certification title is required.');
        }

        if (!normalizedCode) {
            throw new Error('Certification code is required.');
        }

        const schema = await this._getCertificationCatalogSchema();
        if (!schema.codeField) {
            throw new Error("The 'Certifications' list must contain a 'Code' text column.");
        }

        if (!this._isTextFieldType(schema.codeFieldType)) {
            throw new Error(`The '${schema.codeField}' column on the Certifications list must be a text column.`);
        }

        if (!schema.maxSeatsField) {
            throw new Error("The 'Certifications' list must contain a 'MaxSeats' number column.");
        }

        if (!this._isNumericFieldType(schema.maxSeatsFieldType)) {
            throw new Error(`The '${schema.maxSeatsField}' column on the Certifications list must be a Number column.`);
        }

        if (normalizedProvider && schema.providerField && !this._isTextFieldType(schema.providerFieldType)) {
            throw new Error(`The '${schema.providerField}' column on the Certifications list must be a text column.`);
        }

        if (normalizedLink && !schema.linkField) {
            throw new Error("The 'Certifications' list must contain a 'Link' text or hyperlink column.");
        }

        if (normalizedLink && schema.linkField && !this._isLinkCompatibleFieldType(schema.linkFieldType)) {
            throw new Error(`The '${schema.linkField}' column on the Certifications list must be a text or hyperlink column.`);
        }

        const existingItems = await this.fetchCertificationMaxSeats(true);
        const duplicate = existingItems.find((item) =>
            Number(item.id || 0) !== parsedItemId &&
            this._normalizeCertificationCode(item.code) === this._normalizeCertificationCode(normalizedCode)
        );
        if (duplicate) {
            throw new Error('Certification already exists');
        }

        const siteUrl = this._ensureProductionSiteUrl();
        const digest = await this._getFormDigestValue();
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(this._certificationAssignmentListName)}')/items(${parsedItemId})`;
        const payload: Record<string, any> = {
            Title: normalizedTitle,
            [schema.codeField]: normalizedCode,
            [schema.maxSeatsField]: Number(normalizedMaxSeats)
        };

        if (schema.providerField && this._isTextFieldType(schema.providerFieldType)) {
            payload[schema.providerField] = normalizedProvider || 'Other';
        }

        if (schema.linkField) {
            if (this._isUrlFieldType(schema.linkFieldType)) {
                payload[schema.linkField] = normalizedLink
                    ? this._buildUrlFieldPayload(normalizedLink, normalizedTitle)
                    : null;
            } else if (this._isTextFieldType(schema.linkFieldType)) {
                payload[schema.linkField] = normalizedLink || '';
            }
        }

        console.log('[Certifications] Update certification payload', {
            endpoint,
            payload
        });

        const response = await this._getHttpClient().post(
            endpoint,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'MERGE',
                    'X-RequestDigest': digest
                }),
                body: JSON.stringify(payload)
            }
        );

        if (!response.ok) {
            const errorText = await this._readErrorBody(response);
            console.error('[Certifications] Update certification failed', {
                endpoint,
                payload,
                status: response.status,
                statusText: response.statusText,
                responseText: errorText
            });
            throw new Error(`Failed to update certification (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 400) || 'No error details returned.'}`);
        }

        this._invalidateCertificationCountCache();
    }

    public static async createCertificationItem(title: string, maxSeats: number, code: string, options: { skipDuplicateCheck?: boolean; provider?: string; link?: string } = {}): Promise<number> {
        const normalizedTitle = (title || '').toString().trim();
        const normalizedCode = (code || '').toString().trim();
        const parsedMaxSeats = Number(maxSeats);
        const normalizedMaxSeats = Number.isFinite(parsedMaxSeats) && parsedMaxSeats >= 0 ? parsedMaxSeats : 0;
        const normalizedProvider = this._normalizeCertificationProvider(options.provider);
        const normalizedLink = this._normalizeCertificationLink(options.link);

        if (!normalizedTitle) {
            throw new Error('Certification title is required.');
        }

        if (!normalizedCode) {
            throw new Error('Certification code is required.');
        }

        const schema = await this._getCertificationCatalogSchema();
        if (!schema.codeField) {
            throw new Error("The 'Certifications' list must contain a 'Code' text column.");
        }

        if (!this._isTextFieldType(schema.codeFieldType)) {
            throw new Error(`The '${schema.codeField}' column on the Certifications list must be a text column.`);
        }

        if (!schema.maxSeatsField) {
            throw new Error("The 'Certifications' list must contain a 'MaxSeats' number column.");
        }

        if (!this._isNumericFieldType(schema.maxSeatsFieldType)) {
            throw new Error(`The '${schema.maxSeatsField}' column on the Certifications list must be a Number column.`);
        }

        if (normalizedProvider && schema.providerField && !this._isTextFieldType(schema.providerFieldType)) {
            throw new Error(`The '${schema.providerField}' column on the Certifications list must be a text column.`);
        }

        if (normalizedLink && !schema.linkField) {
            throw new Error("The 'Certifications' list must contain a 'Link' text or hyperlink column.");
        }

        if (normalizedLink && schema.linkField && !this._isLinkCompatibleFieldType(schema.linkFieldType)) {
            throw new Error(`The '${schema.linkField}' column on the Certifications list must be a text or hyperlink column.`);
        }

        if (!options.skipDuplicateCheck) {
            const existingItems = await this.fetchCertificationMaxSeats(false);
            const duplicate = existingItems.find((item) =>
                this._normalizeCertificationCode(item.code) === this._normalizeCertificationCode(normalizedCode)
            );

            if (duplicate) {
                throw new Error('Certification already exists');
            }
        }

        const siteUrl = this._ensureProductionSiteUrl();
        const digest = await this._getFormDigestValue();
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(this._certificationAssignmentListName)}')/items`;
        const payload: Record<string, any> = {
            Title: normalizedTitle,
            [schema.codeField]: normalizedCode,
            [schema.maxSeatsField]: Number(normalizedMaxSeats)
        };

        if (normalizedProvider && schema.providerField && this._isTextFieldType(schema.providerFieldType)) {
            payload[schema.providerField] = normalizedProvider;
        }

        if (normalizedLink && schema.linkField) {
            if (this._isUrlFieldType(schema.linkFieldType)) {
                payload[schema.linkField] = this._buildUrlFieldPayload(normalizedLink, normalizedTitle);
            } else if (this._isTextFieldType(schema.linkFieldType)) {
                payload[schema.linkField] = normalizedLink;
            }
        }

        if (Number.isNaN(Number(payload[schema.maxSeatsField]))) {
            payload[schema.maxSeatsField] = 0;
        }

        console.log('[Certifications] Create certification payload', {
            endpoint,
            payload
        });

        const response = await this._getHttpClient().post(
            endpoint,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'X-RequestDigest': digest
                }),
                body: JSON.stringify(payload)
            }
        );

        if (!response.ok) {
            const errorText = await this._readErrorBody(response);
            console.error('[Certifications] Create certification failed', {
                endpoint,
                payload,
                status: response.status,
                statusText: response.statusText,
                responseText: errorText
            });
            throw new Error(`Failed to create certification (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 400) || 'No error details returned.'}`);
        }

        this._invalidateCertificationCountCache();
        const createdItem = await this._readJson<any>(response);
        return Number(createdItem?.Id || createdItem?.id || 0);
    }

    public static async deleteCertificationItem(id: number, certName?: string, certCode?: string): Promise<void> {
        const parsedItemId = Number(id);
        if (!Number.isFinite(parsedItemId) || parsedItemId <= 0) {
            throw new Error('selectedCertification.Id is invalid for deletion.');
        }

        const normalizedCertName = (certName || '').toString().trim();
        const normalizedCertCode = (certCode || '').toString().trim().toUpperCase();
        const activeEnrollmentCount = await this.getEnrollmentCountForCertificationId(parsedItemId, normalizedCertName, normalizedCertCode);
        if (activeEnrollmentCount > 0) {
            throw new Error('Cannot delete certification with active enrollments.');
        }

        const siteUrl = this._ensureProductionSiteUrl();
        const digest = await this._getFormDigestValue();
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(this._certificationAssignmentListName)}')/items(${parsedItemId})`;
        const response = await this._getHttpClient().post(
            endpoint,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'DELETE',
                    'X-RequestDigest': digest
                })
            }
        );

        if (!response.ok) {
            const errorText = await this._readErrorBody(response);
            console.error('[Certifications] Delete certification failed', {
                endpoint,
                id: parsedItemId,
                status: response.status,
                statusText: response.statusText,
                responseText: errorText
            });
            throw new Error(`Failed to delete certification (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 400) || 'No error details returned.'}`);
        }

        this._invalidateCertificationCountCache();
    }

    public static async bulkUpsertCertificationItems(
        items: ICertificationImportRow[]
    ): Promise<{ createdCount: number; updatedCount: number; skippedCount: number; totalProcessed: number; }> {
        const schema = await this._getCertificationCatalogSchema();
        if (!schema.codeField) {
            throw new Error("The 'Certifications' list must contain a 'Code' text column.");
        }

        if (!this._isTextFieldType(schema.codeFieldType)) {
            throw new Error(`The '${schema.codeField}' column on the Certifications list must be a text column.`);
        }

        if (!schema.maxSeatsField) {
            throw new Error("The 'Certifications' list must contain a 'MaxSeats' number column.");
        }

        if (!this._isNumericFieldType(schema.maxSeatsFieldType)) {
            throw new Error(`The '${schema.maxSeatsField}' column on the Certifications list must be a Number column.`);
        }

        const normalizedEntries = new Map<string, ICertificationImportRow>();
        let skippedCount = 0;

        (items || []).forEach((item) => {
            const title = (item?.title || '').toString().trim();
            const code = (item?.code || '').toString().trim().toUpperCase();
            const parsedMaxSeats = Number(item?.maxSeats);
            const maxSeats = Number.isFinite(parsedMaxSeats) && parsedMaxSeats >= 0 ? parsedMaxSeats : 0;
            const provider = this._normalizeCertificationProvider(item?.provider);

            if (!title || !code) {
                skippedCount += 1;
                return;
            }

            const dedupeKey = this._normalizeCertificationCode(code) || title.toLowerCase();
            normalizedEntries.set(dedupeKey, {
                title,
                code,
                maxSeats,
                provider
            });
        });

        const entriesToProcess = Array.from(normalizedEntries.values());
        if (entriesToProcess.length === 0) {
            return {
                createdCount: 0,
                updatedCount: 0,
                skippedCount,
                totalProcessed: 0
            };
        }

        if (
            entriesToProcess.some((entry) => !!entry.provider) &&
            schema.providerField &&
            !this._isTextFieldType(schema.providerFieldType)
        ) {
            throw new Error(`The '${schema.providerField}' column on the Certifications list must be a text column.`);
        }

        const existingItems = await this.fetchCertificationMaxSeats(true);
        const itemsByCode = new Map<string, ICertificationMaxSeatsItem>();
        const itemsByTitle = new Map<string, ICertificationMaxSeatsItem>();

        existingItems.forEach((item) => {
            const normalizedCode = this._normalizeCertificationCode(item.code);
            const normalizedTitle = (item.title || '').toString().trim().toLowerCase();

            if (normalizedCode) {
                itemsByCode.set(normalizedCode, item);
            }

            if (normalizedTitle) {
                itemsByTitle.set(normalizedTitle, item);
            }
        });

        const siteUrl = this._ensureProductionSiteUrl();
        const listName = this._certificationAssignmentListName;
        const escapedListName = this._escapeODataValue(listName);
        const digest = await this._getFormDigestValue();
        let createdCount = 0;
        let updatedCount = 0;

        await this._processInChunks(entriesToProcess, 20, async (chunk) => {
            await Promise.all(chunk.map(async (entry) => {
                const normalizedCode = this._normalizeCertificationCode(entry.code);
                const normalizedTitle = entry.title.toLowerCase();
                const existingItem = itemsByCode.get(normalizedCode) || itemsByTitle.get(normalizedTitle) || null;

                if (existingItem) {
                    const titleChanged = (existingItem.title || '').toString().trim() !== entry.title;
                    const codeChanged = this._normalizeCertificationCode(existingItem.code) !== normalizedCode;
                    const providerChanged =
                        !!entry.provider &&
                        this._normalizeCertificationProvider(existingItem.provider) !== this._normalizeCertificationProvider(entry.provider);

                    if (!titleChanged && !codeChanged && !providerChanged) {
                        skippedCount += 1;
                        return;
                    }

                    const updateEndpoint = `${siteUrl}/_api/web/lists/getbytitle('${escapedListName}')/items(${Number(existingItem.id)})`;
                    const updatePayload: Record<string, any> = {
                        Title: entry.title,
                        [schema.codeField as string]: entry.code
                    };

                    if (entry.provider && schema.providerField && this._isTextFieldType(schema.providerFieldType)) {
                        updatePayload[schema.providerField] = entry.provider;
                    }

                    const updateResponse = await this._getHttpClient().post(
                        updateEndpoint,
                        SPHttpClient.configurations.v1,
                        {
                            headers: this._getJsonHeaders({
                                'IF-MATCH': '*',
                                'X-HTTP-Method': 'MERGE',
                                'X-RequestDigest': digest
                            }),
                            body: JSON.stringify(updatePayload)
                        }
                    );

                    if (!updateResponse.ok) {
                        const errorText = await this._readErrorBody(updateResponse);
                        throw new Error(`Failed to update certification '${entry.title}' (HTTP ${updateResponse.status} ${updateResponse.statusText}): ${errorText.substring(0, 300) || 'No error details returned.'}`);
                    }

                    updatedCount += 1;
                    const updatedItem: ICertificationMaxSeatsItem = {
                        ...existingItem,
                        title: entry.title,
                        code: entry.code,
                        provider: entry.provider || existingItem.provider
                    };
                    itemsByCode.set(normalizedCode, updatedItem);
                    itemsByTitle.set(normalizedTitle, updatedItem);
                    return;
                }

                const createEndpoint = `${siteUrl}/_api/web/lists/getbytitle('${escapedListName}')/items`;
                const createPayload: Record<string, any> = {
                    Title: entry.title,
                    [schema.codeField as string]: entry.code,
                    [schema.maxSeatsField as string]: Number(entry.maxSeats || 0)
                };

                if (entry.provider && schema.providerField && this._isTextFieldType(schema.providerFieldType)) {
                    createPayload[schema.providerField] = entry.provider;
                }

                const createResponse = await this._getHttpClient().post(
                    createEndpoint,
                    SPHttpClient.configurations.v1,
                    {
                        headers: this._getJsonHeaders({
                            'X-RequestDigest': digest
                        }),
                        body: JSON.stringify(createPayload)
                    }
                );

                if (!createResponse.ok) {
                    const errorText = await this._readErrorBody(createResponse);
                    throw new Error(`Failed to create certification '${entry.title}' (HTTP ${createResponse.status} ${createResponse.statusText}): ${errorText.substring(0, 300) || 'No error details returned.'}`);
                }

                const createdItem = await this._readJson<any>(createResponse);
                const createdCertification: ICertificationMaxSeatsItem = {
                    id: Number(createdItem?.Id || createdItem?.id || 0),
                    title: entry.title,
                    code: entry.code,
                    maxSeats: Number(entry.maxSeats || 0),
                    provider: entry.provider || ''
                };
                createdCount += 1;
                itemsByCode.set(normalizedCode, createdCertification);
                itemsByTitle.set(normalizedTitle, createdCertification);
            }));
        }, 120);

        this._invalidateCertificationCountCache();

        return {
            createdCount,
            updatedCount,
            skippedCount,
            totalProcessed: entriesToProcess.length
        };
    }

    private static async _getEnrollmentListSchema(): Promise<IEnrollmentListSchema> {
        const siteUrl = this._ensureProductionSiteUrl();

        if (this._enrollmentListSchemaCache?.siteUrl === siteUrl) {
            return this._enrollmentListSchemaCache.schema;
        }

        if (this._enrollmentListSchemaPromise) {
            return this._enrollmentListSchemaPromise;
        }

        this._enrollmentListSchemaPromise = (async () => {
            const listName = await this._resolveEnrollmentListName();
            const escapedListName = this._escapeODataValue(listName);
            const fieldsData = await this._safeGetJson<any>(
                `${siteUrl}/_api/web/lists/getbytitle('${escapedListName}')/fields?$select=Title,InternalName,StaticName,Hidden,ReadOnlyField,TypeAsString,FieldTypeKind`,
                `${listName} fields`
            );
            const fields = this._toCollection(fieldsData) as ISharePointListFieldInfo[];
            const fieldTypes = this._getFieldTypeMap(fields);
            const schema: IEnrollmentListSchema = {
                assignedToField: this._pickListFieldName(fields, ['AssignedTo', 'Assigned To', 'User', 'Learner']),
                assignedByField: this._pickListFieldName(fields, ['AssignedBy', 'Assigned By']),
                certificationLookupField: this._pickListFieldNameByType(fields, ['CertificationId', 'Certification ID', 'Certification'], ['lookup', 'lookupmulti']),
                userEmailField: this._pickListFieldName(fields, ['UserEmail', 'User Email', 'LearnerEmail', 'AssignedToEmail', 'Assigned To Email', 'Email']),
                userNameField: this._pickListFieldName(fields, ['UserName', 'User Name', 'LearnerName', 'AssignedToName', 'Assigned To Name', 'Name']),
                certCodeField: this._pickListFieldName(fields, ['CertCode', 'Cert Code', 'CertificationCode', 'Certification Code', 'Code']),
                certNameField: this._pickListFieldName(fields, ['CertName', 'Cert Name', 'CertificationName', 'Certification Name']),
                startDateField: this._pickListFieldName(fields, ['StartDate', 'Start Date']),
                endDateField: this._pickListFieldName(fields, ['EndDate', 'End Date']),
                statusField: this._pickListFieldName(fields, ['Status']),
                progressField: this._pickListFieldName(fields, ['Progress']),
                certificateNameField: this._pickListFieldName(fields, ['CertificateName', 'Certificate Name', 'Certification']),
                assignedDateField: this._pickListFieldName(fields, ['AssignedDate', 'Assigned Date']),
                examScheduledDateField: this._pickListFieldName(fields, ['ExamScheduledDate', 'Exam Scheduled Date']),
                rescheduledDateField: this._pickListFieldName(fields, ['RescheduledDate', 'Rescheduled Date']),
                expiryDateField: this._pickListFieldName(fields, ['ExpiryDate', 'Expiry Date', 'ExpirationDate', 'Expiration Date', 'ValidTill', 'Valid Till']),
                completionDateField: this._pickListFieldName(fields, ['CompletionDate', 'Completion Date', 'CompletedOn', 'Completed On']),
                examCodeField: this._pickListFieldName(fields, ['ExamCode', 'Exam Code']),
                assignedByNameField: this._pickListFieldName(fields, ['AssignedByName', 'Assigned By Name']),
                assignedToEmailField: this._pickListFieldName(fields, ['AssignedToEmail', 'Assigned To Email', 'LearnerEmail']),
                pathIdField: this._pickListFieldName(fields, ['PathId', 'Path ID', 'CertificationId', 'Certification ID', 'CertCode', 'Cert Code', 'Code']),
                fieldTypes
            };

            console.log('[Enrollments] Resolved SharePoint field schema', {
                ...schema,
                fieldTypes
            });

            this._enrollmentListSchemaCache = { siteUrl, schema };
            this._enrollmentListSchemaPromise = null;
            return schema;
        })().catch((error) => {
            this._enrollmentListSchemaPromise = null;
            throw error;
        });

        return this._enrollmentListSchemaPromise;
    }

    private static async _getEnrollmentListContext(): Promise<{
        siteUrl: string;
        listName: string;
        escapedListName: string;
        schema: IEnrollmentListSchema;
    }> {
        const siteUrl = this._getSiteUrl();
        const listName = await this._resolveEnrollmentListName();
        const schema = await this._getEnrollmentListSchema();

        return {
            siteUrl,
            listName,
            escapedListName: this._escapeODataValue(listName),
            schema
        };
    }

    private static _getEnrollmentSelectFields(schema: IEnrollmentListSchema): string[] {
        return Array.from(new Set([
            'Id',
            'Title',
            'Created',
            schema.assignedDateField || '',
            schema.assignedByNameField || '',
            schema.examScheduledDateField || '',
            schema.rescheduledDateField || '',
            schema.expiryDateField || '',
            schema.completionDateField || '',
            schema.examCodeField || '',
            schema.statusField || '',
            schema.assignedToEmailField || '',
            schema.userEmailField || '',
            schema.userNameField || '',
            schema.certCodeField || '',
            schema.certNameField || '',
            schema.startDateField || '',
            schema.endDateField || '',
            schema.progressField || '',
            schema.certificateNameField || '',
            schema.pathIdField || '',
            schema.certificationLookupField ? `${schema.certificationLookupField}/Id` : '',
            schema.certificationLookupField ? `${schema.certificationLookupField}/Title` : '',
            schema.certificationLookupField ? `${schema.certificationLookupField}/Code` : '',
            schema.assignedToField ? `${schema.assignedToField}/Id` : '',
            schema.assignedToField ? `${schema.assignedToField}/Title` : '',
            schema.assignedToField ? `${schema.assignedToField}/EMail` : '',
            schema.assignedByField ? `${schema.assignedByField}/Id` : '',
            schema.assignedByField ? `${schema.assignedByField}/Title` : ''
        ].filter((field) => !!field)));
    }

    private static _getEnrollmentExpandFields(schema: IEnrollmentListSchema): string[] {
        return Array.from(new Set([
            schema.certificationLookupField || '',
            schema.assignedToField || '',
            schema.assignedByField || ''
        ].filter((field) => !!field)));
    }

    private static _buildEnrollmentItemsEndpoint(
        escapedListName: string,
        schema: IEnrollmentListSchema,
        options: {
            filters?: string[];
            orderByField?: string;
            top?: number;
        } = {}
    ): string {
        const selectFields = this._getEnrollmentSelectFields(schema);
        const expandFields = this._getEnrollmentExpandFields(schema);
        const queryParts: string[] = [
            `$select=${selectFields.join(',')}`
        ];

        if (expandFields.length > 0) {
            queryParts.push(`$expand=${expandFields.join(',')}`);
        }

        if (options.filters && options.filters.length > 0) {
            queryParts.push(`$filter=${options.filters.join(' and ')}`);
        }

        if (options.orderByField) {
            queryParts.push(`$orderby=${options.orderByField}`);
        }

        if (typeof options.top === 'number' && options.top > 0) {
            queryParts.push(`$top=${options.top}`);
        }

        return `${this._getSiteUrl()}/_api/web/lists/getbytitle('${escapedListName}')/items?${queryParts.join('&')}`;
    }

    private static _buildEnrollmentUserEmailFilter(schema: IEnrollmentListSchema, email: string): string[] {
        const normalizedEmail = (email || '').toString().trim().toLowerCase();
        if (!normalizedEmail) {
            return [];
        }

        return [
            schema.assignedToField ? `${schema.assignedToField}/EMail eq '${this._escapeODataValue(normalizedEmail)}'` : '',
            schema.assignedToEmailField ? `${schema.assignedToEmailField} eq '${this._escapeODataValue(normalizedEmail)}'` : '',
            schema.userEmailField ? `${schema.userEmailField} eq '${this._escapeODataValue(normalizedEmail)}'` : ''
        ].filter((value) => !!value);
    }

    private static _buildEnrollmentCertificationFilter(schema: IEnrollmentListSchema, certName: string): string[] {
        const normalizedCertName = (certName || '').toString().trim();
        if (!normalizedCertName) {
            return [];
        }

        const escapedCertName = this._escapeODataValue(normalizedCertName);
        return [
            schema.certNameField ? `${schema.certNameField} eq '${escapedCertName}'` : '',
            schema.certificateNameField ? `${schema.certificateNameField} eq '${escapedCertName}'` : '',
            `Title eq '${escapedCertName}'`
        ].filter((value, index, array) => !!value && array.indexOf(value) === index);
    }

    private static _buildEnrollmentCertificationCodeFilter(schema: IEnrollmentListSchema, certCode: string): string[] {
        const normalizedCertCode = (certCode || '').toString().trim();
        if (!normalizedCertCode) {
            return [];
        }

        const escapedCertCode = this._escapeODataValue(normalizedCertCode);
        return [
            schema.certCodeField ? `${schema.certCodeField} eq '${escapedCertCode}'` : '',
            schema.pathIdField ? `${schema.pathIdField} eq '${escapedCertCode}'` : ''
        ].filter((value, index, array) => !!value && array.indexOf(value) === index);
    }

    private static _buildEnrollmentCertificationIdFilter(schema: IEnrollmentListSchema, certificationId: number): string[] {
        const normalizedCertificationId = Number(certificationId || 0);
        if (normalizedCertificationId <= 0 || !schema.certificationLookupField) {
            return [];
        }

        return [
            `${schema.certificationLookupField}/Id eq ${normalizedCertificationId}`,
            `${schema.certificationLookupField}Id eq ${normalizedCertificationId}`
        ].filter((value, index, array) => !!value && array.indexOf(value) === index);
    }

    private static _buildEnrollmentSearchFilter(schema: IEnrollmentListSchema, searchText: string): string {
        const normalizedSearch = (searchText || '').toString().trim();
        if (!normalizedSearch) {
            return '';
        }

        const escapedSearch = this._escapeODataValue(normalizedSearch);
        const parts = [
            schema.certificationLookupField ? `substringof('${escapedSearch}', ${schema.certificationLookupField}/Title)` : '',
            schema.certificationLookupField ? `substringof('${escapedSearch}', ${schema.certificationLookupField}/Code)` : '',
            schema.certNameField ? `substringof('${escapedSearch}', ${schema.certNameField})` : '',
            schema.certCodeField ? `substringof('${escapedSearch}', ${schema.certCodeField})` : '',
            `substringof('${escapedSearch}', Title)`
        ].filter((value, index, array) => !!value && array.indexOf(value) === index);

        return parts.length > 0 ? `(${parts.join(' or ')})` : '';
    }

    private static _buildEnrollmentStatusExclusionFilter(schema: IEnrollmentListSchema, statuses: string[] = []): string {
        const statusField = schema.statusField;
        const normalizedStatuses = (statuses || [])
            .map((value) => (value || '').toString().trim())
            .filter((value, index, array) => !!value && array.indexOf(value) === index);

        if (!statusField || normalizedStatuses.length === 0) {
            return '';
        }

        const parts = normalizedStatuses.map((status) => `${statusField} ne '${this._escapeODataValue(status)}'`);
        return parts.length > 0 ? `(${parts.join(' and ')})` : '';
    }

    private static async _getAuditLogListSchema(): Promise<IAuditLogListSchema> {
        const siteUrl = this._ensureProductionSiteUrl();

        if (this._auditLogListSchemaCache?.siteUrl === siteUrl) {
            return this._auditLogListSchemaCache.schema;
        }

        if (this._auditLogListSchemaPromise) {
            return this._auditLogListSchemaPromise;
        }

        this._auditLogListSchemaPromise = (async () => {
            const listName = await this._resolveAuditLogListName();
            const fieldsData = await this._safeGetJson<any>(
                `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/fields?$select=Title,InternalName,StaticName,Hidden,ReadOnlyField`,
                `${listName} fields`
            );
            const fields = this._toCollection(fieldsData) as ISharePointListFieldInfo[];
            const schema: IAuditLogListSchema = {
                userField: this._pickListFieldName(fields, ['UserId', 'User ID', 'User', 'Learner', 'AssignedTo', 'Assigned To']),
                learnerEmailField: this._pickListFieldName(fields, ['LearnerEmail', 'Learner Email', 'UserEmail', 'User Email', 'Email']),
                learnerNameField: this._pickListFieldName(fields, ['LearnerName', 'Learner Name', 'UserName', 'User Name', 'Name']),
                actionField: this._pickListFieldName(fields, ['Action']),
                assignmentNameField: this._pickListFieldName(fields, ['AssignmentName', 'Assignment Name', 'CertificationName', 'Certification Name', 'PathName', 'Path Name']),
                assignmentDateField: this._pickListFieldName(fields, ['AssignmentDate', 'Assignment Date', 'AssignedDate', 'Assigned Date', 'Timestamp']),
                assignedByField: this._pickListFieldName(fields, ['AssignedBy', 'Assigned By']),
                statusField: this._pickListFieldName(fields, ['Status']),
                pathIdField: this._pickListFieldName(fields, ['PathId', 'Path ID']),
                timestampField: this._pickListFieldName(fields, ['Timestamp', 'AssignmentDate', 'Assignment Date', 'AssignedDate', 'Assigned Date'])
            };

            console.log('[Audit Logs] Resolved SharePoint field schema', schema);

            this._auditLogListSchemaCache = { siteUrl, schema };
            this._auditLogListSchemaPromise = null;
            return schema;
        })().catch((error) => {
            this._auditLogListSchemaPromise = null;
            throw error;
        });

        return this._auditLogListSchemaPromise;
    }

    private static _rememberListTitle(listName: string): void {
        const siteUrl = this._getSiteUrl();
        if (this._listTitlesCache?.siteUrl !== siteUrl) {
            return;
        }

        const normalizedTitle = this._normalizeListTitle(listName);
        this._listTitlesCache.titles.add(normalizedTitle);
        this._listTitlesCache.titleMap.set(normalizedTitle, listName);
    }

    private static async _getExistingListTitles(forceRefresh: boolean = false): Promise<Set<string>> {
        const siteUrl = this._getSiteUrl();

        if (!forceRefresh && this._listTitlesCache?.siteUrl === siteUrl) {
            return this._listTitlesCache.titles;
        }

        if (!forceRefresh && this._listTitlesPromise) {
            return this._listTitlesPromise;
        }

        this._listTitlesPromise = (async () => {
            const data = await this._safeGetJson<any>(
                `${siteUrl}/_api/web/lists?$select=Title`,
                'SharePoint list titles'
            );

            const titleMap = new Map<string, string>();
            const titles = new Set<string>();
            this._toCollection(data).forEach((item: any) => {
                const actualTitle = (item?.Title || '').toString().trim();
                const normalizedTitle = this._normalizeListTitle(actualTitle);
                if (!normalizedTitle) {
                    return;
                }

                titles.add(normalizedTitle);
                if (!titleMap.has(normalizedTitle)) {
                    titleMap.set(normalizedTitle, actualTitle);
                }
            });

            this._listTitlesCache = {
                siteUrl,
                titles,
                titleMap
            };

            this._listTitlesPromise = null;
            return titles;
        })().catch((error) => {
            this._listTitlesPromise = null;
            throw error;
        });

        return this._listTitlesPromise;
    }

    private static async _getExistingListTitleMap(forceRefresh: boolean = false): Promise<Map<string, string>> {
        const siteUrl = this._getSiteUrl();
        if (!forceRefresh && this._listTitlesCache?.siteUrl === siteUrl) {
            return this._listTitlesCache.titleMap;
        }

        await this._getExistingListTitles(forceRefresh);
        return this._listTitlesCache?.titleMap || new Map<string, string>();
    }

    private static async _resolveExistingListName(
        candidates: string[],
        fallbackListName: string,
        options?: {
            cacheKey?: '_enrollmentListNameCache' | '_userNotificationListNameCache' | '_assessmentAssignmentListNameCache' | '_contentLibraryListNameCache' | '_auditLogListNameCache' | '_learningAndSkillsListNameCache' | '_certificationCompletionListNameCache';
            forceRefresh?: boolean;
            throwIfMissing?: boolean;
            label?: string;
        }
    ): Promise<string> {
        const cacheKey = options?.cacheKey;
        const forceRefresh = !!options?.forceRefresh;

        if (!forceRefresh && cacheKey) {
            const cachedValue = this._getResolvedListNameCache(cacheKey);
            if (cachedValue) {
                return cachedValue;
            }
        }

        const titleMap = await this._getExistingListTitleMap(forceRefresh);
        const matchedListName = candidates
            .map((candidate) => titleMap.get(this._normalizeListTitle(candidate)) || '')
            .find((value) => !!value) || '';

        if (matchedListName) {
            if (cacheKey) {
                this._setResolvedListNameCache(cacheKey, matchedListName);
            }
            return matchedListName;
        }

        if (options?.throwIfMissing) {
            const errorMessage = `${options?.label || fallbackListName} list not found. Checked: ${candidates.join(', ')}`;
            console.error(`[Lists] ${errorMessage}`);
            throw new Error(errorMessage);
        }

        if (cacheKey) {
            this._setResolvedListNameCache(cacheKey, fallbackListName);
        }

        return fallbackListName;
    }

    private static _getResolvedListNameCache(
        cacheKey: '_enrollmentListNameCache' | '_userNotificationListNameCache' | '_assessmentAssignmentListNameCache' | '_contentLibraryListNameCache' | '_auditLogListNameCache' | '_learningAndSkillsListNameCache' | '_certificationCompletionListNameCache'
    ): string | null {
        switch (cacheKey) {
            case '_enrollmentListNameCache':
                return this._enrollmentListNameCache;
            case '_userNotificationListNameCache':
                return this._userNotificationListNameCache;
            case '_assessmentAssignmentListNameCache':
                return this._assessmentAssignmentListNameCache;
            case '_contentLibraryListNameCache':
                return this._contentLibraryListNameCache;
            case '_auditLogListNameCache':
                return this._auditLogListNameCache;
            case '_learningAndSkillsListNameCache':
                return this._learningAndSkillsListNameCache;
            case '_certificationCompletionListNameCache':
                return this._certificationCompletionListNameCache;
            default:
                return null;
        }
    }

    private static _setResolvedListNameCache(
        cacheKey: '_enrollmentListNameCache' | '_userNotificationListNameCache' | '_assessmentAssignmentListNameCache' | '_contentLibraryListNameCache' | '_auditLogListNameCache' | '_learningAndSkillsListNameCache' | '_certificationCompletionListNameCache',
        value: string
    ): void {
        switch (cacheKey) {
            case '_enrollmentListNameCache':
                this._enrollmentListNameCache = value;
                return;
            case '_userNotificationListNameCache':
                this._userNotificationListNameCache = value;
                return;
            case '_assessmentAssignmentListNameCache':
                this._assessmentAssignmentListNameCache = value;
                return;
            case '_contentLibraryListNameCache':
                this._contentLibraryListNameCache = value;
                return;
            case '_auditLogListNameCache':
                this._auditLogListNameCache = value;
                return;
            case '_learningAndSkillsListNameCache':
                this._learningAndSkillsListNameCache = value;
                return;
            case '_certificationCompletionListNameCache':
                this._certificationCompletionListNameCache = value;
                return;
            default:
                return;
        }
    }

    private static _ensureProductionSiteUrl(): string {
        const siteUrl = this._getSiteUrl();
        if (!siteUrl || siteUrl.toLowerCase().indexOf('localhost') !== -1) {
            throw new Error('SharePoint group membership is available only when this web part runs on a production SharePoint site.');
        }

        return siteUrl;
    }

    private static _getMembershipIdentityKeys(user: any): string[] {
        return Array.from(
            new Set(
                this._getUserLookupKeys(user)
                    .map((key) => key.toLowerCase())
                    .filter((key) => !!key)
            ).values()
        );
    }

    private static _getMembershipRolePriority(role?: string): number {
        switch (role) {
            case 'Owner':
                return 0;
            case 'Member':
                return 1;
            case 'Visitor':
                return 2;
            default:
                return 99;
        }
    }

    private static _isGenericSiteUser(user?: ILearnerDirectoryUser): boolean {
        const siteGroup = (user?.siteGroup || user?.group || '').toString().trim().toLowerCase();
        return siteGroup === 'all site users';
    }

    private static _mergeMembershipUser(existing: ILearnerDirectoryUser, incoming: ILearnerDirectoryUser): ILearnerDirectoryUser {
        const existingIsGeneric = this._isGenericSiteUser(existing);
        const incomingIsGeneric = this._isGenericSiteUser(incoming);

        let preferredUser = existing;
        let secondaryUser = incoming;

        if (existingIsGeneric !== incomingIsGeneric) {
            preferredUser = existingIsGeneric ? incoming : existing;
            secondaryUser = preferredUser === incoming ? existing : incoming;
        } else {
            const existingPriority = this._getMembershipRolePriority(existing.role);
            const incomingPriority = this._getMembershipRolePriority(incoming.role);
            preferredUser = incomingPriority < existingPriority ? incoming : existing;
            secondaryUser = preferredUser === incoming ? existing : incoming;
        }

        return {
            ...secondaryUser,
            ...preferredUser,
            jobTitle: preferredUser.jobTitle || secondaryUser.jobTitle || '',
            department: preferredUser.department || secondaryUser.department || '',
            role: preferredUser.role,
            group: preferredUser.group || secondaryUser.group,
            siteGroup: preferredUser.siteGroup || secondaryUser.siteGroup
        };
    }

    private static _sortMembershipUsers(users: ILearnerDirectoryUser[]): ILearnerDirectoryUser[] {
        return [...users].sort((left, right) => {
            const roleDelta = this._getMembershipRolePriority(left.role) - this._getMembershipRolePriority(right.role);
            if (roleDelta !== 0) {
                return roleDelta;
            }

            const leftName = (left.name || left.Title || '').toString();
            const rightName = (right.name || right.Title || '').toString();
            return leftName.localeCompare(rightName);
        });
    }

    private static _mergeMembershipGroups(groups: ILearnerDirectoryUser[][]): ILearnerDirectoryUser[] {
        const learners: ILearnerDirectoryUser[] = [];
        const learnersByKey = new Map<string, number>();

        groups.forEach((groupUsers) => {
            groupUsers.forEach((user) => {
                const dedupeKeys = this._getMembershipIdentityKeys(user);
                if (dedupeKeys.length === 0) {
                    return;
                }

                const existingLearnerIndex = dedupeKeys
                    .map((key) => learnersByKey.get(key))
                    .find((value) => value !== undefined);

                if (existingLearnerIndex === undefined) {
                    const nextIndex = learners.push(user) - 1;
                    dedupeKeys.forEach((key) => learnersByKey.set(key, nextIndex));
                    return;
                }

                learners[existingLearnerIndex] = this._mergeMembershipUser(learners[existingLearnerIndex], user);
                this._getMembershipIdentityKeys(learners[existingLearnerIndex]).forEach((key) => {
                    learnersByKey.set(key, existingLearnerIndex);
                });
            });
        });

        return this._sortMembershipUsers(learners);
    }

    private static _mapMembershipUser(user: any, siteGroup: 'Owners' | 'Members' | 'Visitors', role: 'Owner' | 'Member' | 'Visitor'): ILearnerDirectoryUser | null {
        const normalizedUser = this._normalizeDirectoryUser(user, role, siteGroup);
        if (!normalizedUser) {
            return null;
        }

        return {
            ...normalizedUser,
            group: siteGroup,
            siteGroup,
            role
        };
    }

    private static _getUserLookupKeys(user: any): string[] {
        const keys = new Set<string>();
        const email = this._extractUserEmail(user).toLowerCase();
        const login = (user?.LoginName || user?.login || '').toString().trim().toLowerCase();

        if (email) {
            keys.add(email);
        }

        if (login) {
            keys.add(login);

            if (login.indexOf('|') !== -1) {
                const loginTail = login.split('|').pop()?.trim() || '';
                if (loginTail) {
                    keys.add(loginTail);
                }
            }
        }

        return Array.from(keys.values());
    }

    private static async _fetchSiteUserProfiles(): Promise<Map<string, { jobTitle: string; department: string }>> {
        const siteUrl = this._getSiteUrl();
        const endpoint = `${siteUrl}/_api/web/siteusers?$select=Id,Title,Email,LoginName,JobTitle`;

        try {
            const data = await this._safeGetJson<any>(endpoint, 'SharePoint site user profiles');
            const profileMap = new Map<string, { jobTitle: string; department: string }>();

            this._toCollection(data).forEach((user: any) => {
                const profile = {
                    jobTitle: user?.JobTitle || '',
                    department: ''
                };

                this._getUserLookupKeys(user).forEach((key) => {
                    profileMap.set(key, profile);
                });
            });

            return profileMap;
        } catch (error) {
            console.error('[Learners] SharePoint site user profile request threw an error', {
                endpoint,
                error
            });
            return new Map<string, { jobTitle: string; department: string }>();
        }
    }

    private static _getUserProfileProperty(profile: any, key: string): string {
        const normalizedKey = (key || '').toString().trim().toLowerCase();
        const properties = Array.isArray(profile?.UserProfileProperties) ? profile.UserProfileProperties : [];
        const matchedProperty = properties.find((property: any) =>
            (property?.Key || '').toString().trim().toLowerCase() === normalizedKey
        );

        return (matchedProperty?.Value || '').toString().trim();
    }

    private static _getPeopleManagerAccountName(user: ILearnerDirectoryUser): string {
        const loginName = (user?.login || user?.LoginName || '').toString().trim();
        if (loginName) {
            return loginName;
        }

        return (user?.email || user?.Email || '').toString().trim();
    }

    private static async _fetchPeopleManagerUserProfiles(users: ILearnerDirectoryUser[]): Promise<Map<string, { jobTitle: string; department: string }>> {
        const siteUrl = this._getSiteUrl();
        const profileMap = new Map<string, { jobTitle: string; department: string }>();
        const uniqueUsers = Array.from(
            new Map(
                (users || [])
                    .filter((user) => !!this._getPeopleManagerAccountName(user))
                    .map((user) => [this._getPeopleManagerAccountName(user).toLowerCase(), user])
            ).values()
        );

        if (uniqueUsers.length === 0) {
            return profileMap;
        }

        await this._processInChunks(uniqueUsers, 10, async (chunk) => {
            await Promise.all(chunk.map(async (user) => {
                const accountName = this._getPeopleManagerAccountName(user);
                if (!accountName) {
                    return;
                }

                const endpoint = `${siteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodeURIComponent(accountName)}'`;

                try {
                    const profile = await this._safeGetJson<any>(endpoint, `PeopleManager profile for ${accountName}`);
                    if (!profile) {
                        return;
                    }

                    const mappedProfile = {
                        jobTitle: this._getUserProfileProperty(profile, 'Title'),
                        department: this._getUserProfileProperty(profile, 'Department')
                    };

                    this._getUserLookupKeys(user).forEach((key) => {
                        profileMap.set(key, mappedProfile);
                    });
                } catch (error) {
                    console.warn('[Learners] PeopleManager profile lookup failed for user', {
                        accountName,
                        error
                    });
                }
            }));
        });

        return profileMap;
    }

    private static async _enrichDirectoryUsers(users: ILearnerDirectoryUser[]): Promise<ILearnerDirectoryUser[]> {
        const mergedUsers = this._mergeMembershipGroups([users || []]);
        if (mergedUsers.length === 0) {
            return [];
        }

        const [peopleManagerProfiles, siteUserProfiles] = await Promise.all([
            this._fetchPeopleManagerUserProfiles(mergedUsers),
            this._fetchSiteUserProfiles()
        ]);

        return mergedUsers.map((user) => {
            const matchingProfile = this._getUserLookupKeys(user)
                .map((key) => peopleManagerProfiles.get(key) || siteUserProfiles.get(key))
                .find((profile) => !!profile);

            if (!matchingProfile) {
                return user;
            }

            return {
                ...user,
                jobTitle: matchingProfile.jobTitle || user.jobTitle || '',
                department: matchingProfile.department || user.department || ''
            };
        });
    }

    private static async _fetchNamedSiteGroupUsers(groupName: string): Promise<ILearnerDirectoryUser[]> {
        const normalizedGroupName = (groupName || '').toString().trim();
        if (!normalizedGroupName) {
            return [];
        }

        const endpoint = this._getApiUrl(
            `/_api/web/sitegroups/getbyname('${this._escapeODataValue(normalizedGroupName)}')/users?$select=Id,Title,Email,LoginName`
        );

        try {
            const data = await this._safeGetJson<any>(endpoint, `${normalizedGroupName} group users`);
            return this._toCollection(data)
                .map((user: any) => {
                    const normalizedUser = this._normalizeDirectoryUser(user, 'Member', normalizedGroupName);
                    if (!normalizedUser) {
                        return null;
                    }

                    return {
                        ...normalizedUser,
                        group: normalizedGroupName,
                        siteGroup: normalizedGroupName,
                        role: 'Member'
                    } as ILearnerDirectoryUser;
                })
                .filter((user: ILearnerDirectoryUser | null): user is ILearnerDirectoryUser => !!user);
        } catch (error) {
            this._logResponseIssueOnce(
                `site-group-by-name:${normalizedGroupName}`,
                `[Learners] SharePoint group '${normalizedGroupName}' is unavailable. Falling back to the synced learner directory.`,
                {
                    endpoint,
                    error
                }
            );
            return [];
        }
    }

    private static async _getAssociatedSiteGroups(): Promise<Array<{
        groupId: number;
        siteGroup: 'Owners' | 'Members' | 'Visitors';
        role: 'Owner' | 'Member' | 'Visitor';
    }>> {
        const endpoint = this._getApiUrl(
            `/_api/web?$select=AssociatedMemberGroup/Id,AssociatedOwnerGroup/Id,AssociatedVisitorGroup/Id&$expand=AssociatedMemberGroup,AssociatedOwnerGroup,AssociatedVisitorGroup`
        );

        try {
            const data = await this._safeGetJson<any>(endpoint, 'associated SharePoint site groups');
            const groups = [
                {
                    groupId: Number(data?.AssociatedMemberGroup?.Id || 0),
                    siteGroup: 'Members' as const,
                    role: 'Member' as const
                },
                {
                    groupId: Number(data?.AssociatedVisitorGroup?.Id || 0),
                    siteGroup: 'Visitors' as const,
                    role: 'Visitor' as const
                },
                {
                    groupId: Number(data?.AssociatedOwnerGroup?.Id || 0),
                    siteGroup: 'Owners' as const,
                    role: 'Owner' as const
                }
            ].filter((group) => Number.isFinite(group.groupId) && group.groupId > 0);

            if (groups.length > 0) {
                return groups;
            }
        } catch (error) {
            console.warn('[Learners] Failed to resolve associated SharePoint site groups, using fallback ids', {
                endpoint,
                error
            });
        }

        return this._assessmentAssignmentGroups;
    }

    private static async _fetchAllSiteUsers(): Promise<ILearnerDirectoryUser[]> {
        const siteUrl = this._getSiteUrl();
        const endpoint = `${siteUrl}/_api/web/siteusers?$select=Id,Title,Email,LoginName,JobTitle,IsSiteAdmin,PrincipalType`;

        try {
            const data = await this._safeGetJson<any>(endpoint, 'SharePoint site users');
            return this._toCollection(data)
                .map((user: any) => {
                    const principalType = Number(user?.PrincipalType || 0);
                    if (principalType && principalType !== 1) {
                        return null;
                    }

                    const normalizedUser = this._normalizeDirectoryUser(user);
                    if (!normalizedUser) {
                        return null;
                    }

                    return {
                        ...normalizedUser,
                        group: 'All Site Users',
                        siteGroup: 'All Site Users',
                        role: user?.IsSiteAdmin ? 'Owner' : 'Member',
                        jobTitle: (user?.JobTitle || '').toString().trim()
                    } as ILearnerDirectoryUser;
                })
                .filter((user: ILearnerDirectoryUser | null): user is ILearnerDirectoryUser => !!user);
        } catch (error) {
            console.error('[Learners] SharePoint site users request threw an error', {
                endpoint,
                error
            });
            return [];
        }
    }

    private static async _fetchSiteGroupUsers(
        groupId: number,
        siteGroup: 'Owners' | 'Members' | 'Visitors',
        role: 'Owner' | 'Member' | 'Visitor'
    ): Promise<ILearnerDirectoryUser[]> {
        const endpoint = this._getApiUrl(`/_api/web/sitegroups/getbyid(${groupId})/users?$select=Id,Title,Email,LoginName`);

        try {
            const data = await this._safeGetJson<any>(endpoint, `${siteGroup} group users`);
            return this._toCollection(data)
                .map((user: any) => this._mapMembershipUser(user, siteGroup, role))
                .filter((user: ILearnerDirectoryUser | null): user is ILearnerDirectoryUser => !!user);
        } catch (error) {
            console.error('[Access] SharePoint site group request threw an error', {
                endpoint,
                groupId,
                siteGroup,
                role,
                error
            });
            return [];
        }
    }

    public static async getDefaultSiteGroupMembership(forceRefresh: boolean = false): Promise<ISiteMembershipSnapshot> {
        if (!this._spHttpClient) {
            throw new Error('SharePointService not initialized. Call init() with context.spHttpClient first.');
        }

        const siteUrl = this._getSiteUrl();
        if (!forceRefresh) {
            const cachedSnapshot = this._readMembershipSnapshotCache(siteUrl);
            if (cachedSnapshot) {
                return cachedSnapshot;
            }

            if (this._membershipSnapshotPromise) {
                return this._membershipSnapshotPromise;
            }
        }

        const membershipPromise = (async () => {
            const siteGroups = await this._getAssociatedSiteGroups();
            const groupedUsers = await Promise.all(
                siteGroups.map((groupConfig) =>
                    this._fetchSiteGroupUsers(groupConfig.groupId, groupConfig.siteGroup, groupConfig.role)
                )
            );
            const owners = groupedUsers[siteGroups.findIndex((group) => group.siteGroup === 'Owners')] || [];
            const members = groupedUsers[siteGroups.findIndex((group) => group.siteGroup === 'Members')] || [];
            const visitors = groupedUsers[siteGroups.findIndex((group) => group.siteGroup === 'Visitors')] || [];
            const flattenedUsers = this._mergeMembershipGroups([owners, members, visitors]);
            const [peopleManagerProfiles, siteUserProfiles] = await Promise.all([
                this._fetchPeopleManagerUserProfiles(flattenedUsers),
                this._fetchSiteUserProfiles()
            ]);

            const applyProfiles = (users: ILearnerDirectoryUser[]): ILearnerDirectoryUser[] =>
                users.map((user: any) => {
                    const matchingProfile = this._getUserLookupKeys(user)
                        .map((key) => peopleManagerProfiles.get(key) || siteUserProfiles.get(key))
                        .find((profile) => !!profile);

                    if (!matchingProfile) {
                        return user;
                    }

                    return {
                        ...user,
                        jobTitle: matchingProfile.jobTitle || user.jobTitle || '',
                        department: matchingProfile.department || user.department || ''
                    };
                });

            const ownersWithProfiles = applyProfiles(owners);
            const membersWithProfiles = applyProfiles(members);
            const visitorsWithProfiles = applyProfiles(visitors);

            const snapshot: ISiteMembershipSnapshot = {
                owners: this._sortMembershipUsers(ownersWithProfiles),
                members: this._sortMembershipUsers(membersWithProfiles),
                visitors: this._sortMembershipUsers(visitorsWithProfiles),
                learners: this._mergeMembershipGroups([ownersWithProfiles, membersWithProfiles, visitorsWithProfiles])
            };

            this._writeMembershipSnapshotCache(siteUrl, snapshot);
            return snapshot;
        })();

        this._membershipSnapshotPromise = membershipPromise;

        try {
            return await membershipPromise;
        } finally {
            if (this._membershipSnapshotPromise === membershipPromise) {
                this._membershipSnapshotPromise = null;
            }
        }
    }

    private static async _getCurrentUserGroupTitles(): Promise<string[]> {
        const siteUrl = this._getSiteUrl();
        const primaryEndpoint =
            `${siteUrl}/_api/web/currentuser?$select=Id,Title,Email,LoginName,Groups/Id,Groups/Title&$expand=Groups`;
        const fallbackEndpoint =
            `${siteUrl}/_api/web/currentuser/groups?$select=Id,Title`;

        const extractGroupTitles = (payload: any): string[] => {
            const groups = this._toCollection(
                payload?.Groups ||
                payload?.groups ||
                payload?.value ||
                payload?.d?.results ||
                payload?.d?.Groups
            );

            return groups
                .map((group: any) => (group?.Title || '').toString().trim())
                .filter((title: string) => !!title);
        };

        try {
            const primaryData = await this._safeGetJson<any>(primaryEndpoint, 'current user groups');
            const primaryTitles = extractGroupTitles(primaryData);
            if (primaryTitles.length > 0) {
                return primaryTitles;
            }
        } catch (error) {
            console.warn('[Access] Primary current user groups lookup failed', {
                endpoint: primaryEndpoint,
                error
            });
        }

        const fallbackData = await this._safeGetJson<any>(fallbackEndpoint, 'current user groups fallback');
        return extractGroupTitles(fallbackData);
    }

    private static _getContextBasedAdminAccess(): { isOwner: boolean; isMember: boolean; isVisitor: boolean; canAccessAdmin: boolean; currentUserRole: 'Owner' | 'Member' | 'Visitor' | 'Learner' | 'Unknown'; } {
        const permissions = this._context?.pageContext?.web?.permissions;
        const hasPermission = permissions?.hasPermission?.bind(permissions);

        if (!hasPermission) {
            return {
                isOwner: false,
                isMember: false,
                isVisitor: false,
                canAccessAdmin: false,
                currentUserRole: 'Unknown'
            };
        }

        const canManageWeb = hasPermission(SPPermission.manageWeb);
        const canEditContent =
            hasPermission(SPPermission.editListItems) ||
            hasPermission(SPPermission.addListItems) ||
            hasPermission(SPPermission.deleteListItems);
        const canViewPages = hasPermission(SPPermission.viewPages);

        const isOwner = !!canManageWeb;
        const isMember = !isOwner && !!canEditContent;
        const isVisitor = !isOwner && !isMember && !!canViewPages;
        const currentUserRole =
            isOwner ? 'Owner' :
            isMember ? 'Member' :
            isVisitor ? 'Visitor' :
            'Learner';

        return {
            isOwner,
            isMember,
            isVisitor,
            canAccessAdmin: isOwner || isMember,
            currentUserRole
        };
    }

    public static async getCurrentUserAdminAccess(userEmail?: string, forceRefresh: boolean = false): Promise<IAdminPortalAccessState> {
        const fallbackState: IAdminPortalAccessState = {
            owners: [],
            members: [],
            visitors: [],
            learners: [],
            currentUserRole: 'Unknown',
            canAccessAdmin: false,
            accessCheckFailed: false
        };

        try {
            const membershipSnapshot = forceRefresh
                ? await this.getDefaultSiteGroupMembership(true).catch(() => fallbackState)
                : (this._readMembershipSnapshotCache(this._getSiteUrl()) || fallbackState);
            const currentUserGroupTitles = await this._getCurrentUserGroupTitles();
            const normalizedGroupTitles = currentUserGroupTitles.map((title) => title.toLowerCase());
            const permissionAccess = this._getContextBasedAdminAccess();
            const isOwner = normalizedGroupTitles.some((title) => title.indexOf('owner') !== -1) || permissionAccess.isOwner;
            const isMember = (!isOwner && normalizedGroupTitles.some((title) => title.indexOf('member') !== -1)) || permissionAccess.isMember;
            const isVisitor = !isOwner && !isMember && (
                normalizedGroupTitles.some((title) => title.indexOf('visitor') !== -1) ||
                permissionAccess.isVisitor
            );

            const currentUserRole =
                isOwner ? 'Owner' :
                isMember ? 'Member' :
                isVisitor ? 'Visitor' :
                permissionAccess.currentUserRole !== 'Unknown' ? permissionAccess.currentUserRole :
                'Learner';

            console.log('[Access] Current user group titles', {
                userEmail: (userEmail || this._context?.pageContext?.user?.email || '').toString().trim().toLowerCase(),
                currentUserGroupTitles,
                normalizedGroupTitles,
                isOwner,
                isMember,
                isVisitor,
                permissionAccess,
                currentUserRole
            });

            return {
                ...membershipSnapshot,
                currentUserRole,
                canAccessAdmin: isOwner || isMember || permissionAccess.canAccessAdmin,
                accessCheckFailed: false
            };
        } catch (error) {
            console.error('[Access] Failed to determine current user admin access', error);
            return {
                ...fallbackState,
                accessCheckFailed: true
            };
        }
    }

    private static async _resolveTaxonomyListName(): Promise<string> {
        return this._resolveExistingListName(
            this._taxonomyListCandidates,
            this._taxonomyListCandidates[0]
        );
    }

    private static async _resolveEnrollmentListName(): Promise<string> {
        this._enrollmentListNameCache = await this._resolveExistingListName(
            this._enrollmentListCandidates,
            this._enrollmentListCandidates[0],
            {
                cacheKey: '_enrollmentListNameCache'
            }
        );
        console.log('[Enrollments] Resolved list name', this._enrollmentListNameCache);
        return this._enrollmentListNameCache;
    }

    private static async _resolveAssignmentNotificationListName(): Promise<string> {
        return this._resolveEnrollmentListName();
    }

    private static async _resolveUserNotificationListName(): Promise<string> {
        this._userNotificationListNameCache = await this._resolveExistingListName(
            this._userNotificationListCandidates,
            'LMS_Notifications',
            {
                cacheKey: '_userNotificationListNameCache'
            }
        );
        return this._userNotificationListNameCache;
    }

    private static async _resolveContentLibraryListName(): Promise<string> {
        if (this._contentLibraryListNameCache) {
            return this._contentLibraryListNameCache;
        }

        const siteUrl = this._getSiteUrl();
        const listsData = await this._safeGetJson<any>(
            `${siteUrl}/_api/web/lists?$select=Title,BaseTemplate,Hidden`,
            'content library list candidates'
        );
        const siteLists = this._toCollection(listsData);
        const genericListCandidate = siteLists.find((list: any) => {
            const normalizedTitle = this._normalizeListTitle(list?.Title || '');
            const isSupportedTitle =
                this._contentLibraryListCandidates
                    .map((candidate) => this._normalizeListTitle(candidate))
                    .indexOf(normalizedTitle) !== -1;

            return isSupportedTitle && Number(list?.BaseTemplate || 0) === 100 && list?.Hidden !== true;
        });

        this._contentLibraryListNameCache = (genericListCandidate?.Title || this._contentLibraryListName || '').toString();
        return this._contentLibraryListNameCache || this._contentLibraryListName;
    }

    private static async _resolveAssessmentAssignmentListName(): Promise<string> {
        this._assessmentAssignmentListNameCache = await this._resolveExistingListName(
            this._assessmentAssignmentListCandidates,
            this._assessmentAssignmentListCandidates[0],
            {
                cacheKey: '_assessmentAssignmentListNameCache'
            }
        );
        return this._assessmentAssignmentListNameCache;
    }

    private static async _resolveAuditLogListName(): Promise<string> {
        this._auditLogListNameCache = await this._resolveExistingListName(
            this._auditLogListCandidates,
            this._auditLogListName,
            {
                cacheKey: '_auditLogListNameCache'
            }
        );
        return this._auditLogListNameCache;
    }

    private static async _resolveLearningAndSkillsListName(forceRefresh: boolean = false): Promise<string> {
        return this._resolveExistingListName(
            this._learningAndSkillsListCandidates,
            this._learningAndSkillsListCandidates[0],
            {
                cacheKey: '_learningAndSkillsListNameCache',
                forceRefresh,
                throwIfMissing: true,
                label: 'Learning and Skills'
            }
        );
    }

    private static async _resolveCertificationCompletionListName(forceRefresh: boolean = false): Promise<string> {
        this._certificationCompletionListNameCache = await this._resolveExistingListName(
            this._certificationCompletionListCandidates,
            this._certificationCompletionListCandidates[0],
            {
                cacheKey: '_certificationCompletionListNameCache',
                forceRefresh
            }
        );

        return this._certificationCompletionListNameCache;
    }

    private static async _ensureTextFieldOnList(listName: string, internalName: string, displayName: string = internalName): Promise<void> {
        const siteUrl = this._getSiteUrl();
        const escapedListName = this._escapeODataValue(listName);
        const fieldsData = await this._safeGetJson<any>(
            `${siteUrl}/_api/web/lists/getbytitle('${escapedListName}')/fields?$select=Title,InternalName,StaticName,Hidden,ReadOnlyField`,
            `${listName} fields`
        );
        const fields = this._toCollection(fieldsData) as ISharePointListFieldInfo[];
        const normalizedInternalName = this._normalizeFieldKey(internalName);
        const existingField = fields.find((field) => {
            const fieldNames = [field.InternalName, field.StaticName, field.Title]
                .map((value) => this._normalizeFieldKey((value || '').toString()))
                .filter((value) => !!value);
            return fieldNames.indexOf(normalizedInternalName) !== -1;
        });

        if (existingField) {
            return;
        }

        const digest = await this._getFormDigestValue();
        const schemaXml = `<Field Type="Text" Name="${this._escapeXmlAttribute(internalName)}" StaticName="${this._escapeXmlAttribute(internalName)}" DisplayName="${this._escapeXmlAttribute(displayName)}" MaxLength="255" Group="LMS Generated Columns" />`;
        const response = await this._getHttpClient().post(
            `${siteUrl}/_api/web/lists/getbytitle('${escapedListName}')/fields/createfieldasxml`,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'X-RequestDigest': digest
                }),
                body: JSON.stringify({
                    parameters: {
                        SchemaXml: schemaXml,
                        Options: 0
                    }
                })
            }
        );

        if (!response.ok) {
            const errorText = await this._readErrorBody(response);
            if ((errorText || '').toLowerCase().indexOf('duplicate') !== -1 || (errorText || '').toLowerCase().indexOf('already exists') !== -1) {
                return;
            }

            throw new Error(`Failed to ensure ${internalName} on ${listName} (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 200) || 'No error details returned.'}`);
        }
    }

    private static async _ensureDateFieldOnList(listName: string, internalName: string, displayName: string = internalName): Promise<void> {
        const siteUrl = this._getSiteUrl();
        const escapedListName = this._escapeODataValue(listName);
        const fieldsData = await this._safeGetJson<any>(
            `${siteUrl}/_api/web/lists/getbytitle('${escapedListName}')/fields?$select=Title,InternalName,StaticName,Hidden,ReadOnlyField`,
            `${listName} fields`
        );
        const fields = this._toCollection(fieldsData) as ISharePointListFieldInfo[];
        const normalizedInternalName = this._normalizeFieldKey(internalName);
        const existingField = fields.find((field) => {
            const fieldNames = [field.InternalName, field.StaticName, field.Title]
                .map((value) => this._normalizeFieldKey((value || '').toString()))
                .filter((value) => !!value);
            return fieldNames.indexOf(normalizedInternalName) !== -1;
        });

        if (existingField) {
            return;
        }

        const digest = await this._getFormDigestValue();
        const schemaXml = `<Field Type="DateTime" Name="${this._escapeXmlAttribute(internalName)}" StaticName="${this._escapeXmlAttribute(internalName)}" DisplayName="${this._escapeXmlAttribute(displayName)}" Format="DateOnly" FriendlyDisplayFormat="Disabled" Group="LMS Generated Columns" />`;
        const response = await this._getHttpClient().post(
            `${siteUrl}/_api/web/lists/getbytitle('${escapedListName}')/fields/createfieldasxml`,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'X-RequestDigest': digest
                }),
                body: JSON.stringify({
                    parameters: {
                        SchemaXml: schemaXml,
                        Options: 0
                    }
                })
            }
        );

        if (!response.ok) {
            const errorText = await this._readErrorBody(response);
            if ((errorText || '').toLowerCase().indexOf('duplicate') !== -1 || (errorText || '').toLowerCase().indexOf('already exists') !== -1) {
                return;
            }

            throw new Error(`Failed to ensure ${internalName} on ${listName} (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 200) || 'No error details returned.'}`);
        }
    }

    private static async _ensureContentLibraryList(): Promise<string> {
        const listName = await this._resolveContentLibraryListName();
        await this._ensureList(listName);
        await Promise.all([
            this._ensureTextFieldOnList(listName, 'FileLink'),
            this._ensureTextFieldOnList(listName, 'UploadedBy'),
            this._ensureTextFieldOnList(listName, 'AssignedTo'),
            this._ensureTextFieldOnList(listName, 'Status'),
            this._ensureTextFieldOnList(listName, 'FolderName'),
            this._ensureTextFieldOnList(listName, 'AssetType'),
            this._ensureTextFieldOnList(listName, 'Description'),
            this._ensureTextFieldOnList(listName, 'FileSize')
        ]);
        this._contentLibraryListSchemaCache = null;
        this._contentLibraryListSchemaPromise = null;
        return listName;
    }

    private static async _ensureCertificationCompletionList(): Promise<string> {
        const listName = await this._resolveCertificationCompletionListName();
        await this._ensureList(listName);
        await Promise.all([
            this._ensureTextFieldOnList(listName, 'CertID'),
            this._ensureDateFieldOnList(listName, 'ExamDate'),
            this._ensureDateFieldOnList(listName, 'RenewalDate'),
            this._ensureTextFieldOnList(listName, 'ExamCode')
        ]);
        this._certificationCompletionListSchemaCache = null;
        this._certificationCompletionListSchemaPromise = null;
        return listName;
    }

    private static async _tryEnsureOptionalEnrollmentField(
        listName: string,
        fieldName: string,
        ensureField: () => Promise<void>
    ): Promise<void> {
        try {
            await ensureField();
        } catch (error) {
            const rawMessage = error instanceof Error ? error.message : `${error || ''}`;
            const normalizedMessage = rawMessage.toLowerCase();
            const logMessage = normalizedMessage.indexOf('total size of the columns in this list exceeds the limit') !== -1
                ? `[Enrollments] Skipping optional '${fieldName}' field creation because the list has reached the SharePoint column-size limit.`
                : `[Enrollments] Optional '${fieldName}' field is unavailable. Continuing with the existing enrollment schema.`;

            this._logResponseIssueOnce(
                `enrollment-optional-field:${listName}:${fieldName}`,
                logMessage,
                {
                    listName,
                    fieldName,
                    error
                }
            );
        }
    }

    private static async _ensureEnrollmentCompletionFields(listName: string): Promise<void> {
        await Promise.all([
            this._tryEnsureOptionalEnrollmentField(listName, 'CompletionDate', () => this._ensureDateFieldOnList(listName, 'CompletionDate')),
            this._tryEnsureOptionalEnrollmentField(listName, 'ExamCode', () => this._ensureTextFieldOnList(listName, 'ExamCode'))
        ]);
        this._enrollmentListSchemaCache = null;
        this._enrollmentListSchemaPromise = null;
    }

    private static _escapeODataValue(value: string): string {
        return (value || '').replace(/'/g, "''");
    }

    private static _toCollection(data: any): any[] {
        return data?.value || data?.d?.results || [];
    }

    private static _parseAssessmentPayload(value: string): IAssessmentAssignmentDefinition | null {
        if (!value) {
            return null;
        }

        try {
            return JSON.parse(value) as IAssessmentAssignmentDefinition;
        } catch (error) {
            console.warn('[AssessmentAssignments] Failed to parse assessment payload', error);
            return null;
        }
    }

    private static async _processInChunks<T>(
        items: T[],
        chunkSize: number,
        handler: (chunk: T[]) => Promise<void>,
        delayBetweenChunksMs: number = 0
    ): Promise<void> {
        for (let index = 0; index < items.length; index += chunkSize) {
            await handler(items.slice(index, index + chunkSize));

            if (delayBetweenChunksMs > 0 && index + chunkSize < items.length) {
                await this._delay(delayBetweenChunksMs);
            }
        }
    }

    private static _mapEnrollmentItem(item: any, schema?: IEnrollmentListSchema | null): IEnrollment {
        const certificationLookup = [
            (schema?.certificationLookupField ? this._readFieldValue(item, schema.certificationLookupField) : undefined),
            item?.CertificationId,
            item?.Certification
        ].find((value: any) => value && typeof value === 'object') || null;
        const assignedUser = [
            (schema?.assignedToField ? this._readFieldValue(item, schema.assignedToField) : undefined),
            item?.AssignedTo,
            item?.Assigned_x0020_To,
            item?.User,
            item?.Learner
        ].find((value: any) => value && typeof value === 'object') || null;
        const assignedByUser = [
            (schema?.assignedByField ? this._readFieldValue(item, schema.assignedByField) : undefined),
            item?.AssignedBy,
            item?.Assigned_x0020_By
        ].find((value: any) => value && typeof value === 'object') || null;
        const userEmail = (
            assignedUser?.EMail ||
            assignedUser?.Email ||
            (schema?.assignedToEmailField ? this._readFieldValue(item, schema.assignedToEmailField) : undefined) ||
            (schema?.userEmailField ? this._readFieldValue(item, schema.userEmailField) : undefined) ||
            item?.AssignedToEmail ||
            item?.UserEmail ||
            item?.LearnerEmail ||
            ''
        ).toString().trim();
        const userName = (
            assignedUser?.Title ||
            (schema?.userNameField ? this._readFieldValue(item, schema.userNameField) : undefined) ||
            item?.UserName ||
            item?.AssignedToName ||
            item?.LearnerName ||
            userEmail
        ).toString().trim();
        const certName = (
            certificationLookup?.Title ||
            (schema?.certNameField ? this._readFieldValue(item, schema.certNameField) : undefined) ||
            item?.CertName ||
            item?.CertificateName ||
            item?.Certification ||
            item?.Title ||
            ''
        ).toString().trim();
        const certCode = (
            certificationLookup?.Code ||
            (schema?.certCodeField ? this._readFieldValue(item, schema.certCodeField) : undefined) ||
            item?.CertCode ||
            item?.CertificationCode ||
            item?.Code ||
            certName ||
            `CERT-${item?.Id || 'UNKNOWN'}`
        ).toString().trim();
        const assignedDate = ((schema?.assignedDateField ? this._readFieldValue(item, schema.assignedDateField) : undefined) || item?.AssignedDate || item?.StartDate || item?.Created || '').toString();
        const storedEndDate = ((schema?.endDateField ? this._readFieldValue(item, schema.endDateField) : undefined) || item?.EndDate || '').toString();
        const examScheduledDate = ((schema?.examScheduledDateField ? this._readFieldValue(item, schema.examScheduledDateField) : undefined) || item?.ExamScheduledDate || storedEndDate || '').toString();
        const rescheduledDate = ((schema?.rescheduledDateField ? this._readFieldValue(item, schema.rescheduledDateField) : undefined) || item?.RescheduledDate || '').toString();
        const expiryDate = ((schema?.expiryDateField ? this._readFieldValue(item, schema.expiryDateField) : undefined) || item?.ExpiryDate || item?.ExpirationDate || storedEndDate || '').toString();
        const completionDate = ((schema?.completionDateField ? this._readFieldValue(item, schema.completionDateField) : undefined) || item?.CompletionDate || item?.CompletedOn || '').toString();
        const examCode = ((schema?.examCodeField ? this._readFieldValue(item, schema.examCodeField) : undefined) || item?.ExamCode || '').toString().trim();
        const assignedByName = (
            assignedByUser?.Title ||
            (schema?.assignedByNameField ? this._readFieldValue(item, schema.assignedByNameField) : undefined) ||
            item?.AssignedByName ||
            ''
        ).toString().trim();
        const displayExamDate = rescheduledDate || examScheduledDate;
        const pathId = (
            certificationLookup?.Id ||
            (schema?.pathIdField ? this._readFieldValue(item, schema.pathIdField) : undefined) ||
            item?.PathId ||
            item?.CertificationId ||
            item?.CertCode ||
            item?.CertificationCode ||
            item?.Code ||
            certCode
        ).toString().trim();

        return {
            id: item?.Id,
            userEmail,
            userName,
            certificationId: Number(certificationLookup?.Id || item?.CertificationIdId || item?.CertificationId || 0) || undefined,
            userId: Number(assignedUser?.Id || item?.UserId || item?.AssignedToId || 0) || undefined,
            certCode,
            certName,
            startDate: (item?.StartDate || assignedDate || '').toString(),
            endDate: (storedEndDate || displayExamDate || '').toString(),
            status: this._normalizeEnrollmentStatus(item?.Status),
            progress: typeof item?.Progress === 'number' ? item.Progress : Number(item?.Progress || 0),
            certificateName: (item?.CertificateName || certificationLookup?.Title || certName).toString(),
            assignedDate,
            assignedByName,
            examScheduledDate,
            rescheduledDate,
            expiryDate,
            completionDate,
            examCode,
            assignedByAdmin: !!assignedByName,
            listStatus: (item?.Status || '').toString(),
            pathId
        };
    }

    private static _getEnrollmentPathId(enrollment: Partial<IEnrollment> | any): string {
        return (
            enrollment?.pathId ||
            enrollment?.certCode ||
            enrollment?.certName ||
            enrollment?.certificateName ||
            ''
        ).toString().trim();
    }

    private static _dedupeEnrollments(enrollments: IEnrollment[]): IEnrollment[] {
        const byKey = new Map<string, IEnrollment>();

        enrollments.forEach((enrollment) => {
            const dedupeKey = [
                (enrollment.userEmail || '').toLowerCase(),
                (enrollment.certCode || enrollment.certName || '').toLowerCase()
            ].join('::');

            if (!dedupeKey.replace(/[:]/g, '')) {
                return;
            }

            const existing = byKey.get(dedupeKey);
            const existingTimestamp = new Date(existing?.assignedDate || existing?.startDate || 0).getTime();
            const incomingTimestamp = new Date(enrollment.assignedDate || enrollment.startDate || 0).getTime();

            if (!existing || incomingTimestamp >= existingTimestamp) {
                byKey.set(dedupeKey, enrollment);
            }
        });

        return Array.from(byKey.values());
    }

    private static _getAuditLogUser(item: any, schema: IAuditLogListSchema): any {
        const user = this._readFieldValue(item, schema.userField);
        return user && typeof user === 'object' ? user : null;
    }

    private static _getAuditLogUserId(item: any, schema: IAuditLogListSchema): number | undefined {
        const user = this._getAuditLogUser(item, schema);
        const userIdValue =
            user?.Id ||
            (schema.userField ? item?.[`${schema.userField}Id`] : undefined) ||
            item?.UserId ||
            item?.AssignedToId ||
            0;
        const userId = Number(userIdValue || 0);
        return userId > 0 ? userId : undefined;
    }

    private static _getAuditLogUserEmail(item: any, schema: IAuditLogListSchema): string {
        const user = this._getAuditLogUser(item, schema);
        return (
            user?.EMail ||
            user?.Email ||
            this._readFieldValue(item, schema.learnerEmailField) ||
            ''
        ).toString().trim();
    }

    private static _getAuditLogUserName(item: any, schema: IAuditLogListSchema): string {
        const user = this._getAuditLogUser(item, schema);
        return (
            user?.Title ||
            this._readFieldValue(item, schema.learnerNameField) ||
            this._getAuditLogUserEmail(item, schema)
        ).toString().trim();
    }

    private static _getAuditLogPathId(item: any, schema: IAuditLogListSchema): string {
        return (
            this._readFieldValue(item, schema.pathIdField) ||
            this._readFieldValue(item, schema.assignmentNameField) ||
            item?.PathId ||
            item?.AssignmentName ||
            ''
        ).toString().trim();
    }

    private static _mapAuditLogItem(item: any, schema: IAuditLogListSchema): IAuditLogRecord {
        return {
            id: Number(item?.Id || item?.id || 0),
            title: (item?.Title || '').toString().trim(),
            learnerEmail: this._getAuditLogUserEmail(item, schema),
            learnerName: this._getAuditLogUserName(item, schema),
            action: (this._readFieldValue(item, schema.actionField) || item?.Action || '').toString().trim(),
            assignmentName: (this._readFieldValue(item, schema.assignmentNameField) || item?.AssignmentName || '').toString().trim(),
            assignmentDate: (this._readFieldValue(item, schema.assignmentDateField) || item?.AssignmentDate || item?.Created || '').toString(),
            assignedById: Number(
                (schema.assignedByField ? item?.[`${schema.assignedByField}Id`] : undefined) ||
                item?.AssignedById ||
                0
            ) || undefined,
            status: (this._readFieldValue(item, schema.statusField) || item?.Status || '').toString().trim(),
            created: (item?.Created || '').toString(),
            pathId: this._getAuditLogPathId(item, schema),
            userId: this._getAuditLogUserId(item, schema),
            timestamp: (this._readFieldValue(item, schema.timestampField) || item?.Timestamp || item?.Created || '').toString()
        };
    }

    private static _normalizeEnrollmentStatus(status?: string): string {
        const normalized = (status || '').toString().trim().toLowerCase();

        if (normalized === 'completed') {
            return 'completed';
        }

        if (normalized === 'in progress' || normalized === 'in-progress') {
            return 'in-progress';
        }

        if (normalized === 'new' || normalized === 'viewed' || normalized === 'assigned' || normalized === 'scheduled' || normalized === 'rescheduled') {
            return 'scheduled';
        }

        return normalized || 'scheduled';
    }

    private static async _updateListItemStatus(listName: string, itemId: number, status: string): Promise<void> {
        const siteUrl = this._getSiteUrl();
        const digest = await this._getFormDigestValue();

        await this._spHttpClient.post(
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items(${itemId})`,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'MERGE',
                    'X-RequestDigest': digest
                }),
                body: JSON.stringify({
                    Status: status
                })
            }
        );
    }

    private static async _getSiteUserIdByEmail(email: string): Promise<number | null> {
        const trimmedEmail = (email || '').trim().toLowerCase();
        if (!trimmedEmail) {
            return null;
        }

        const siteUrl = this._getSiteUrl();

        try {
            const digest = await this._getFormDigestValue();
            const ensuredUser = await this._safePostJson<any>(
                `${siteUrl}/_api/web/ensureuser`,
                {
                    headers: this._getJsonHeaders({
                        'X-RequestDigest': digest
                    }),
                    body: JSON.stringify({
                        logonName: trimmedEmail
                    })
                },
                `ensure user ${trimmedEmail}`
            );
            const ensuredUserId = Number(ensuredUser?.Id || ensuredUser?.d?.Id || 0);
            if (ensuredUserId > 0) {
                return ensuredUserId;
            }
        } catch (error) {
            console.warn('[Assignments] Could not resolve site user by email', {
                email: trimmedEmail,
                error
            });
        }

        const fallbackEndpoint =
            `${siteUrl}/_api/web/siteusers?$select=Id,EMail&$filter=EMail eq '${this._escapeODataValue(trimmedEmail)}'`;

        try {
            const data = await this._safeGetJson<any>(fallbackEndpoint, `site user ${trimmedEmail}`);
            const users = this._toCollection(data);
            return users.length > 0 ? Number(users[0].Id || 0) : null;
        } catch (error) {
            console.warn('[Assignments] Could not resolve site user by email from siteusers fallback', {
                email: trimmedEmail,
                fallbackEndpoint,
                error
            });
            return null;
        }
    }

    private static async _getCurrentContextUserId(): Promise<number | null> {
        const contextUserId = Number(this._context?.pageContext?.legacyPageContext?.userId || 0);
        if (contextUserId > 0) {
            return contextUserId;
        }

        const contextEmail = this._context?.pageContext?.user?.email || '';
        return this._getSiteUserIdByEmail(contextEmail);
    }

    private static _getCurrentContextUserName(): string {
        return (
            this._context?.pageContext?.user?.displayName ||
            this._context?.pageContext?.legacyPageContext?.userDisplayName ||
            ''
        ).toString().trim();
    }

    private static async _syncUserAssignmentNotification(enrollment: IEnrollment, storedStatus: string): Promise<void> {
        const listName = await this._resolveUserNotificationListName();

        if (listName === 'Notifications') {
            await this._ensureList(listName);
            const siteUrl = this._getSiteUrl();
            const digest = await this._getFormDigestValue();

            await this._spHttpClient.post(
                `${siteUrl}/_api/web/lists/getbytitle('Notifications')/items`,
                SPHttpClient.configurations.v1,
                {
                    headers: this._getJsonHeaders({
                        'X-RequestDigest': digest
                    }),
                    body: JSON.stringify({
                        Title: 'New Certification Assigned',
                        Description: enrollment.certName,
                        UserEmail: enrollment.userEmail,
                        Status: storedStatus.toLowerCase() === 'viewed' ? 'Viewed' : 'Unread',
                        AssignedDate: enrollment.assignedDate || new Date().toISOString()
                    })
                }
            );

            return;
        }

        await this.addNotification({
            title: enrollment.certName,
            text: 'New certification assigned by Admin',
            targetEmail: enrollment.userEmail,
            type: 'assignment',
            time: enrollment.assignedDate || new Date().toISOString(),
            read: false
        });
    }

    private static async _getEnrollmentEditRoleDefinitionId(): Promise<number | null> {
        if (this._enrollmentEditRoleDefinitionId !== undefined) {
            return this._enrollmentEditRoleDefinitionId;
        }

        if (this._enrollmentEditRoleDefinitionPromise) {
            return this._enrollmentEditRoleDefinitionPromise;
        }

        const siteUrl = this._getSiteUrl();
        const endpoint = `${siteUrl}/_api/web/roledefinitions?$select=Id,Name,RoleTypeKind,Hidden&$orderby=Order asc`;

        this._enrollmentEditRoleDefinitionPromise = (async () => {
            try {
                const roleDefinitionData = await this._safeGetJson<any>(endpoint, 'enrollment edit role definitions');
                const roleDefinitions = this._toCollection(roleDefinitionData)
                    .map((item: any) => ({
                        id: Number(item?.Id || item?.id || 0),
                        name: (item?.Name || '').toString().trim(),
                        roleTypeKind: Number(item?.RoleTypeKind || 0),
                        hidden: item?.Hidden === true
                    }))
                    .filter((item) => item.id > 0 && !item.hidden);

                const preferredRole =
                    roleDefinitions.find((item) => item.roleTypeKind === 6) ||
                    roleDefinitions.find((item) => item.roleTypeKind === 3) ||
                    roleDefinitions.find((item) => ['edit', 'contribute', 'design'].indexOf(item.name.toLowerCase()) !== -1) ||
                    null;

                this._enrollmentEditRoleDefinitionId = preferredRole?.id || null;
                return this._enrollmentEditRoleDefinitionId;
            } catch (error) {
                this._logResponseIssueOnce(
                    'enrollment-edit-role-definition-missing',
                    '[Enrollments] Could not resolve an editable role definition for learner access. Enrollment save will continue without permission sync.',
                    {
                        endpoint,
                        error
                    }
                );
                this._enrollmentEditRoleDefinitionId = null;
                return null;
            } finally {
                this._enrollmentEditRoleDefinitionPromise = null;
            }
        })();

        return this._enrollmentEditRoleDefinitionPromise;
    }

    private static async _ensureEnrollmentLearnerEditAccess(itemId: number, assignedToId: number): Promise<void> {
        const normalizedItemId = Number(itemId || 0);
        const normalizedAssignedToId = Number(assignedToId || 0);

        if (normalizedItemId <= 0 || normalizedAssignedToId <= 0) {
            return;
        }

        const context = await this._getEnrollmentListContext();
        const roleDefinitionId = await this._getEnrollmentEditRoleDefinitionId();

        if (!roleDefinitionId) {
            return;
        }

        const detailEndpoint =
            `${context.siteUrl}/_api/web/lists/getbytitle('${context.escapedListName}')/items(${normalizedItemId})` +
            `?$select=Id,HasUniqueRoleAssignments`;
        const detailData = await this._safeGetJson<any>(detailEndpoint, `enrollment permissions ${normalizedItemId}`);
        const hasUniqueRoleAssignments = !!(detailData?.HasUniqueRoleAssignments || detailData?.d?.HasUniqueRoleAssignments);
        const digest = await this._getFormDigestValue();

        if (!hasUniqueRoleAssignments) {
            const breakInheritanceResponse = await this._getHttpClient().post(
                `${context.siteUrl}/_api/web/lists/getbytitle('${context.escapedListName}')/items(${normalizedItemId})/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`,
                SPHttpClient.configurations.v1,
                {
                    headers: this._getJsonHeaders({
                        'X-RequestDigest': digest
                    })
                }
            );

            if (!breakInheritanceResponse.ok) {
                const errorText = await this._readErrorBody(breakInheritanceResponse);
                throw new Error(`Failed to enable unique permissions on enrollment ${normalizedItemId} (HTTP ${breakInheritanceResponse.status} ${breakInheritanceResponse.statusText}): ${errorText.substring(0, 300) || 'No error details returned.'}`);
            }
        }

        const assignResponse = await this._getHttpClient().post(
            `${context.siteUrl}/_api/web/lists/getbytitle('${context.escapedListName}')/items(${normalizedItemId})/roleassignments/addroleassignment(principalid=${normalizedAssignedToId},roledefid=${roleDefinitionId})`,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'X-RequestDigest': digest
                })
            }
        );

        if (!assignResponse.ok) {
            const errorText = await this._readErrorBody(assignResponse);
            const normalizedErrorText = errorText.toLowerCase();
            if (
                assignResponse.status === 409 ||
                normalizedErrorText.indexOf('already exists') !== -1 ||
                normalizedErrorText.indexOf('role assignment already exists') !== -1
            ) {
                return;
            }

            throw new Error(`Failed to grant learner edit access on enrollment ${normalizedItemId} (HTTP ${assignResponse.status} ${assignResponse.statusText}): ${errorText.substring(0, 300) || 'No error details returned.'}`);
        }
    }

    public static async ensureLearnerEditAccessForEnrollments(enrollments: Array<Partial<IEnrollment>>): Promise<void> {
        const siteKey = (this._getSiteUrl() || '').toString().trim().toLowerCase();
        const candidateTargets = Array.from(
            new Map(
                (enrollments || [])
                    .map((item) => {
                        const itemId = Number(item?.id || 0);
                        const userId = Number(item?.userId || 0);
                        const userName = (item?.userName || '').toString().trim().toLowerCase();
                        const userEmail = (
                            (item as any)?.userEmail ||
                            (item as any)?.email ||
                            ''
                        ).toString().trim().toLowerCase();
                        const learnerEmailAlias = userEmail.indexOf('@') > -1 ? userEmail.split('@')[0] : userEmail;
                        const assignedByName = (item?.assignedByName || '').toString().trim().toLowerCase();
                        const isLikelyAdminAssigned =
                            !!assignedByName &&
                            assignedByName !== userName &&
                            assignedByName !== userEmail &&
                            assignedByName !== learnerEmailAlias;

                        if (itemId <= 0 || (!userId && !userEmail) || !isLikelyAdminAssigned) {
                            return null;
                        }

                        const syncKey = `${siteKey}::${itemId}::${userEmail || userId}`;
                        return [
                            syncKey,
                            {
                                id: itemId,
                                userId,
                                userEmail,
                                syncKey
                            }
                        ] as const;
                    })
                    .filter((entry): entry is readonly [string, { id: number; userId: number; userEmail: string; syncKey: string }] => !!entry)
            ).values()
        ).filter((entry) => !this._enrollmentLearnerEditAccessSyncKeys.has(entry.syncKey));

        if (candidateTargets.length === 0) {
            return;
        }

        await this._processInChunks(
            candidateTargets,
            10,
            async (chunk) => {
                await Promise.all(
                    chunk.map(async (entry) => {
                        let resolvedUserId = Number(entry.userId || 0);

                        if (resolvedUserId <= 0 && entry.userEmail) {
                            const resolvedSiteUserId = await this._getSiteUserIdByEmail(entry.userEmail).catch((error) => {
                                console.warn('[Enrollments] Failed to resolve learner id while syncing edit access.', {
                                    enrollmentId: entry.id,
                                    userEmail: entry.userEmail,
                                    error
                                });
                                return 0;
                            });

                            resolvedUserId = Number(resolvedSiteUserId || 0);
                        }

                        if (resolvedUserId <= 0) {
                            return;
                        }

                        await this._ensureEnrollmentLearnerEditAccess(entry.id, resolvedUserId);
                        this._enrollmentLearnerEditAccessSyncKeys.add(entry.syncKey);
                    })
                );
            },
            100
        );
    }

    public static async getLearningAndSkillsItems(forceRefresh: boolean = false): Promise<Array<{ id: number; name: string; title: string; }>> {
        try {
            const siteUrl = this._ensureProductionSiteUrl();
            const listName = await this._resolveLearningAndSkillsListName(forceRefresh);
            const endpoint =
                `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items` +
                `?$select=Id,Title&$top=5000`;
            const data = await this._safeGetJson<any>(endpoint, `${listName} items`);
            return this._toCollection(data).map((item: any) => ({
                id: Number(item?.Id || item?.id || 0),
                title: (item?.Title || '').toString().trim(),
                name: (item?.Name || item?.Title || '').toString().trim()
            }));
        } catch (error) {
            this._logResponseIssueOnce('learning-and-skills-list-missing', '[LearningAndSkills] List lookup failed. Continuing without Learning and Skills items.', {
                requestedCandidates: this._learningAndSkillsListCandidates,
                error
            });
            return [];
        }
    }

    private static async _readJson<T>(response: any): Promise<T | null> {
        if (!response) {
            console.error('No response object provided to _readJson');
            return null;
        }

        const requestUrl = response.url || 'unknown SharePoint API URL';
        const contentType = response.headers?.get?.('content-type') || '';
        const responseText = await response.text().catch(() => '');

        if (!response.ok) {
            const errorMsg = `SharePoint API Error ${response.status} for ${requestUrl}: ${responseText.substring(0, 200) || 'No error details'}`;
            console.error(errorMsg, {
                requestUrl,
                status: response.status,
                statusText: response.statusText,
                responsePreview: responseText.substring(0, 500)
            });
            throw new Error(errorMsg);
        }

        const trimmedResponse = responseText.trim();
        if (!trimmedResponse) {
            return null;
        }

        const lowerResponse = trimmedResponse.toLowerCase();
        const looksLikeHtml =
            lowerResponse.indexOf('<!doctype html') === 0 ||
            lowerResponse.indexOf('<html') === 0 ||
            lowerResponse.indexOf('<head') === 0 ||
            lowerResponse.indexOf('<body') === 0;

        if (looksLikeHtml || (contentType && contentType.toLowerCase().indexOf('json') === -1 && trimmedResponse.indexOf('<') === 0)) {
            const errorMsg =
                `SharePoint returned non-JSON content for ${requestUrl}: ${trimmedResponse.substring(0, 120)}. ` +
                `This usually means an authentication redirect or an incorrect _api endpoint.`;
            console.error(errorMsg, {
                requestUrl,
                contentType,
                responsePreview: trimmedResponse.substring(0, 500)
            });
            throw new Error(errorMsg);
        }

        try {
            return JSON.parse(responseText) as T;
        } catch (error) {
            console.error('[SharePoint API] Failed to parse JSON response', {
                requestUrl,
                contentType,
                responsePreview: trimmedResponse.substring(0, 500),
                error
            });
            throw new Error(`SharePoint returned invalid JSON for ${requestUrl}.`);
        }
    }

    private static _formatDisplayDate(value?: string): string {
        if (!value) {
            return new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
        }

        const date = new Date(value);
        if (Number.isNaN(date.getTime())) {
            return value;
        }

        return date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    }

    private static _formatFileSize(length: number | string): string {
        const numericLength = typeof length === 'string' ? parseInt(length, 10) : length;

        if (!numericLength || Number.isNaN(numericLength)) {
            return 'N/A';
        }

        if (numericLength < 1024) {
            return `${numericLength} B`;
        }

        if (numericLength < 1024 * 1024) {
            return `${(numericLength / 1024).toFixed(1)} KB`;
        }

        if (numericLength < 1024 * 1024 * 1024) {
            return `${(numericLength / (1024 * 1024)).toFixed(1)} MB`;
        }

        return `${(numericLength / (1024 * 1024 * 1024)).toFixed(1)} GB`;
    }

    private static _guessAssetType(fileName: string): string {
        const lowerName = (fileName || '').toLowerCase();

        if (lowerName.endsWith('.mp4') || lowerName.endsWith('.mov') || lowerName.endsWith('.avi') || lowerName.endsWith('.mkv')) {
            return 'VIDEO';
        }

        if (lowerName.endsWith('.pdf')) {
            return 'PDF';
        }

        if (lowerName.endsWith('.xlsx') || lowerName.endsWith('.xls') || lowerName.endsWith('.csv')) {
            return 'EXCEL';
        }

        if (lowerName.endsWith('.doc') || lowerName.endsWith('.docx') || lowerName.endsWith('.txt') || lowerName.endsWith('.md')) {
            return 'DOC';
        }

        if (lowerName.endsWith('.ppt') || lowerName.endsWith('.pptx')) {
            return 'PPT';
        }

        if (lowerName.endsWith('.zip') || lowerName.endsWith('.scorm')) {
            return 'SCORM';
        }

        if (lowerName.startsWith('http://') || lowerName.startsWith('https://')) {
            return 'LINK';
        }

        return 'DOC';
    }

    private static _buildAbsoluteUrl(serverRelativeUrl: string): string {
        if (!serverRelativeUrl) {
            return '';
        }

        if (serverRelativeUrl.startsWith('http://') || serverRelativeUrl.startsWith('https://')) {
            return serverRelativeUrl;
        }

        try {
            return `${new URL(this._siteUrl).origin}${serverRelativeUrl}`;
        } catch (error) {
            return serverRelativeUrl;
        }
    }

    private static _decodeUriComponentSafely(value: string): string {
        try {
            return decodeURIComponent(value);
        } catch (error) {
            return value;
        }
    }

    private static _normalizeContentAssetLookupKey(candidate?: string): string {
        const normalizedCandidate = (candidate || '').toString().trim();
        if (!normalizedCandidate) {
            return '';
        }

        const serverRelativeUrl = this._resolveServerRelativeUrlFromString(normalizedCandidate);
        if (serverRelativeUrl) {
            return this._decodeUriComponentSafely(serverRelativeUrl)
                .split('#')[0]
                .split('?')[0]
                .replace(/\\/g, '/')
                .toLowerCase();
        }

        try {
            const resolvedUrl = new URL(normalizedCandidate, this._siteUrl);
            return this._decodeUriComponentSafely(`${resolvedUrl.origin}${resolvedUrl.pathname}`)
                .replace(/\\/g, '/')
                .toLowerCase();
        } catch (error) {
            return this._decodeUriComponentSafely(normalizedCandidate)
                .split('#')[0]
                .split('?')[0]
                .replace(/\\/g, '/')
                .toLowerCase();
        }
    }

    private static _isMissingSharePointResource(status: number, errorText: string): boolean {
        const normalizedError = (errorText || '').toString().trim().toLowerCase();
        return status === 404 ||
            normalizedError.indexOf('does not exist') >= 0 ||
            normalizedError.indexOf('not found') >= 0 ||
            normalizedError.indexOf('cannot find resource') >= 0;
    }

    private static async _recycleContentAssetFile(serverRelativeUrl: string, digest: string): Promise<void> {
        const normalizedServerRelativeUrl = this._decodeUriComponentSafely((serverRelativeUrl || '').toString().trim());
        if (!normalizedServerRelativeUrl) {
            return;
        }

        const siteUrl = normalizedServerRelativeUrl.toLowerCase().indexOf(this._documents1SiteServerRelativePath.toLowerCase()) === 0
            ? this._getDocuments1SiteUrl()
            : this._getSiteUrl();
        const escapedAssetUrl = this._escapeODataValue(normalizedServerRelativeUrl);
        const endpoints = [
            `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapedAssetUrl}')/recycle()`,
            `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${escapedAssetUrl}')/recycle()`
        ];
        let lastError: Error | null = null;

        for (const endpoint of endpoints) {
            try {
                const response = await this._getHttpClient().post(
                    endpoint,
                    SPHttpClient.configurations.v1,
                    {
                        headers: this._getJsonHeaders({
                            'X-RequestDigest': digest
                        })
                    }
                );

                if (response.ok) {
                    return;
                }

                const errorText = await this._readErrorBody(response);
                if (this._isMissingSharePointResource(response.status, errorText)) {
                    return;
                }

                lastError = new Error(`Failed to recycle file from Documents1 (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 300) || 'No error details returned.'}`);
            } catch (error) {
                lastError = error instanceof Error ? error : new Error(String(error));
            }
        }

        if (lastError) {
            throw lastError;
        }
    }

    private static _getContentAssetMapKey(asset: Partial<IContentAsset> | any): string {
        const normalizedPathKey = this._normalizeContentAssetLookupKey(asset?.path || asset?.url || '');
        if (normalizedPathKey) {
            return normalizedPathKey;
        }

        const fallbackKey = asset?.name || asset?.id || '';
        return fallbackKey.toString().trim().toLowerCase();
    }

    private static _normalizeContentAssetName(value?: string): string {
        return this._decodeUriComponentSafely((value || '').toString().trim())
            .replace(/\s+/g, ' ')
            .toLowerCase();
    }

    private static _normalizeContentAssetFolder(value?: string): string {
        return this._decodeUriComponentSafely((value || '').toString().trim())
            .replace(/\\/g, '/')
            .replace(/^\/+|\/+$/g, '')
            .replace(/\s+/g, ' ')
            .toLowerCase();
    }

    private static _getContentAssetAliasKeys(asset: Partial<IContentAsset> | any): string[] {
        const aliasKeys = new Set<string>();
        const primaryKey = this._getContentAssetMapKey(asset);
        if (primaryKey) {
            aliasKeys.add(primaryKey);
        }

        const normalizedName = this._normalizeContentAssetName(asset?.name);
        const normalizedFolder = this._normalizeContentAssetFolder(
            asset?.folderName ||
            this._extractFolderNameFromServerRelativeUrl(
                this._resolveServerRelativeUrlFromString(asset?.path || asset?.url || '')
            )
        );

        if (normalizedName && normalizedFolder) {
            aliasKeys.add(`folder:${normalizedFolder}|name:${normalizedName}`);
        } else if (normalizedName) {
            aliasKeys.add(`name:${normalizedName}`);
        }

        return Array.from(aliasKeys);
    }

    private static _getContentLibraryItemLookupKey(item: any, schema: IContentLibraryListSchema): string {
        const fileLink = this._buildNormalizedContentFileUrl(
            (schema.fileLinkField ? this._readFieldValue(item, schema.fileLinkField) : undefined) || item?.FileLink || ''
        );
        const normalizedFileKey = this._normalizeContentAssetLookupKey(fileLink);
        if (normalizedFileKey) {
            return normalizedFileKey;
        }

        const normalizedTitle = this._normalizeContentAssetName(item?.Title || '');
        const normalizedFolder = this._normalizeContentAssetFolder(
            (schema.folderNameField ? this._readFieldValue(item, schema.folderNameField) : undefined) || item?.FolderName || ''
        );

        if (normalizedTitle && normalizedFolder) {
            return `folder:${normalizedFolder}|name:${normalizedTitle}`;
        }

        if (normalizedTitle) {
            return `name:${normalizedTitle}`;
        }

        return '';
    }

    private static _dedupeContentLibraryItems(items: any[], schema: IContentLibraryListSchema): any[] {
        const dedupedItems = new Map<string, any>();
        const passthroughItems: any[] = [];

        items.forEach((item: any) => {
            const lookupKey = this._getContentLibraryItemLookupKey(item, schema);
            if (!lookupKey) {
                passthroughItems.push(item);
                return;
            }

            if (!dedupedItems.has(lookupKey)) {
                dedupedItems.set(lookupKey, item);
            }
        });

        return [...Array.from(dedupedItems.values()), ...passthroughItems];
    }

    private static _mapContentLibraryItem(item: any, schema: IContentLibraryListSchema): IContentAsset {
        const fileLink = this._buildNormalizedContentFileUrl(
            (schema.fileLinkField ? this._readFieldValue(item, schema.fileLinkField) : undefined) || item?.FileLink || ''
        );
        const serverRelativeUrl = this._resolveServerRelativeUrlFromString(fileLink);
        const folderName = (
            (schema.folderNameField ? this._readFieldValue(item, schema.folderNameField) : undefined) ||
            this._extractFolderNameFromServerRelativeUrl(serverRelativeUrl)
        ).toString().trim();
        const uploadedBy = ((schema.uploadedByField ? this._readFieldValue(item, schema.uploadedByField) : undefined) || item?.UploadedBy || '').toString().trim().toLowerCase();
        const assignedTo = ((schema.assignedToField ? this._readFieldValue(item, schema.assignedToField) : undefined) || item?.AssignedTo || '').toString().trim().toLowerCase();
        const assetType = ((schema.assetTypeField ? this._readFieldValue(item, schema.assetTypeField) : undefined) || item?.AssetType || '').toString().trim();
        const size = ((schema.fileSizeField ? this._readFieldValue(item, schema.fileSizeField) : undefined) || item?.FileSize || '').toString().trim();
        const description = ((schema.descriptionField ? this._readFieldValue(item, schema.descriptionField) : undefined) || item?.Description || '').toString().trim();
        const status = ((schema.statusField ? this._readFieldValue(item, schema.statusField) : undefined) || item?.Status || 'Uploaded').toString().trim();

        return {
            id: Number(item?.Id || item?.id || 0),
            name: (item?.Title || '').toString().trim(),
            type: assetType || this._guessAssetType(fileLink || item?.Title || ''),
            owner: uploadedBy || 'SharePoint',
            status: status || 'Uploaded',
            dateAdded: this._formatDisplayDate(item?.Modified || item?.Created || ''),
            size: size || '',
            description: description || (folderName ? `Stored in Documents1/${folderName}.` : 'Stored in Documents1.'),
            url: fileLink,
            path: serverRelativeUrl,
            folderName,
            uploadedBy,
            assignedTo
        };
    }

    private static _mergeContentAssetCollections(metadataAssets: IContentAsset[], libraryFiles: IContentAsset[], includeUnmappedFiles: boolean): IContentAsset[] {
        const mergedAssets = new Map<string, IContentAsset>();
        const assetAliases = new Map<string, string>();

        const registerAsset = (asset: IContentAsset, preferredKey?: string): string => {
            const aliasKeys = this._getContentAssetAliasKeys(asset);
            const primaryKey = preferredKey || aliasKeys[0] || `${asset?.id || asset?.name || Date.now()}`.toString().trim().toLowerCase();
            mergedAssets.set(primaryKey, asset);
            aliasKeys.forEach((aliasKey) => {
                if (aliasKey) {
                    assetAliases.set(aliasKey, primaryKey);
                }
            });
            return primaryKey;
        };

        const resolveExistingKey = (asset: IContentAsset): string => {
            const aliasKeys = this._getContentAssetAliasKeys(asset);
            for (const aliasKey of aliasKeys) {
                const existingKey = assetAliases.get(aliasKey);
                if (existingKey) {
                    return existingKey;
                }
            }

            return '';
        };

        metadataAssets.forEach((asset) => {
            const existingKey = resolveExistingKey(asset);
            if (existingKey) {
                registerAsset({
                    ...(mergedAssets.get(existingKey) || {}),
                    ...asset
                }, existingKey);
                return;
            }

            registerAsset(asset);
        });

        libraryFiles.forEach((file) => {
            const existingKey = resolveExistingKey(file);
            if (!existingKey && !includeUnmappedFiles) {
                return;
            }

            const registrationKey = existingKey || this._getContentAssetAliasKeys(file)[0] || this._getContentAssetMapKey(file);
            if (!registrationKey) {
                return;
            }

            const existing = existingKey ? mergedAssets.get(existingKey) : undefined;

            registerAsset({
                id: existing?.id || file.id,
                name: existing?.name || file.name,
                type: existing?.type || file.type,
                owner: existing?.owner || file.owner,
                status: existing?.status || file.status,
                dateAdded: existing?.dateAdded || file.dateAdded,
                size: existing?.size || file.size,
                description: existing?.description || file.description || 'Stored in Documents1.',
                url: file.url || existing?.url,
                path: file.path || existing?.path,
                folderName: existing?.folderName || file.folderName || this._extractFolderNameFromServerRelativeUrl(file.path || file.url || ''),
                uploadedBy: existing?.uploadedBy,
                assignedTo: existing?.assignedTo
            }, registrationKey);
        });

        return Array.from(mergedAssets.values()).sort((a, b) => (a.name || '').localeCompare(b.name || ''));
    }

    private static _getDocuments1SiteUrl(): string {
        try {
            const currentSiteUrl = new URL(this._getSiteUrl());
            return `${currentSiteUrl.origin}${this._documents1SiteServerRelativePath}`;
        } catch (error) {
            return this._getSiteUrl();
        }
    }

    private static _normalizeContentFolderName(folderName?: string): string {
        const normalizedInput = (folderName || '').toString().trim();
        const normalizedKey = normalizedInput.toLowerCase();
        const departmentFolderMap: Record<string, string> = {
            engineering: 'Engineering',
            finance: 'Finance',
            hr: 'HR',
            presales: 'Presales',
            sales: 'Sales'
        };

        if (departmentFolderMap[normalizedKey]) {
            return departmentFolderMap[normalizedKey];
        }

        return normalizedInput
            .replace(/[~"#%&*:<>?/\\{|}]/g, '-')
            .replace(/\s+/g, '-')
            .replace(/-+/g, '-')
            .replace(/^-|-$/g, '');
    }

    private static _getContentLibraryRootServerRelativeUrl(): string {
        return this._documents1LibraryServerRelativePath;
    }

    private static _buildDocuments1FolderServerRelativeUrl(folderName?: string): string {
        const normalizedFolderName = this._normalizeContentFolderName(folderName);
        const rootPath = this._getContentLibraryRootServerRelativeUrl();

        if (!normalizedFolderName) {
            return rootPath;
        }

        return `${rootPath}/${normalizedFolderName}`.replace(/\/{2,}/g, '/');
    }

    private static _extractFolderNameFromServerRelativeUrl(serverRelativeUrl?: string): string {
        const normalizedPath = (serverRelativeUrl || '').toString().trim();
        if (!normalizedPath) {
            return '';
        }

        const documentsRoot = this._getContentLibraryRootServerRelativeUrl().toLowerCase();
        const normalizedLowerPath = normalizedPath.toLowerCase();
        if (normalizedLowerPath.indexOf(documentsRoot) !== 0) {
            return '';
        }

        const relativePath = normalizedPath.substring(documentsRoot.length).replace(/^\/+/, '');
        const segments = relativePath.split('/').filter((segment) => !!segment);
        return segments.length > 1 ? segments[0] : '';
    }

    private static async _createFolderIfNotExists(folderName?: string): Promise<string> {
        const normalizedFolderName = this._normalizeContentFolderName(folderName);
        const targetFolderPath = this._buildDocuments1FolderServerRelativeUrl(normalizedFolderName);

        if (!normalizedFolderName) {
            return targetFolderPath;
        }

        const siteUrl = this._getDocuments1SiteUrl();
        const escapedTargetFolderPath = this._escapeODataValue(targetFolderPath);

        try {
            const response = await this._getHttpClient().get(
                `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapedTargetFolderPath}')/Exists`,
                SPHttpClient.configurations.v1,
                {
                    headers: this._getJsonHeaders()
                }
            );

            if (response.ok) {
                const existsData = await this._readJson<any>(response).catch(() => null);
                const existsValue =
                    existsData?.value ??
                    existsData?.Exists ??
                    existsData?.d?.Exists ??
                    existsData?.d?.GetFolderByServerRelativeUrl?.Exists;

                if (existsValue === true) {
                    return targetFolderPath;
                }
            }
        } catch (error) {
            // Folder existence fallback below
        }

        try {
            const response = await this._getHttpClient().get(
                `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapedTargetFolderPath}')?$select=ServerRelativeUrl,Name`,
                SPHttpClient.configurations.v1,
                {
                    headers: this._getJsonHeaders()
                }
            );

            if (response.ok) {
                return targetFolderPath;
            }
        } catch (error) {
            // Folder create fallback below
        }

        const digest = await this._getFormDigestValue();
        const createEndpoint = `${siteUrl}/_api/web/folders`;
        const createResponse = await this._getHttpClient().post(
            createEndpoint,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'X-RequestDigest': digest
                }),
                body: JSON.stringify({
                    __metadata: { type: 'SP.Folder' },
                    ServerRelativeUrl: targetFolderPath
                })
            }
        );

        if (!createResponse.ok) {
            const errorText = await this._readErrorBody(createResponse);
            const normalizedError = errorText.toLowerCase();
            if (normalizedError.indexOf('already exists') === -1 && normalizedError.indexOf('a file or folder with the name') === -1) {
                throw new Error(`Failed to create Documents1/${normalizedFolderName} folder (HTTP ${createResponse.status} ${createResponse.statusText}): ${errorText.substring(0, 250) || 'No error details returned.'}`);
            }
        }

        const verifyResponse = await this._getHttpClient().get(
            `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapedTargetFolderPath}')?$select=ServerRelativeUrl,Name`,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders()
            }
        );

        if (!verifyResponse.ok) {
            const errorText = await this._readErrorBody(verifyResponse);
            throw new Error(`Folder verification failed for Documents1/${normalizedFolderName} (HTTP ${verifyResponse.status} ${verifyResponse.statusText}): ${errorText.substring(0, 250) || 'No error details returned.'}`);
        }

        return targetFolderPath;
    }

    private static _resolveServerRelativeAssetUrl(asset: IContentAsset): string {
        return this._resolveServerRelativeUrlFromString(asset?.url || asset?.path || '');
    }

    private static _resolveServerRelativeUrlFromString(candidate?: string): string {
        const normalizedCandidate = (candidate || '').toString().trim();

        if (!normalizedCandidate) {
            return '';
        }

        if (normalizedCandidate.startsWith('/')) {
            return normalizedCandidate;
        }

        if (normalizedCandidate.startsWith('http://') || normalizedCandidate.startsWith('https://')) {
            try {
                const assetUrl = new URL(normalizedCandidate);
                const siteOrigin = new URL(this._siteUrl).origin;
                return assetUrl.origin === siteOrigin ? assetUrl.pathname : '';
            } catch (error) {
                return '';
            }
        }

        return '';
    }

    private static _buildNormalizedContentFileUrl(candidate?: string): string {
        const normalizedCandidate = (candidate || '').toString().trim();
        if (!normalizedCandidate) {
            return '';
        }

        const serverRelativeUrl = this._resolveServerRelativeUrlFromString(normalizedCandidate);
        if (serverRelativeUrl) {
            return this._buildAbsoluteUrl(serverRelativeUrl);
        }

        return normalizedCandidate;
    }

    // @ts-ignore - retained as a fallback helper for future non-Documents1 scenarios
    private static async _getDocumentLibrary(): Promise<any | null> {
        if (this._documentLibraryPromise) {
            return this._documentLibraryPromise;
        }

        this._documentLibraryPromise = (async () => {
            const listsData = await this._safeGetJson<any>(
                `${this._getSiteUrl()}/_api/web/lists?$select=Title,BaseTemplate,RootFolder/ServerRelativeUrl&$expand=RootFolder`,
                'document libraries'
            );
            const allLists = this._toCollection(listsData);

            const preferredLib = allLists.find((list: any) =>
                list?.BaseTemplate === 101 && (list?.Title === 'Documents' || list?.Title === 'Shared Documents')
            );

            return preferredLib ||
                allLists.find((list: any) => list?.BaseTemplate === 101 && list?.RootFolder?.ServerRelativeUrl) ||
                null;
        })().catch((error) => {
            this._documentLibraryPromise = null;
            throw error;
        });

        return this._documentLibraryPromise;
    }

    private static async _getFormDigestValue(): Promise<string> {
        const siteUrl = this._getSiteUrl();
        const now = Date.now();

        if (this._formDigestCache?.siteUrl === siteUrl && this._formDigestCache.expiresAt > now) {
            return this._formDigestCache.value;
        }

        if (this._formDigestPromise) {
            return this._formDigestPromise;
        }

        this._formDigestPromise = (async (): Promise<string> => {
            try {
                const data = await this._safePostJson<any>(
                    `${siteUrl}/_api/contextinfo`,
                    {
                        headers: this._getJsonHeaders(),
                        body: ''
                    },
                    'SharePoint form digest'
                );

                const formDigest =
                    data?.FormDigestValue ||
                    data?.GetContextWebInformation?.FormDigestValue ||
                    data?.d?.GetContextWebInformation?.FormDigestValue;
                if (!formDigest) {
                    throw new Error(`Form digest value not found in SharePoint contextinfo response. Response: ${JSON.stringify(data).substring(0, 300)}`);
                }

                const digestTimeoutSeconds = Number(
                    data?.FormDigestTimeoutSeconds ||
                    data?.GetContextWebInformation?.FormDigestTimeoutSeconds ||
                    data?.d?.GetContextWebInformation?.FormDigestTimeoutSeconds ||
                    1800
                );

                this._formDigestCache = {
                    siteUrl,
                    value: formDigest,
                    expiresAt: Date.now() + Math.max((digestTimeoutSeconds * 1000) - 30000, 60000)
                };

                return formDigest;
            } finally {
                this._formDigestPromise = null;
            }
        })();

        try {
            return await this._formDigestPromise;
        } catch (error) {
            console.error('Error fetching form digest:', error);
            throw error;
        }
    }

    public static async uploadFile(file: File, folderName?: string): Promise<{ name: string; url: string; serverRelativeUrl: string; folderName?: string }> {
        if (!file) {
            throw new Error('No file provided for upload.');
        }

        if (!this._getHttpClient()) {
            throw new Error('SharePointService not initialized. Call init() first with context and spHttpClient.');
        }

        try {
            const siteUrl = this._getDocuments1SiteUrl();
            const targetFolderPath = await this._createFolderIfNotExists(folderName);
            const escapedTargetFolderPath = this._escapeODataValue(targetFolderPath);
            const encodedFileName = encodeURIComponent(file.name).replace(/'/g, '%27');
            const uploadUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapedTargetFolderPath}')/Files/add(url='${encodedFileName}',overwrite=true)`;

            console.log('Uploading to:', targetFolderPath);

            console.log('[Upload-Docs1] Uploading file to SharePoint', {
                fileName: file.name,
                folderName: this._normalizeContentFolderName(folderName) || 'root',
                siteUrl,
                folderPath: targetFolderPath,
                uploadUrl
            });

            let response = await this._getHttpClient().post(
                uploadUrl,
                SPHttpClient.configurations.v1,
                {
                    headers: this._getJsonHeaders({
                        'Content-Type': 'application/octet-stream'
                    }),
                    body: file
                }
            );

            if (!response.ok && response.status === 404) {
                await new Promise((resolve) => setTimeout(resolve, 600));
                await this._createFolderIfNotExists(folderName);
                response = await this._getHttpClient().post(
                    uploadUrl,
                    SPHttpClient.configurations.v1,
                    {
                        headers: this._getJsonHeaders({
                            'Content-Type': 'application/octet-stream'
                        }),
                        body: file
                    }
                );
            }

            if (!response.ok) {
                const errorText = await this._readErrorBody(response);
                const isHtmlError = errorText.trim().toLowerCase().indexOf('<html') !== -1;
                console.error('[Upload-Docs1] SharePoint upload failed', {
                    status: response.status,
                    statusText: response.statusText,
                    responsePreview: errorText.substring(0, 500)
                });
                if (isHtmlError) {
                    throw new Error(`SharePoint returned HTML for the upload request (HTTP ${response.status}). Verify the current site authentication and that the 'Documents1' library exists.`);
                }
                throw new Error(`SharePoint upload failed (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 300) || 'No error details returned.'}`);
            }

            const data = await this._readJson<any>(response);
            
            if (!data || !data.ServerRelativeUrl) {
                throw new Error('SharePoint upload succeeded but returned no file metadata.');
            }

            const result = {
                name: data.Name || file.name,
                url: this._buildAbsoluteUrl(data.ServerRelativeUrl),
                serverRelativeUrl: data.ServerRelativeUrl,
                folderName: this._extractFolderNameFromServerRelativeUrl(data.ServerRelativeUrl)
            };

            console.log(`[Upload-Docs1] âœ“ Successfully uploaded: ${result.name}`);
            return result;

        } catch (error) {
            console.error('[Upload-Docs1] âœ— Upload error:', error);
            throw error;
        }
    }

    // Keep compatibility with old name if needed, or just redirect
    public static async uploadFileToDocuments1(file: File, folderName?: string): Promise<{ name: string; url: string; serverRelativeUrl: string; folderName?: string }> {
        return this.uploadFile(file, folderName);
    }

    public static async getFilesFromDocuments1(): Promise<IContentAsset[]> {
        if (!this._getHttpClient()) {
            throw new Error('SharePointService not initialized. Call init() first.');
        }

        try {
            const siteUrl = this._getDocuments1SiteUrl();
            const rootFolderPath = this._buildDocuments1FolderServerRelativeUrl();
            const escapedRootFolderPath = this._escapeODataValue(rootFolderPath);
            const filesEndpoint = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapedRootFolderPath}')/Files`;
            const foldersEndpoint =
                `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapedRootFolderPath}')/Folders` +
                `?$select=Name,ServerRelativeUrl,Files/Name,Files/Length,Files/TimeCreated,Files/TimeLastModified,Files/ServerRelativeUrl,Files/UniqueId` +
                `&$expand=Files`;

            console.log('[GetFiles-Docs1] Fetching files from Documents1', { siteUrl, filesEndpoint, foldersEndpoint });

            const [filesResponse, foldersResponse] = await Promise.all([
                this._getHttpClient().get(
                    filesEndpoint,
                    SPHttpClient.configurations.v1,
                    {
                        headers: this._getJsonHeaders()
                    }
                ),
                this._getHttpClient().get(
                    foldersEndpoint,
                    SPHttpClient.configurations.v1,
                    {
                        headers: this._getJsonHeaders()
                    }
                )
            ]);

            if (!filesResponse.ok) {
                const errorText = await this._readErrorBody(filesResponse);
                const isHtmlError = errorText.trim().toLowerCase().indexOf('<html') !== -1;
                console.error('[GetFiles-Docs1] SharePoint root file listing failed', {
                    status: filesResponse.status,
                    statusText: filesResponse.statusText,
                    responsePreview: errorText.substring(0, 500)
                });
                if (isHtmlError) {
                    throw new Error(`SharePoint returned HTML (HTTP ${filesResponse.status}) instead of file list. Verify 'Documents1' exists.`);
                }
                throw new Error(`Failed to fetch Documents1 root files (HTTP ${filesResponse.status} ${filesResponse.statusText}): ${errorText.substring(0, 200) || 'No error details returned.'}`);
            }

            if (!foldersResponse.ok) {
                const errorText = await this._readErrorBody(foldersResponse);
                const isHtmlError = errorText.trim().toLowerCase().indexOf('<html') !== -1;
                console.error('[GetFiles-Docs1] SharePoint folder listing failed', {
                    status: foldersResponse.status,
                    statusText: foldersResponse.statusText,
                    responsePreview: errorText.substring(0, 500)
                });
                if (isHtmlError) {
                    throw new Error(`SharePoint returned HTML (HTTP ${foldersResponse.status}) instead of folder list. Verify 'Documents1' exists.`);
                }
                throw new Error(`Failed to fetch Documents1 folders (HTTP ${foldersResponse.status} ${foldersResponse.statusText}): ${errorText.substring(0, 200) || 'No error details returned.'}`);
            }

            const filesData = await this._readJson<any>(filesResponse);
            const foldersData = await this._readJson<any>(foldersResponse);
            const rootFiles = filesData.value || [];
            const folderEntries = (foldersData.value || []).filter((folder: any) => (folder?.Name || '').toLowerCase() !== 'forms');
            const nestedFiles = folderEntries.flatMap((folder: any) =>
                (folder?.Files || []).map((file: any) => ({
                    ...file,
                    FolderName: folder.Name
                }))
            );
            const items = [...rootFiles, ...nestedFiles];

            console.log(`[GetFiles-Docs1] Found ${items.length} files including folders`);

            return items.map((file: any) => ({
                id: file.UniqueId || Math.random(),
                name: file.Name || 'Unknown File',
                type: this._guessAssetType(file.Name),
                owner: 'SharePoint',
                status: 'Published',
                dateAdded: this._formatDisplayDate(file.TimeLastModified || file.TimeCreated),
                size: this._formatFileSize(file.Length),
                description: `Managed via Documents1 SharePoint Library`,
                url: this._buildAbsoluteUrl(file.ServerRelativeUrl),
                path: file.ServerRelativeUrl,
                folderName: file.FolderName || this._extractFolderNameFromServerRelativeUrl(file.ServerRelativeUrl)
            }));

        } catch (error) {
            console.error('[GetFiles-Docs1] Error:', error);
            throw error;
        }
    }

    // Keep compatibility
    public static async getDocuments1Files(): Promise<IContentAsset[]> {
        return this.getFilesFromDocuments1();
    }

    public static async refreshDocuments1Files(): Promise<IContentAsset[]> {
        console.log('[RefreshDocs1] Starting Documents1 file refresh');
        return this.getDocuments1Files();
    }

    public static async getDocuments(): Promise<IContentAsset[]> {
        return this.getFilesFromDocuments1();
    }

    private static async _getCustomContentAssets(): Promise<IContentAsset[]> {
        const listName = await this._ensureContentLibraryList();
        const schema = await this._getContentLibraryListSchema();
        const selectFields = [
            'Id',
            'Title',
            'Created',
            'Modified',
            schema.fileLinkField,
            schema.uploadedByField,
            schema.assignedToField,
            schema.statusField,
            schema.folderNameField || '',
            schema.assetTypeField || '',
            schema.descriptionField || '',
            schema.fileSizeField || ''
        ].filter((field, index, collection) => !!field && collection.indexOf(field) === index);
        const endpoint =
            this._getApiUrl(`/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items`) +
            `?$select=${selectFields.join(',')}` +
            `&$orderby=Modified desc` +
            `&$top=5000`;
        const data = await this._safeGetJson<any>(endpoint, `${listName} content assets`);
        return this._dedupeContentLibraryItems(this._toCollection(data), schema)
            .map((item: any) => this._mapContentLibraryItem(item, schema))
            .filter((asset: IContentAsset) => !!asset.url);
    }

    private static async _ensureList(listName: string, baseTemplate: number = 100): Promise<void> {
        if (this._ensuredLists.has(listName)) {
            return;
        }

        const existingEnsure = this._pendingListEnsures.get(listName);
        if (existingEnsure) {
            return existingEnsure;
        }

        const lastFailureAt = this._listFailureTimestamps.get(listName);
        if (lastFailureAt && (Date.now() - lastFailureAt) < this._listRetryCooldownMs) {
            return;
        }

        const ensurePromise = (async () => {
            const siteUrl = this._getSiteUrl();
            const normalizedListName = this._normalizeListTitle(listName);

            try {
                const existingTitles = await this._getExistingListTitles();
                if (existingTitles.has(normalizedListName)) {
                    this._ensuredLists.add(listName);
                    this._listFailureTimestamps.delete(listName);
                    return;
                }

                const digest = await this._getFormDigestValue();
                await this._safePostJson<any>(
                    `${siteUrl}/_api/web/lists`,
                    {
                        headers: this._getJsonHeaders({
                            'X-RequestDigest': digest
                        }),
                        body: JSON.stringify({
                            AllowContentTypes: true,
                            BaseTemplate: baseTemplate,
                            ContentTypesEnabled: true,
                            Description: `LMS System List for ${listName}`,
                            Title: listName
                        })
                    },
                    `create ${listName} list`
                );

                this._rememberListTitle(listName);
                this._ensuredLists.add(listName);
                this._listFailureTimestamps.delete(listName);
            } catch (error) {
                this._listFailureTimestamps.set(listName, Date.now());
                throw error;
            } finally {
                this._pendingListEnsures.delete(listName);
            }
        })();

        this._pendingListEnsures.set(listName, ensurePromise);
        return ensurePromise;
    }

    public static async getContentAssets(): Promise<IContentAsset[]> {
        const metadataAssets = await this._getCustomContentAssets().catch(() => [] as IContentAsset[]);
        if (metadataAssets.length > 0) {
            return this._mergeContentAssetCollections(metadataAssets, [] as IContentAsset[], false);
        }

        return this.getFilesFromDocuments1().catch(() => [] as IContentAsset[]);
    }

    public static async getContentAssetsForUser(userEmail?: string): Promise<IContentAsset[]> {
        void userEmail;
        return this.getContentAssets();
    }

    private static async _deleteContentLibraryMetadataItemById(listName: string, itemId: number, digest: string): Promise<void> {
        const siteUrl = this._getSiteUrl();
        const response = await this._getHttpClient().post(
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items(${itemId})`,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'DELETE',
                    'X-RequestDigest': digest
                })
            }
        );

        if (!response.ok) {
            const errorText = await this._readErrorBody(response);
            if (!this._isMissingSharePointResource(response.status, errorText)) {
                throw new Error(`Failed to delete duplicate content metadata from ${listName} (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 300) || 'No error details returned.'}`);
            }
        }
    }

    private static async _cleanupDuplicateContentLibraryItems(listName: string, schema: IContentLibraryListSchema, normalizedFileKey: string, preferredItemId: number, digest: string): Promise<void> {
        if (!normalizedFileKey) {
            return;
        }

        const siteUrl = this._getSiteUrl();
        const endpoint =
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items` +
            `?$select=Id,Title,Modified,${schema.fileLinkField}${schema.folderNameField ? `,${schema.folderNameField}` : ''}` +
            `&$orderby=Modified desc` +
            `&$top=5000`;
        const data = await this._safeGetJson<any>(endpoint, `${listName} content metadata`);
        const matchingItems = this._toCollection(data)
            .filter((item: any) => this._getContentLibraryItemLookupKey(item, schema) === normalizedFileKey)
            .sort((left: any, right: any) => {
                if (preferredItemId > 0) {
                    if (Number(left?.Id || 0) === preferredItemId) {
                        return -1;
                    }

                    if (Number(right?.Id || 0) === preferredItemId) {
                        return 1;
                    }
                }

                return Number(right?.Id || 0) - Number(left?.Id || 0);
            });

        const duplicateItems = matchingItems.slice(1);
        for (const duplicateItem of duplicateItems) {
            const duplicateItemId = Number(duplicateItem?.Id || duplicateItem?.id || 0);
            if (duplicateItemId > 0) {
                await this._deleteContentLibraryMetadataItemById(listName, duplicateItemId, digest);
            }
        }
    }

    public static async addContentAsset(asset: IContentAsset): Promise<void> {
        const listName = await this._ensureContentLibraryList();
        const siteUrl = this._getSiteUrl();
        const schema = await this._getContentLibraryListSchema();
        const normalizedTitle = (asset?.name || '').toString().trim();
        const normalizedFileUrl = this._resolveServerRelativeUrlFromString(asset?.path || asset?.url || '') ||
            this._buildNormalizedContentFileUrl(asset?.path || asset?.url || '');
        const normalizedFolderName = (
            asset?.folderName ||
            this._extractFolderNameFromServerRelativeUrl(this._resolveServerRelativeAssetUrl(asset))
        ).toString().trim();
        const normalizedUploadedBy = ((asset?.uploadedBy || asset?.owner || this.getCurrentContextUserEmail()) || '').toString().trim().toLowerCase();
        const normalizedAssignedTo = (asset?.assignedTo || '').toString().trim().toLowerCase();
        const normalizedStatus = (asset?.status || 'Uploaded').toString().trim() || 'Uploaded';
        const normalizedType = (asset?.type || this._guessAssetType(normalizedFileUrl || normalizedTitle)).toString().trim();
        const normalizedDescription = (asset?.description || '').toString().trim();
        const normalizedSize = (asset?.size || '').toString().trim();

        if (!normalizedTitle) {
            throw new Error('Content asset title is required.');
        }

        if (!normalizedFileUrl) {
            throw new Error('Content asset FileLink is required.');
        }

        const digest = await this._getFormDigestValue();
        const selectFields = Array.from(new Set([
            'Id',
            'Title',
            schema.fileLinkField,
            schema.assignedToField
        ])).join(',');
        const lookupEndpoint = `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items?$select=${selectFields}&$top=5000`;
        const existingItemsData = await this._safeGetJson<any>(lookupEndpoint, `${listName} content metadata`);
        const requestedId = Number(asset?.id || 0);
        const normalizedFileKey = this._normalizeContentAssetLookupKey(normalizedFileUrl);
        const existingItem = this._toCollection(existingItemsData).find((item: any) => {
            if (requestedId > 0 && Number(item?.Id || item?.id || 0) === requestedId) {
                return true;
            }

            const existingFileUrl = this._buildNormalizedContentFileUrl(this._readFieldValue(item, schema.fileLinkField));
            return !!existingFileUrl && this._normalizeContentAssetLookupKey(existingFileUrl) === normalizedFileKey;
        });

        const payload: Record<string, any> = {
            Title: normalizedTitle,
            [schema.fileLinkField]: normalizedFileUrl,
            [schema.uploadedByField]: normalizedUploadedBy,
            [schema.assignedToField]: normalizedAssignedTo,
            [schema.statusField]: normalizedStatus
        };

        if (schema.folderNameField) {
            payload[schema.folderNameField] = normalizedFolderName;
        }

        if (schema.assetTypeField) {
            payload[schema.assetTypeField] = normalizedType;
        }

        if (schema.descriptionField) {
            payload[schema.descriptionField] = normalizedDescription;
        }

        if (schema.fileSizeField) {
            payload[schema.fileSizeField] = normalizedSize;
        }

        const endpoint = existingItem
            ? `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items(${Number(existingItem.Id || existingItem.id || 0)})`
            : `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items`;
        const headers = existingItem
            ? this._getJsonHeaders({
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE',
                'X-RequestDigest': digest
            })
            : this._getJsonHeaders({
                'X-RequestDigest': digest
            });
        const response = await this._getHttpClient().post(
            endpoint,
            SPHttpClient.configurations.v1,
            {
                headers,
                body: JSON.stringify(payload)
            }
        );

        if (!response.ok) {
            const errorText = await this._readErrorBody(response);
            throw new Error(`Failed to save content metadata into ${listName} (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 300) || 'No error details returned.'}`);
        }

        const savedItemData = existingItem ? null : await this._readJson<any>(response).catch(() => null);
        const savedItemId = existingItem
            ? Number(existingItem?.Id || existingItem?.id || 0)
            : Number(savedItemData?.Id || savedItemData?.d?.Id || savedItemData?.value?.Id || 0);

        try {
            await this._cleanupDuplicateContentLibraryItems(listName, schema, normalizedFileKey, savedItemId, digest);
        } catch (cleanupError) {
            console.warn('[ContentLibrary] Failed to clean duplicate metadata rows.', cleanupError);
        }
    }

    public static async deleteContentAsset(assetOrId: number | IContentAsset): Promise<void> {
        const listName = await this._ensureContentLibraryList();
        const siteUrl = this._getSiteUrl();
        const asset = typeof assetOrId === 'number' ? null : assetOrId;
        const digest = await this._getFormDigestValue();

        if (asset) {
            const serverRelativeUrl = this._resolveServerRelativeAssetUrl(asset);
            if (serverRelativeUrl) {
                await this._recycleContentAssetFile(serverRelativeUrl, digest);
            }
        }

        const schema = await this._getContentLibraryListSchema();
        let metadataItemId = typeof assetOrId === 'number' ? Number(assetOrId) : 0;

        if (!metadataItemId && asset && schema.fileLinkField) {
            const normalizedFileUrl = this._buildNormalizedContentFileUrl(asset.url || asset.path || '');
            if (normalizedFileUrl) {
                const normalizedFileKey = this._normalizeContentAssetLookupKey(normalizedFileUrl);
                const endpoint =
                    `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items` +
                    `?$select=Id,${schema.fileLinkField}`;
                const data = await this._safeGetJson<any>(endpoint, `${listName} content metadata`);
                const existingItem = this._toCollection(data).find((item: any) => {
                    const existingFileUrl = this._buildNormalizedContentFileUrl(this._readFieldValue(item, schema.fileLinkField));
                    return !!existingFileUrl && this._normalizeContentAssetLookupKey(existingFileUrl) === normalizedFileKey;
                });
                metadataItemId = Number(existingItem?.Id || existingItem?.id || 0);
            }
        }

        if (!metadataItemId) {
            return;
        }

        try {
            const response = await this._getHttpClient().post(
                `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items(${metadataItemId})`,
                SPHttpClient.configurations.v1,
                {
                    headers: this._getJsonHeaders({
                        'IF-MATCH': '*',
                        'X-HTTP-Method': 'DELETE',
                        'X-RequestDigest': digest
                    })
                }
            );

            if (!response.ok) {
                const errorText = await this._readErrorBody(response);
                if (!this._isMissingSharePointResource(response.status, errorText)) {
                    throw new Error(`Failed to delete content metadata from ${listName} (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 300) || 'No error details returned.'}`);
                }
            }
        } catch (error) {
            throw error;
        }
    }

    public static async getEnrollments(
        userEmail: string,
        searchText: string = '',
        options: { excludeStatuses?: string[]; } = {}
    ): Promise<IEnrollment[]> {
        const trimmedEmail = (userEmail || '').trim().toLowerCase();
        const trimmedSearch = (searchText || '').toString().trim();
        const context = await this._getEnrollmentListContext();
        const orderByField = context.schema.assignedDateField || 'Created';
        const queryAttempts: Array<{ endpoint: string; label: string }> = [];
        const emailFilters = this._buildEnrollmentUserEmailFilter(context.schema, trimmedEmail);
        const searchFilter = this._buildEnrollmentSearchFilter(context.schema, trimmedSearch);
        const statusExclusionFilter = this._buildEnrollmentStatusExclusionFilter(context.schema, options.excludeStatuses || []);

        emailFilters.forEach((filter, index) => {
            const combinedFilters = [filter, searchFilter, statusExclusionFilter].filter((value) => !!value);
            queryAttempts.push({
                endpoint: this._buildEnrollmentItemsEndpoint(context.escapedListName, context.schema, {
                    filters: combinedFilters,
                    orderByField: `${orderByField} desc`
                }),
                label: index === 0 ? 'enrollment records by user email' : 'enrollment records by alternate email field'
            });
        });

        if (queryAttempts.length === 0) {
            queryAttempts.push({
                endpoint: this._buildEnrollmentItemsEndpoint(context.escapedListName, context.schema, {
                    filters: [searchFilter, statusExclusionFilter].filter((value) => !!value),
                    orderByField: `${orderByField} desc`
                }),
                label: `${context.listName} enrollment records`
            });
        }

        let items: any[] = [];
        let resolvedEndpoint = queryAttempts[0].endpoint;

        for (let index = 0; index < queryAttempts.length; index += 1) {
            const attempt = queryAttempts[index];
            resolvedEndpoint = attempt.endpoint;
            try {
                const data = await this._safeGetJson<any>(attempt.endpoint, attempt.label);
                items = this._toCollection(data);
                break;
            } catch (error) {
                const isLastAttempt = index === queryAttempts.length - 1;
                if (isLastAttempt) {
                    const fallbackEndpoint = this._buildEnrollmentItemsEndpoint(context.escapedListName, context.schema, {
                        orderByField: 'Created desc'
                    });
                    console.warn('[Enrollments] Filtered query failed. Retrying with a plain list query.', {
                        endpoint: attempt.endpoint,
                        error
                    });
                    const fallbackData = await this._safeGetJson<any>(
                        fallbackEndpoint,
                        `${context.listName} enrollment records fallback`
                    );
                    resolvedEndpoint = fallbackEndpoint;
                    items = this._toCollection(fallbackData);
                }
            }
        }

        const excludedStatuses = new Set(
            (options.excludeStatuses || [])
                .map((value) => (value || '').toString().trim().toLowerCase())
                .filter((value) => !!value)
        );
        const mapped = items
            .map((item: any) => this._mapEnrollmentItem(item, context.schema))
            .filter((item: IEnrollment) => !trimmedEmail || (item.userEmail || '').toLowerCase() === trimmedEmail)
            .filter((item: IEnrollment) => {
                if (excludedStatuses.size === 0) {
                    return true;
                }

                const normalizedStatus = (item.listStatus || item.status || '').toString().trim().toLowerCase();
                return !excludedStatuses.has(normalizedStatus);
            })
            .filter((item: IEnrollment) => {
                if (!trimmedSearch) {
                    return true;
                }

                const normalizedSearch = trimmedSearch.toLowerCase();
                return [
                    item.certName,
                    item.certCode,
                    item.certificateName
                ].some((value) => (value || '').toString().toLowerCase().indexOf(normalizedSearch) > -1);
            });

        console.log('[Enrollments] API response', {
            listName: context.listName,
            endpoint: resolvedEndpoint,
            userEmail: trimmedEmail,
            count: mapped.length,
            items: mapped
        });

        return this._dedupeEnrollments(mapped);
    }

    public static async getDepartmentProgressDashboard(): Promise<IDepartmentProgressSummary[]> {
        const enrollments = await this.getEnrollments('');
        const uniqueUsers = Array.from(
            new Map(
                (enrollments || [])
                    .filter((enrollment) => !!(enrollment?.userEmail || '').toString().trim())
                    .map((enrollment) => {
                        const normalizedEmail = (enrollment.userEmail || '').toString().trim().toLowerCase();
                        return [
                            normalizedEmail,
                            {
                                id: normalizedEmail,
                                email: normalizedEmail,
                                Email: normalizedEmail,
                                login: normalizedEmail,
                                LoginName: normalizedEmail,
                                name: enrollment.userName || normalizedEmail,
                                Title: enrollment.userName || normalizedEmail,
                                role: 'Member',
                                siteGroup: 'Members'
                            } as ILearnerDirectoryUser
                        ];
                    })
            ).values()
        );

        const peopleManagerProfiles = await this._fetchPeopleManagerUserProfiles(uniqueUsers).catch(() =>
            new Map<string, { jobTitle: string; department: string }>()
        );

        const mergedLearners: IDepartmentProgressLearner[] = enrollments.map((enrollment) => {
            const lookupKeys = this._getUserLookupKeys({
                Email: enrollment.userEmail,
                LoginName: enrollment.userEmail
            });
            const matchingProfile = lookupKeys
                .map((key) => peopleManagerProfiles.get(key))
                .find((profile) => !!profile);
            const normalizedStatus = (enrollment.status || '').toString().trim().toLowerCase();

            return {
                learner: (enrollment.userName || enrollment.userEmail || 'Not Available').toString(),
                learnerEmail: (enrollment.userEmail || '').toString(),
                path: (enrollment.certName || enrollment.certificateName || enrollment.certCode || 'Not Available').toString(),
                progress: Number(enrollment.progress || 0),
                department: (matchingProfile?.department || 'Not Available').toString(),
                status: normalizedStatus || 'not-started'
            };
        });

        const groupedDepartments = new Map<string, IDepartmentProgressLearner[]>();
        mergedLearners.forEach((learner) => {
            const department = (learner.department || 'Not Available').toString().trim() || 'Not Available';
            const existingLearners = groupedDepartments.get(department) || [];
            existingLearners.push({
                ...learner,
                department
            });
            groupedDepartments.set(department, existingLearners);
        });

        return Array.from(groupedDepartments.entries())
            .map(([department, learners]) => {
                const uniqueLearners = new Set(
                    learners
                        .map((learner) => (learner.learnerEmail || '').toString().trim().toLowerCase())
                        .filter((email) => !!email)
                );
                const completedCount = learners.filter((learner) =>
                    learner.status === 'completed' || Number(learner.progress || 0) >= 100
                ).length;
                const inProgressCount = learners.filter((learner) => {
                    const progress = Number(learner.progress || 0);
                    return !(learner.status === 'completed' || progress >= 100) && progress > 0;
                }).length;
                const notStartedCount = learners.filter((learner) => {
                    const progress = Number(learner.progress || 0);
                    return !(learner.status === 'completed' || progress >= 100) && progress <= 0;
                }).length;
                const enrolledCount = completedCount + inProgressCount;
                const totalLearners = uniqueLearners.size;
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
                    learners: [...learners].sort((left, right) => {
                        const progressDelta = Number(right.progress || 0) - Number(left.progress || 0);
                        if (progressDelta !== 0) {
                            return progressDelta;
                        }

                        return (left.learner || '').localeCompare(right.learner || '');
                    })
                } as IDepartmentProgressSummary;
            })
            .sort((left, right) => left.department.localeCompare(right.department));
    }

    public static async getEnrollmentCountForPath(pathId: string, certName?: string): Promise<number> {
        const normalizedPathId = (pathId || '').toString().trim();
        const normalizedCertName = (certName || '').toString().trim().toLowerCase();
        if (!normalizedPathId && !normalizedCertName) {
            return 0;
        }

        if (normalizedCertName) {
            return this.getEnrollmentCountForCertification(certName || '');
        }

        const context = await this._getEnrollmentListContext();
        if (!context.schema.pathIdField) {
            throw new Error("The 'Enrollments' list must expose a PathId field to calculate seat usage.");
        }

        const endpoint = this._buildEnrollmentItemsEndpoint(context.escapedListName, context.schema, {
            filters: [`${context.schema.pathIdField} eq '${this._escapeODataValue(normalizedPathId)}'`]
        });
        const data = await this._safeGetJson<any>(endpoint, `enrollment count for path ${normalizedPathId}`);
        const items = this._toCollection(data);
        const uniqueEmails = new Set<string>();

        items.forEach((item: any) => {
            const mapped = this._mapEnrollmentItem(item, context.schema);
            const email = (mapped.userEmail || '').toLowerCase();
            if (email) {
                uniqueEmails.add(email);
            }
        });

        console.log('[Enrollments] Seat count response', {
            listName: context.listName,
            pathId: normalizedPathId,
            endpoint,
            count: uniqueEmails.size
        });

        return uniqueEmails.size;
    }

    public static async getEnrollmentCountForCertification(certName: string, certCode?: string): Promise<number> {
        const normalizedCertName = (certName || '').toString().trim();
        const normalizedCertCode = (certCode || '').toString().trim();
        if (!normalizedCertName && !normalizedCertCode) {
            return 0;
        }

        try {
            const context = await this._getEnrollmentListContext();
            const certificationFilters = [
                ...this._buildEnrollmentCertificationCodeFilter(context.schema, normalizedCertCode),
                ...this._buildEnrollmentCertificationFilter(context.schema, normalizedCertName)
            ].filter((value, index, array) => !!value && array.indexOf(value) === index);
            if (certificationFilters.length === 0) {
                throw new Error('Certification field is not available on the Enrollment list.');
            }

            const primaryFilter = certificationFilters[0];
            const endpoint = this._buildEnrollmentItemsEndpoint(context.escapedListName, context.schema, {
                filters: [primaryFilter]
            });
            const data = await this._safeGetJson<any>(endpoint, `enrollment count for certification ${normalizedCertName}`);
            const items = this._toCollection(data);
            const uniqueEmails = new Set<string>();

            items.forEach((item: any) => {
                const mapped = this._mapEnrollmentItem(item, context.schema);
                const email = (mapped.userEmail || '').toLowerCase();
                const mappedCertCode = (mapped.certCode || mapped.pathId || '').toString().trim().toLowerCase();
                const mappedCertName = (mapped.certName || mapped.certificateName || '').toString().trim().toLowerCase();
                const matchesCode = !!normalizedCertCode && mappedCertCode === normalizedCertCode.toLowerCase();
                const matchesTitle = !!normalizedCertName && mappedCertName === normalizedCertName.toLowerCase();
                if (email && (matchesCode || (!normalizedCertCode && matchesTitle) || (matchesCode && matchesTitle))) {
                    uniqueEmails.add(email);
                }
            });

            console.log('[Enrollment] Certification seat count response', {
                listName: context.listName,
                certification: normalizedCertName,
                certCode: normalizedCertCode,
                endpoint,
                count: uniqueEmails.size
            });

            return uniqueEmails.size;
        } catch (error) {
            console.warn('[Enrollment] Direct certification count failed. Falling back to in-memory filtering.', {
                certification: normalizedCertName,
                certCode: normalizedCertCode,
                error
            });
            const enrollments = await this.getEnrollments('');
            return new Set(
                enrollments
                    .filter((item) => {
                        const itemCode = (item.certCode || item.pathId || '').toString().trim().toLowerCase();
                        const itemTitle = (item.certName || item.certificateName || '').toString().trim().toLowerCase();
                        return (!!normalizedCertCode && itemCode === normalizedCertCode.toLowerCase()) ||
                            (!normalizedCertCode && !!normalizedCertName && itemTitle === normalizedCertName.toLowerCase());
                    })
                    .map((item) => (item.userEmail || '').toLowerCase())
                    .filter((value) => !!value)
            ).size;
        }
    }

    public static async getEnrollmentCountForCertificationId(certificationId: number, certName?: string, certCode?: string): Promise<number> {
        const normalizedCertificationId = Number(certificationId || 0);
        if (normalizedCertificationId <= 0) {
            return this.getEnrollmentCountForCertification(certName || '', certCode || '');
        }

        try {
            const context = await this._getEnrollmentListContext();
            const certificationIdFilters = this._buildEnrollmentCertificationIdFilter(context.schema, normalizedCertificationId);
            if (certificationIdFilters.length === 0) {
                throw new Error('Certification lookup field is not available on the Enrollment list.');
            }

            const endpoint = this._buildEnrollmentItemsEndpoint(context.escapedListName, context.schema, {
                filters: [`(${certificationIdFilters.join(' or ')})`]
            });
            const data = await this._safeGetJson<any>(endpoint, `enrollment count for certification id ${normalizedCertificationId}`);
            const items = this._toCollection(data);
            const uniqueEmails = new Set<string>();

            items.forEach((item: any) => {
                const mapped = this._mapEnrollmentItem(item, context.schema);
                const email = (mapped.userEmail || '').toLowerCase();
                if (email && Number(mapped.certificationId || 0) === normalizedCertificationId) {
                    uniqueEmails.add(email);
                }
            });

            console.log('[Enrollment] Certification seat count by id response', {
                listName: context.listName,
                certificationId: normalizedCertificationId,
                endpoint,
                count: uniqueEmails.size
            });

            return uniqueEmails.size;
        } catch (error) {
            console.warn('[Enrollment] Direct certification-id count failed. Falling back to certification-based filtering.', {
                certificationId: normalizedCertificationId,
                certName,
                certCode,
                error
            });

            const enrollments = await this.getEnrollments('');
            const uniqueEmails = new Set(
                enrollments
                    .filter((item) => Number(item.certificationId || 0) === normalizedCertificationId)
                    .map((item) => (item.userEmail || '').toLowerCase())
                    .filter((value) => !!value)
            );

            if (uniqueEmails.size > 0) {
                return uniqueEmails.size;
            }

            return this.getEnrollmentCountForCertification(certName || '', certCode || '');
        }
    }

    private static async _resolveCertificationCatalogId(certificationId: number, certName?: string, certCode?: string): Promise<number> {
        const normalizedCertificationId = Number(certificationId || 0);
        if (normalizedCertificationId > 0) {
            return normalizedCertificationId;
        }

        const certification = await this.getCertificationDetailsByCodeOrTitle(certCode || '', certName || '', true);
        return Number(certification?.id || 0);
    }

    public static async syncCertificationAssignedLearnerCount(certificationId: number, certName?: string, certCode?: string): Promise<number> {
        const resolvedCertificationId = await this._resolveCertificationCatalogId(certificationId, certName, certCode);
        if (resolvedCertificationId <= 0) {
            return 0;
        }

        const assignedLearnerCount = await this.getEnrollmentCountForCertificationId(resolvedCertificationId, certName || '', certCode || '');
        await this.updateCertificationEnrolledCount(resolvedCertificationId, assignedLearnerCount);
        return assignedLearnerCount;
    }

    private static async _syncCertificationAssignedLearnerCounts(records: Array<Partial<IEnrollment> | null | undefined>): Promise<void> {
        const targets = Array.from(
            new Map(
                (records || [])
                    .map((record) => {
                        if (!record) {
                            return null;
                        }

                        const certificationId = Number(record.certificationId || 0);
                        const certName = (record.certName || record.certificateName || '').toString().trim();
                        const certCode = (record.certCode || record.pathId || '').toString().trim();
                        const dedupeKey = certificationId > 0
                            ? `id:${certificationId}`
                            : `lookup:${this._normalizeCertificationCode(certCode) || certName.toLowerCase()}`;

                        if ((!certName && !certCode) && certificationId <= 0) {
                            return null;
                        }

                        return [
                            dedupeKey,
                            {
                                certificationId,
                                certName,
                                certCode
                            }
                        ] as const;
                    })
                    .filter((entry): entry is readonly [string, { certificationId: number; certName: string; certCode: string; }] => !!entry)
            ).values()
        );

        for (const target of targets) {
            await this.syncCertificationAssignedLearnerCount(target.certificationId, target.certName, target.certCode).catch((error) => {
                console.warn('[Certifications] Enrollment saved but assigned learner count sync failed.', {
                    certificationId: target.certificationId,
                    certName: target.certName,
                    certCode: target.certCode,
                    error
                });
            });
        }
    }

    public static async getEnrollmentSeatUsageMap(): Promise<Map<string, number>> {
        const seatUsageMap = new Map<string, Set<string>>();
        const enrollments = await this.getEnrollments('');

        enrollments.forEach((item) => {
            const normalizedEmail = (item.userEmail || '').toString().trim().toLowerCase();
            const normalizedCode = this._normalizeCertificationCode(item.certCode || item.pathId || '');
            const normalizedTitle = (item.certName || item.certificateName || '').toString().trim().toLowerCase();

            if (!normalizedEmail) {
                return;
            }

            if (normalizedCode) {
                if (!seatUsageMap.has(normalizedCode)) {
                    seatUsageMap.set(normalizedCode, new Set<string>());
                }
                seatUsageMap.get(normalizedCode)?.add(normalizedEmail);
            }

            if (normalizedTitle) {
                if (!seatUsageMap.has(normalizedTitle)) {
                    seatUsageMap.set(normalizedTitle, new Set<string>());
                }
                seatUsageMap.get(normalizedTitle)?.add(normalizedEmail);
            }
        });

        const normalizedCounts = new Map<string, number>();
        seatUsageMap.forEach((emails, key) => {
            normalizedCounts.set(key, emails.size);
        });

        console.log('[Enrollment] Seat usage map generated', {
            enrollmentCount: enrollments.length,
            certificationCount: normalizedCounts.size
        });

        return normalizedCounts;
    }

    public static async hasEnrollmentForUserCertification(userEmail: string, certName: string, certCode?: string): Promise<boolean> {
        const normalizedEmail = (userEmail || '').toString().trim().toLowerCase();
        const normalizedCertName = (certName || '').toString().trim();
        const normalizedCertCode = (certCode || '').toString().trim();
        if (!normalizedEmail || (!normalizedCertName && !normalizedCertCode)) {
            return false;
        }

        try {
            const context = await this._getEnrollmentListContext();
            const emailFilters = this._buildEnrollmentUserEmailFilter(context.schema, normalizedEmail);
            const certificationFilters = [
                ...this._buildEnrollmentCertificationCodeFilter(context.schema, normalizedCertCode),
                ...this._buildEnrollmentCertificationFilter(context.schema, normalizedCertName)
            ].filter((value, index, array) => !!value && array.indexOf(value) === index);
            if (emailFilters.length === 0 || certificationFilters.length === 0) {
                throw new Error('Enrollment email or certification fields are not available on the list.');
            }

            const combinedFilters = [
                `(${emailFilters.join(' or ')})`,
                `(${certificationFilters.join(' or ')})`
            ];
            const endpoint = this._buildEnrollmentItemsEndpoint(context.escapedListName, context.schema, {
                filters: combinedFilters,
                top: 25
            });
            const data = await this._safeGetJson<any>(endpoint, `duplicate enrollment check for ${normalizedEmail}`);
            const items = this._toCollection(data);
            const alreadyEnrolled = items.some((item: any) => {
                const mapped = this._mapEnrollmentItem(item, context.schema);
                const mappedCertCode = (mapped.certCode || mapped.pathId || '').toString().trim().toLowerCase();
                return (mapped.userEmail || '').toLowerCase() === normalizedEmail &&
                    (
                        (!!normalizedCertCode && mappedCertCode === normalizedCertCode.toLowerCase()) ||
                        (!normalizedCertCode && (mapped.certName || mapped.certificateName || '').toString().trim().toLowerCase() === normalizedCertName.toLowerCase())
                    );
            });

            console.log('[Enrollment] Duplicate check response', {
                listName: context.listName,
                endpoint,
                userEmail: normalizedEmail,
                certification: normalizedCertName,
                certCode: normalizedCertCode,
                alreadyEnrolled
            });

            return alreadyEnrolled;
        } catch (error) {
            console.warn('[Enrollment] Direct duplicate check failed. Falling back to scoped enrollment read.', {
                userEmail: normalizedEmail,
                certification: normalizedCertName,
                certCode: normalizedCertCode,
                error
            });
            const enrollments = await this.getEnrollments(normalizedEmail);
            return enrollments.some((item) =>
                (item.userEmail || '').toLowerCase() === normalizedEmail &&
                (
                    (!!normalizedCertCode && (item.certCode || item.pathId || '').toString().trim().toLowerCase() === normalizedCertCode.toLowerCase()) ||
                    (!normalizedCertCode && (item.certName || item.certificateName || '').toString().trim().toLowerCase() === normalizedCertName.toLowerCase())
                )
            );
        }
    }

    public static async hasEnrollmentForUserCertificationId(
        userEmail: string,
        certificationId: number,
        certName?: string,
        certCode?: string
    ): Promise<boolean> {
        const normalizedEmail = (userEmail || '').toString().trim().toLowerCase();
        const normalizedCertificationId = Number(certificationId || 0);
        if (!normalizedEmail || normalizedCertificationId <= 0) {
            return this.hasEnrollmentForUserCertification(userEmail, certName || '', certCode || '');
        }

        try {
            const context = await this._getEnrollmentListContext();
            const emailFilters = this._buildEnrollmentUserEmailFilter(context.schema, normalizedEmail);
            const certificationIdFilters = this._buildEnrollmentCertificationIdFilter(context.schema, normalizedCertificationId);
            if (emailFilters.length === 0 || certificationIdFilters.length === 0) {
                throw new Error('Enrollment email or certification lookup fields are not available on the list.');
            }

            const endpoint = this._buildEnrollmentItemsEndpoint(context.escapedListName, context.schema, {
                filters: [
                    `(${emailFilters.join(' or ')})`,
                    `(${certificationIdFilters.join(' or ')})`
                ],
                top: 25
            });
            const data = await this._safeGetJson<any>(endpoint, `duplicate enrollment check for certification id ${normalizedCertificationId}`);
            const items = this._toCollection(data);
            const alreadyEnrolled = items.some((item: any) => {
                const mapped = this._mapEnrollmentItem(item, context.schema);
                return (mapped.userEmail || '').toLowerCase() === normalizedEmail &&
                    Number(mapped.certificationId || 0) === normalizedCertificationId;
            });

            console.log('[Enrollment] Duplicate check by certification id response', {
                listName: context.listName,
                userEmail: normalizedEmail,
                certificationId: normalizedCertificationId,
                endpoint,
                alreadyEnrolled
            });

            return alreadyEnrolled;
        } catch (error) {
            console.warn('[Enrollment] Direct certification-id duplicate check failed. Falling back to user enrollment read.', {
                userEmail: normalizedEmail,
                certificationId: normalizedCertificationId,
                certName,
                certCode,
                error
            });

            const enrollments = await this.getEnrollments(normalizedEmail);
            const alreadyEnrolled = enrollments.some((item) =>
                (item.userEmail || '').toLowerCase() === normalizedEmail &&
                Number(item.certificationId || 0) === normalizedCertificationId
            );

            if (alreadyEnrolled) {
                return true;
            }

            return this.hasEnrollmentForUserCertification(userEmail, certName || '', certCode || '');
        }
    }

    public static async createEnrollmentForCertificationAssignment(params: {
        userEmail: string;
        userName: string;
        certName: string;
        certCode?: string;
        examScheduledDate: string;
        assignedByName?: string;
        assignedById?: number;
        assignedByAdmin?: boolean;
        pathId?: string;
    }): Promise<number> {
        const normalizedUserEmail = (params.userEmail || '').toString().trim().toLowerCase();
        const normalizedUserName = (params.userName || '').toString().trim();
        const requestedCertName = (params.certName || '').toString().trim();
        const examScheduledDate = (params.examScheduledDate || '').toString().trim();

        if (!normalizedUserEmail || !requestedCertName) {
            throw new Error('Certification not found');
        }

        if (!examScheduledDate) {
            throw new Error('A valid exam date is required');
        }

        const certification = await this.getCertificationDetailsByCodeOrTitle(params.certCode || '', requestedCertName, true);
        if (!certification) {
            throw new Error('Certification not found');
        }

        const canonicalCertName = (certification.title || requestedCertName).toString().trim();
        const canonicalCertCode = (certification.code || params.certCode || canonicalCertName).toString().trim();

        const alreadyEnrolled = certification.id
            ? await this.hasEnrollmentForUserCertificationId(normalizedUserEmail, certification.id, canonicalCertName, canonicalCertCode)
            : await this.hasEnrollmentForUserCertification(normalizedUserEmail, canonicalCertName, canonicalCertCode);
        if (alreadyEnrolled) {
            throw new Error('Already enrolled');
        }

        const nowIso = new Date().toISOString();
        return this.addOrUpdateEnrollment({
            userEmail: normalizedUserEmail,
            userName: normalizedUserName || normalizedUserEmail,
            certificationId: certification.id,
            certCode: canonicalCertCode,
            certName: canonicalCertName,
            pathId: (params.pathId || canonicalCertCode).toString().trim(),
            startDate: nowIso,
            endDate: examScheduledDate,
            status: 'scheduled',
            progress: 0,
            assignedByAdmin: params.assignedByAdmin !== false,
            assignedDate: nowIso,
            assignedByName: params.assignedByName,
            assignedById: params.assignedById,
            examScheduledDate
        }, {
            failIfExists: true,
            minimalFieldsOnly: false
        });
    }

    public static async fetchUserEnrollments(userEmail?: string, searchText: string = ''): Promise<IEnrollment[]> {
        const normalizedEmail = (userEmail || this.getCurrentContextUserEmail()).toString().trim().toLowerCase();
        const currentUserId = this.getCurrentContextUserId() || await this._getSiteUserIdByEmail(normalizedEmail);
        const enrollments = await this.getEnrollments(normalizedEmail, searchText);
        const syncResult = await this.syncDeletedAuditLogs(normalizedEmail, currentUserId || undefined, enrollments);
        const deletedPathIds = new Set(
            syncResult.deletedLogs
                .map((log) => (log.pathId || log.assignmentName || '').toString().trim().toLowerCase())
                .filter((value) => !!value)
        );

        if (syncResult.removedEnrollmentCount > 0) {
            const refreshed = await this.getEnrollments(normalizedEmail, searchText);
            return refreshed.filter((item) => !deletedPathIds.has(this._getEnrollmentPathId(item).toLowerCase()));
        }

        return enrollments.filter((item) => !deletedPathIds.has(this._getEnrollmentPathId(item).toLowerCase()));
    }

    public static getCurrentContextUserEmail(): string {
        return (this._context?.pageContext?.user?.email || '').toString().trim();
    }

    public static getCurrentContextUserId(): number | null {
        const contextUserId = Number(this._context?.pageContext?.legacyPageContext?.userId || 0);
        return contextUserId > 0 ? contextUserId : null;
    }

    public static getCurrentContextUserName(): string {
        return this._getCurrentContextUserName();
    }

    public static emitEnrollmentRefreshSignal(): void {
        if (typeof window === 'undefined') {
            return;
        }

        window.dispatchEvent(new Event(LMS_ENROLLMENTS_REFRESH_EVENT));
    }

    private static _parseDateOnlyInput(value: string): Date | null {
        const normalizedValue = (value || '').toString().trim();
        if (!normalizedValue) {
            return null;
        }

        const dateOnlyMatch = normalizedValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (dateOnlyMatch) {
            const year = Number(dateOnlyMatch[1]);
            const month = Number(dateOnlyMatch[2]);
            const day = Number(dateOnlyMatch[3]);
            const parsedDate = new Date(year, month - 1, day);

            if (
                parsedDate.getFullYear() === year &&
                parsedDate.getMonth() === month - 1 &&
                parsedDate.getDate() === day
            ) {
                return parsedDate;
            }

            return null;
        }

        const parsedValue = new Date(normalizedValue);
        if (Number.isNaN(parsedValue.getTime())) {
            return null;
        }

        return new Date(parsedValue.getFullYear(), parsedValue.getMonth(), parsedValue.getDate());
    }

    private static _normalizeCompletionExamDate(examDate: string): string {
        const parsedExamDate = this._parseDateOnlyInput(examDate);
        if (!parsedExamDate) {
            throw new Error('A valid exam date is required.');
        }

        const today = new Date();
        const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
        if (parsedExamDate.getTime() > todayStart.getTime()) {
            throw new Error('Completion date cannot be in the future');
        }

        return new Date(Date.UTC(
            parsedExamDate.getFullYear(),
            parsedExamDate.getMonth(),
            parsedExamDate.getDate()
        )).toISOString();
    }

    private static _normalizeCertificationCompletionDates(examDate: string, renewalDate: string): { examDateIso: string; renewalDateIso: string; } {
        const examDateIso = this._normalizeCompletionExamDate(examDate);
        const parsedExamDate = this._parseDateOnlyInput(examDate);
        const parsedRenewalDate = this._parseDateOnlyInput(renewalDate);
        if (!parsedRenewalDate) {
            throw new Error('A valid renewal date is required.');
        }

        if (!parsedExamDate || parsedRenewalDate.getTime() <= parsedExamDate.getTime()) {
            throw new Error('Renewal date must be after exam date');
        }

        return {
            examDateIso,
            renewalDateIso: new Date(Date.UTC(
                parsedRenewalDate.getFullYear(),
                parsedRenewalDate.getMonth(),
                parsedRenewalDate.getDate()
            )).toISOString()
        };
    }

    private static _mapCertificationCompletionRecordItem(item: any, schema: ICertificationCompletionListSchema): ICertificationCompletionRecord {
        return {
            id: Number(item?.Id || item?.id || 0),
            title: (item?.Title || '').toString().trim(),
            certId: (this._readFieldValue(item, schema.certIdField) || '').toString().trim(),
            examDate: (this._readFieldValue(item, schema.examDateField) || '').toString(),
            renewalDate: (this._readFieldValue(item, schema.renewalDateField) || '').toString(),
            examCode: (this._readFieldValue(item, schema.examCodeField) || '').toString().trim(),
            created: (item?.Created || '').toString(),
            modified: (item?.Modified || '').toString(),
            authorEmail: (item?.Author?.EMail || item?.Author?.Email || '').toString().trim().toLowerCase(),
            authorName: (item?.Author?.Title || '').toString().trim()
        };
    }

    private static _normalizeRenewalMatchValue(value?: string): string {
        return (value || '').toString().trim().toLowerCase();
    }

    private static _appendRenewalMatchKeys(keySet: Set<string>, learnerEmail: string, values: Array<string | undefined>): void {
        const normalizedEmail = (learnerEmail || '').toString().trim().toLowerCase();
        if (!normalizedEmail) {
            return;
        }

        values.forEach((value) => {
            const normalizedValue = this._normalizeRenewalMatchValue(value);
            if (normalizedValue) {
                keySet.add(`${normalizedEmail}::${normalizedValue}`);
            }
        });
    }

    public static async fetchUserCertificationCompletions(userEmail?: string): Promise<ICertificationCompletionRecord[]> {
        const normalizedUserEmail = (userEmail || this.getCurrentContextUserEmail() || '').toString().trim().toLowerCase();
        if (!normalizedUserEmail) {
            return [];
        }

        try {
            const siteUrl = this._ensureProductionSiteUrl();
            const listName = await this._ensureCertificationCompletionList();
            const schema = await this._getCertificationCompletionListSchema();
            const escapedListName = this._escapeODataValue(listName);
            const selectFields = Array.from(new Set([
                'Id',
                'Title',
                'Created',
                'Modified',
                'Author/EMail',
                'Author/Title',
                schema.certIdField || '',
                schema.examDateField || '',
                schema.renewalDateField || '',
                schema.examCodeField || ''
            ].filter((field) => !!field)));
            const endpoint =
                `${siteUrl}/_api/web/lists/getbytitle('${escapedListName}')/items` +
                `?$select=${selectFields.join(',')}` +
                `&$expand=Author` +
                `&$orderby=Modified desc` +
                `&$top=5000` +
                `&_=${Date.now()}`;

            const data = await this._safeGetJson<any>(endpoint, `${listName} completion records`);
            return this._toCollection(data)
                .map((item: any) => this._mapCertificationCompletionRecordItem(item, schema))
                .filter((item) => item.authorEmail === normalizedUserEmail);
        } catch (error) {
            console.warn('[uploadlist] Failed to fetch completion records for current user.', {
                userEmail: normalizedUserEmail,
                error
            });
            return [];
        }
    }

    public static async getUpcomingRenewalRecords(daysAhead: number = 30): Promise<IUpcomingRenewalRecord[]> {
        const normalizedDaysAhead = Math.max(1, Math.floor(Number(daysAhead) || 30));

        try {
            const siteUrl = this._ensureProductionSiteUrl();
            const listName = await this._ensureCertificationCompletionList();
            const schema = await this._getCertificationCompletionListSchema();
            if (!schema.renewalDateField) {
                return [];
            }

            const today = new Date();
            const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
            const windowEnd = new Date(todayStart);
            windowEnd.setDate(windowEnd.getDate() + normalizedDaysAhead);
            windowEnd.setHours(23, 59, 59, 999);

            const escapedListName = this._escapeODataValue(listName);
            const selectFields = Array.from(new Set([
                'Id',
                'Title',
                'Created',
                'Modified',
                'Author/EMail',
                'Author/Title',
                schema.certIdField || '',
                schema.examDateField || '',
                schema.renewalDateField || '',
                schema.examCodeField || ''
            ].filter((field) => !!field)));
            const filter = `${schema.renewalDateField} ge datetime'${todayStart.toISOString()}' and ${schema.renewalDateField} le datetime'${windowEnd.toISOString()}'`;
            const endpoint =
                `${siteUrl}/_api/web/lists/getbytitle('${escapedListName}')/items` +
                `?$select=${selectFields.join(',')}` +
                `&$expand=Author` +
                `&$filter=${encodeURIComponent(filter)}` +
                `&$orderby=${schema.renewalDateField} asc` +
                `&$top=5000` +
                `&_=${Date.now()}`;

            const [completionData, enrollments] = await Promise.all([
                this._safeGetJson<any>(endpoint, `${listName} upcoming renewals`),
                this.getEnrollments('', '', {
                    excludeStatuses: ['Not Started']
                }).catch(() => [] as IEnrollment[])
            ]);

            const completedEnrollmentKeys = new Set<string>();
            const completedEnrollmentByKey = new Map<string, IEnrollment>();

            (enrollments || []).forEach((enrollment) => {
                const normalizedStatus = this._normalizeEnrollmentStatus(enrollment?.status || enrollment?.listStatus);
                const progress = Number(enrollment?.progress || 0);
                if (normalizedStatus !== 'completed' && progress < 100) {
                    return;
                }

                const learnerEmail = (enrollment?.userEmail || '').toString().trim().toLowerCase();
                if (!learnerEmail) {
                    return;
                }

                const matchValues = [
                    enrollment?.certCode,
                    enrollment?.pathId,
                    enrollment?.certName,
                    enrollment?.certificateName,
                    enrollment?.examCode
                ];
                this._appendRenewalMatchKeys(completedEnrollmentKeys, learnerEmail, matchValues);

                matchValues.forEach((value) => {
                    const normalizedValue = this._normalizeRenewalMatchValue(value);
                    if (normalizedValue) {
                        completedEnrollmentByKey.set(`${learnerEmail}::${normalizedValue}`, enrollment);
                    }
                });
            });

            return this._toCollection(completionData)
                .map((item: any) => this._mapCertificationCompletionRecordItem(item, schema))
                .map((record) => {
                    const learnerEmail = (record.authorEmail || '').toString().trim().toLowerCase();
                    const matchCandidates = [
                        record.certId,
                        record.examCode,
                        record.title
                    ]
                        .map((value) => this._normalizeRenewalMatchValue(value))
                        .filter((value) => !!value);
                    const matchedEnrollment = matchCandidates
                        .map((value) => completedEnrollmentByKey.get(`${learnerEmail}::${value}`))
                        .find((value): value is IEnrollment => !!value);
                    const isCompleted = !!learnerEmail && matchCandidates.some((value) => completedEnrollmentKeys.has(`${learnerEmail}::${value}`));
                    if (!isCompleted) {
                        return null;
                    }

                    const parsedRenewalDate = new Date(record.renewalDate);
                    if (Number.isNaN(parsedRenewalDate.getTime())) {
                        return null;
                    }

                    const normalizedRenewalDate = new Date(parsedRenewalDate.getFullYear(), parsedRenewalDate.getMonth(), parsedRenewalDate.getDate());
                    const daysUntilRenewal = Math.ceil((normalizedRenewalDate.getTime() - todayStart.getTime()) / (24 * 60 * 60 * 1000));
                    if (daysUntilRenewal < 0 || daysUntilRenewal > normalizedDaysAhead) {
                        return null;
                    }

                    return {
                        ...record,
                        learnerEmail,
                        learnerName: (matchedEnrollment?.userName || record.authorName || learnerEmail).toString().trim(),
                        daysUntilRenewal,
                        urgency: daysUntilRenewal <= 7 ? 'urgent' : 'soon'
                    } as IUpcomingRenewalRecord;
                })
                .filter((record): record is IUpcomingRenewalRecord => !!record)
                .sort((left, right) => {
                    const dateDelta = new Date(left.renewalDate).getTime() - new Date(right.renewalDate).getTime();
                    if (dateDelta !== 0) {
                        return dateDelta;
                    }

                    return (left.learnerName || '').localeCompare(right.learnerName || '');
                });
        } catch (error) {
            console.warn('[uploadlist] Failed to fetch upcoming renewal records.', {
                daysAhead: normalizedDaysAhead,
                error
            });
            return [];
        }
    }

    public static async markCertificationCompleted(params: {
        certificationName: string;
        certId: string;
        examDate: string;
        renewalDate: string;
        examCode?: string;
    }): Promise<ICertificationCompletionRecord> {
        const certificationName = (params.certificationName || '').toString().trim();
        const certId = (params.certId || '').toString().trim();
        const examDate = (params.examDate || '').toString().trim();
        const renewalDate = (params.renewalDate || '').toString().trim();
        const examCode = (params.examCode || '').toString().trim();

        if (!certificationName) {
            throw new Error('Certification name is required.');
        }

        if (!certId) {
            throw new Error('CertID is required.');
        }

        const siteUrl = this._ensureProductionSiteUrl();
        const listName = await this._ensureCertificationCompletionList();
        const schema = await this._getCertificationCompletionListSchema();
        if (!schema.certIdField || !schema.examDateField || !schema.renewalDateField) {
            throw new Error("The 'uploadlist' list must contain CertID, ExamDate, and RenewalDate columns.");
        }

        const normalizedDates = this._normalizeCertificationCompletionDates(examDate, renewalDate);
        const digest = await this._getFormDigestValue();
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items`;
        const payload = {
            Title: certificationName,
            [schema.certIdField]: certId,
            [schema.examDateField]: normalizedDates.examDateIso,
            [schema.renewalDateField]: normalizedDates.renewalDateIso,
            ...(schema.examCodeField && examCode ? { [schema.examCodeField]: examCode } : {})
        };

        const createdItem = await this._safePostJson<any>(
            endpoint,
            {
                headers: this._getJsonHeaders({
                    'X-RequestDigest': digest
                }),
                body: JSON.stringify(payload)
            },
            `${listName} completion item`
        );

        return {
            id: Number(createdItem?.Id || createdItem?.id || 0),
            title: certificationName,
            certId,
            examDate: payload[schema.examDateField],
            renewalDate: payload[schema.renewalDateField],
            examCode,
            created: (createdItem?.Created || '').toString(),
            modified: (createdItem?.Modified || '').toString(),
            authorEmail: (createdItem?.Author?.EMail || createdItem?.Author?.Email || this.getCurrentContextUserEmail() || '').toString().trim().toLowerCase(),
            authorName: (createdItem?.Author?.Title || this.getCurrentContextUserName() || '').toString().trim()
        };
    }

    public static async updateCertificationCompletionRecord(itemId: number, params: {
        certificationName: string;
        certId: string;
        examDate: string;
        renewalDate: string;
        examCode?: string;
    }): Promise<ICertificationCompletionRecord> {
        const normalizedItemId = Number(itemId);
        if (!Number.isFinite(normalizedItemId) || normalizedItemId <= 0) {
            throw new Error('A valid completion record ID is required.');
        }

        const certificationName = (params.certificationName || '').toString().trim();
        const certId = (params.certId || '').toString().trim();
        const examCode = (params.examCode || '').toString().trim();
        if (!certId) {
            throw new Error('CertID is required.');
        }

        const siteUrl = this._ensureProductionSiteUrl();
        const listName = await this._ensureCertificationCompletionList();
        const schema = await this._getCertificationCompletionListSchema();
        if (!schema.examDateField || !schema.renewalDateField) {
            throw new Error("The 'uploadlist' list must contain ExamDate and RenewalDate columns.");
        }

        const normalizedDates = this._normalizeCertificationCompletionDates(params.examDate, params.renewalDate);
        const digest = await this._getFormDigestValue();
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items(${normalizedItemId})`;
        const payload = {
            [schema.examDateField]: normalizedDates.examDateIso,
            [schema.renewalDateField]: normalizedDates.renewalDateIso,
            ...(schema.examCodeField ? { [schema.examCodeField]: examCode } : {})
        };

        const response = await this._getHttpClient().post(
            endpoint,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'MERGE',
                    'X-RequestDigest': digest
                }),
                body: JSON.stringify(payload)
            }
        );

        if (!response.ok) {
            const errorText = await this._readErrorBody(response);
            throw new Error(`Failed to update completion record (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 400) || 'No error details returned.'}`);
        }

        return {
            id: normalizedItemId,
            title: certificationName,
            certId,
            examDate: normalizedDates.examDateIso,
            renewalDate: normalizedDates.renewalDateIso,
            examCode,
            modified: new Date().toISOString(),
            authorEmail: (this.getCurrentContextUserEmail() || '').toString().trim().toLowerCase(),
            authorName: (this.getCurrentContextUserName() || '').toString().trim()
        };
    }

    public static async deleteCertificationCompletionRecord(itemId: number): Promise<void> {
        const normalizedItemId = Number(itemId);
        if (!Number.isFinite(normalizedItemId) || normalizedItemId <= 0) {
            throw new Error('A valid completion record ID is required.');
        }

        const siteUrl = this._ensureProductionSiteUrl();
        const listName = await this._ensureCertificationCompletionList();
        const digest = await this._getFormDigestValue();
        const response = await this._getHttpClient().post(
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items(${normalizedItemId})`,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'DELETE',
                    'X-RequestDigest': digest
                })
            }
        );

        if (!response.ok) {
            const errorText = await this._readErrorBody(response);
            throw new Error(`Failed to delete completion record (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 400) || 'No error details returned.'}`);
        }
    }

    private static _resolveStoredEnrollmentStatusAfterCompletionUndo(enrollment?: Partial<IEnrollment> | null): 'Assigned' | 'Scheduled' | 'Rescheduled' {
        const rescheduledDate = (enrollment?.rescheduledDate || '').toString().trim();
        if (rescheduledDate) {
            return 'Rescheduled';
        }

        const scheduledDate = (
            enrollment?.examScheduledDate ||
            enrollment?.endDate ||
            enrollment?.startDate ||
            ''
        ).toString().trim();
        if (scheduledDate) {
            return 'Scheduled';
        }

        return 'Assigned';
    }

    public static async undoEnrollmentCompletion(params: {
        enrollmentId?: number;
        userEmail?: string;
        certificationId?: number;
        certName?: string;
        certCode?: string;
    }): Promise<number> {
        const normalizedEnrollmentId = Number(params.enrollmentId || 0);
        const normalizedUserEmail = (params.userEmail || '').toString().trim().toLowerCase();
        const normalizedCertName = (params.certName || '').toString().trim();
        const normalizedCertCode = (params.certCode || '').toString().trim();
        const normalizedCertificationId = Number(params.certificationId || 0);

        const listName = await this._resolveEnrollmentListName();
        await this._ensureList(listName);
        await this._ensureEnrollmentCompletionFields(listName);
        const context = await this._getEnrollmentListContext();

        let targetEnrollmentId = normalizedEnrollmentId;
        let currentEnrollment: IEnrollment | null = null;

        if (targetEnrollmentId > 0) {
            try {
                const detailEndpoint =
                    `${context.siteUrl}/_api/web/lists/getbytitle('${context.escapedListName}')/items(${targetEnrollmentId})` +
                    `?$select=${this._getEnrollmentSelectFields(context.schema).join(',')}` +
                    `${this._getEnrollmentExpandFields(context.schema).length > 0 ? `&$expand=${this._getEnrollmentExpandFields(context.schema).join(',')}` : ''}`;
                const enrollmentItem = await this._safeGetJson<any>(detailEndpoint, `enrollment ${targetEnrollmentId}`);
                currentEnrollment = this._mapEnrollmentItem(enrollmentItem, context.schema);
            } catch (error) {
                console.warn('[Enrollments] Failed to read enrollment before undoing completion.', {
                    listName: context.listName,
                    enrollmentId: targetEnrollmentId,
                    error
                });
            }
        }

        if (targetEnrollmentId <= 0 && normalizedUserEmail) {
            const existingEnrollments = await this.getEnrollments(normalizedUserEmail).catch(() => [] as IEnrollment[]);
            currentEnrollment = existingEnrollments.find((item) =>
                (normalizedCertificationId > 0 && Number(item.certificationId || 0) === normalizedCertificationId) ||
                (!!normalizedCertCode && (item.certCode || '').toString().trim().toLowerCase() === normalizedCertCode.toLowerCase()) ||
                (!!normalizedCertName && (item.certName || item.certificateName || '').toString().trim().toLowerCase() === normalizedCertName.toLowerCase())
            ) || null;
            targetEnrollmentId = Number(currentEnrollment?.id || 0);
        }

        if (targetEnrollmentId <= 0) {
            return 0;
        }

        const digest = await this._getFormDigestValue();
        const revertedStatus = this._resolveStoredEnrollmentStatusAfterCompletionUndo(currentEnrollment);
        const payload: Record<string, any> = {
            Title: (currentEnrollment?.certName || currentEnrollment?.certificateName || normalizedCertName || normalizedCertCode || 'Enrollment').toString().trim()
        };

        if (context.schema.statusField) {
            payload[context.schema.statusField] = revertedStatus;
        } else {
            payload.Status = revertedStatus;
        }

        if (context.schema.progressField) {
            payload[context.schema.progressField] = 0;
        }

        if (context.schema.completionDateField) {
            payload[context.schema.completionDateField] = null;
        }

        if (context.schema.examCodeField) {
            payload[context.schema.examCodeField] = null;
        }

        const response = await this._getHttpClient().post(
            `${context.siteUrl}/_api/web/lists/getbytitle('${context.escapedListName}')/items(${targetEnrollmentId})`,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'MERGE',
                    'X-RequestDigest': digest
                }),
                body: JSON.stringify(payload)
            }
        );

        if (!response.ok) {
            const errorText = await this._readErrorBody(response);
            throw new Error(`Failed to undo enrollment completion (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 400) || 'No error details returned.'}`);
        }

        this.emitEnrollmentRefreshSignal();
        return targetEnrollmentId;
    }

    public static async syncEnrollmentCompletion(params: {
        enrollmentId?: number;
        userEmail?: string;
        userName?: string;
        certificationId?: number;
        certName: string;
        certCode?: string;
        examDate: string;
        examCode?: string;
    }): Promise<number> {
        const normalizedEnrollmentId = Number(params.enrollmentId || 0);
        const normalizedUserEmail = (params.userEmail || '').toString().trim().toLowerCase();
        const normalizedUserName = (params.userName || '').toString().trim();
        const normalizedCertName = (params.certName || '').toString().trim();
        const normalizedCertCode = (params.certCode || '').toString().trim();
        const normalizedExamCode = (params.examCode || '').toString().trim();
        const normalizedCertificationId = Number(params.certificationId || 0);
        const normalizedExamDateIso = this._normalizeCompletionExamDate(params.examDate);

        if (!normalizedCertName && !normalizedCertCode && normalizedEnrollmentId <= 0) {
            throw new Error('A certification reference is required to sync the enrollment completion.');
        }

        const listName = await this._resolveEnrollmentListName();
        await this._ensureList(listName);
        await this._ensureEnrollmentCompletionFields(listName);
        const context = await this._getEnrollmentListContext();

        let targetEnrollmentId = normalizedEnrollmentId;
        if (targetEnrollmentId <= 0 && normalizedUserEmail) {
            const existingEnrollments = await this.getEnrollments(normalizedUserEmail).catch(() => [] as IEnrollment[]);
            const matchedEnrollment = existingEnrollments.find((item) =>
                (normalizedCertificationId > 0 && Number(item.certificationId || 0) === normalizedCertificationId) ||
                (!!normalizedCertCode && (item.certCode || '').toString().trim().toLowerCase() === normalizedCertCode.toLowerCase()) ||
                (!!normalizedCertName && (item.certName || item.certificateName || '').toString().trim().toLowerCase() === normalizedCertName.toLowerCase())
            );
            targetEnrollmentId = Number(matchedEnrollment?.id || 0);
        }

        if (targetEnrollmentId <= 0) {
            if (!normalizedUserEmail) {
                throw new Error('Learner email is required to create a completed enrollment record.');
            }

            return this.addOrUpdateEnrollment({
                userEmail: normalizedUserEmail,
                userName: normalizedUserName || normalizedUserEmail,
                certificationId: normalizedCertificationId > 0 ? normalizedCertificationId : undefined,
                certCode: normalizedCertCode || normalizedExamCode || normalizedCertName,
                certName: normalizedCertName || normalizedCertCode || normalizedExamCode,
                certificateName: normalizedCertName || normalizedCertCode || normalizedExamCode,
                startDate: normalizedExamDateIso,
                endDate: normalizedExamDateIso,
                examScheduledDate: normalizedExamDateIso,
                assignedDate: normalizedExamDateIso,
                completionDate: normalizedExamDateIso,
                examCode: normalizedExamCode || normalizedCertCode,
                status: 'completed',
                progress: 100,
                pathId: normalizedCertCode || normalizedCertName || normalizedExamCode
            });
        }

        const digest = await this._getFormDigestValue();
        const payload: Record<string, any> = {
            Title: normalizedCertName || normalizedCertCode || 'Enrollment'
        };

        if (context.schema.statusField) {
            payload[context.schema.statusField] = 'Completed';
        } else {
            payload.Status = 'Completed';
        }

        if (context.schema.completionDateField) {
            payload[context.schema.completionDateField] = normalizedExamDateIso;
        }

        if (context.schema.examScheduledDateField) {
            payload[context.schema.examScheduledDateField] = normalizedExamDateIso;
        }

        if (context.schema.endDateField) {
            payload[context.schema.endDateField] = normalizedExamDateIso;
        }

        if (context.schema.progressField) {
            payload[context.schema.progressField] = 100;
        }

        if (context.schema.examCodeField && normalizedExamCode) {
            payload[context.schema.examCodeField] = normalizedExamCode;
        }

        const response = await this._getHttpClient().post(
            `${context.siteUrl}/_api/web/lists/getbytitle('${context.escapedListName}')/items(${targetEnrollmentId})`,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'MERGE',
                    'X-RequestDigest': digest
                }),
                body: JSON.stringify(payload)
            }
        );

        if (!response.ok) {
            const errorText = await this._readErrorBody(response);
            throw new Error(`Failed to update enrollment completion (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 400) || 'No error details returned.'}`);
        }

        return targetEnrollmentId;
    }

    public static async addOrUpdateEnrollment(
        enrollment: IEnrollment,
        options: { failIfExists?: boolean; minimalFieldsOnly?: boolean; skipCertificationCountSync?: boolean } = {}
    ): Promise<number> {
        const context = await this._getEnrollmentListContext();
        await this._ensureList(context.listName);
        await this._ensureEnrollmentCompletionFields(context.listName);
        const refreshedContext = await this._getEnrollmentListContext();

        const normalizedRequestStatus = (enrollment.status || '').toLowerCase();
        const storedStatus =
            normalizedRequestStatus === 'completed' ? 'Completed' :
                normalizedRequestStatus === 'rescheduled' ? 'Rescheduled' :
                    normalizedRequestStatus === 'assigned' ? 'Assigned' :
                        normalizedRequestStatus === 'scheduled' ? 'Scheduled' :
                            (enrollment.assignedByAdmin ? 'Scheduled' : (enrollment.status || 'Assigned'));

        const normalizeIsoValue = (value?: string): string => {
            const rawValue = (value || '').toString().trim();
            if (!rawValue) {
                return '';
            }

            const parsedValue = new Date(rawValue);
            return Number.isNaN(parsedValue.getTime()) ? rawValue : parsedValue.toISOString();
        };

        const assignedDate = normalizeIsoValue(enrollment.assignedDate || enrollment.startDate || new Date().toISOString());
        const examScheduledDate = normalizeIsoValue(enrollment.examScheduledDate || '');
        const rescheduledDate = normalizeIsoValue(enrollment.rescheduledDate || '');
        const startDate = normalizeIsoValue(enrollment.startDate || assignedDate);
        const endDate = normalizeIsoValue(enrollment.endDate || examScheduledDate || rescheduledDate || '');
        const expiryDate = normalizeIsoValue(enrollment.expiryDate || endDate || examScheduledDate || rescheduledDate || '');
        const completionDate = normalizeIsoValue(
            enrollment.completionDate ||
            (normalizedRequestStatus === 'completed' ? (enrollment.examScheduledDate || enrollment.endDate || enrollment.startDate || assignedDate) : '')
        );
        const examCode = (enrollment.examCode || '').toString().trim();
        const assignedByName = enrollment.assignedByName || this._getCurrentContextUserName();
        const assignedById = enrollment.assignedById || await this._getCurrentContextUserId();
        const assignedToId = await this._getSiteUserIdByEmail(enrollment.userEmail);
        const certificationLookupId = Number(enrollment.certificationId || 0);
        const normalizedUserEmail = (enrollment.userEmail || '').toString().trim().toLowerCase();
        const normalizedCertName = (enrollment.certName || '').toString().trim();
        const normalizedCertCode = (enrollment.certCode || '').toString().trim();
        const enrollmentPathId = this._getEnrollmentPathId(enrollment) || (enrollment.certCode || enrollment.certName || '').toString().trim();
        const existingItems: IEnrollment[] = await this.getEnrollments(enrollment.userEmail).catch(() => [] as IEnrollment[]);
        const existingItem = existingItems.find((item) =>
            (item.userEmail || '').toLowerCase() === normalizedUserEmail &&
            (
                (certificationLookupId > 0 && Number(item.certificationId || 0) === certificationLookupId) ||
                (!!normalizedCertCode && (item.certCode || '').toLowerCase() === normalizedCertCode.toLowerCase()) ||
                (!!normalizedCertName && (item.certName || '').toLowerCase() === normalizedCertName.toLowerCase())
            )
        );
        const requestedItemId = Number(enrollment.id || 0);

        if (existingItem?.id && options.failIfExists) {
            throw new Error('Already enrolled');
        }

        const itemId = requestedItemId > 0 ? requestedItemId : (existingItem?.id || 0);

        if (options.failIfExists && itemId <= 0) {
            const duplicateExists = certificationLookupId > 0
                ? await this.hasEnrollmentForUserCertificationId(normalizedUserEmail, certificationLookupId, normalizedCertName, normalizedCertCode)
                : await this.hasEnrollmentForUserCertification(normalizedUserEmail, normalizedCertName, normalizedCertCode);

            if (duplicateExists) {
                throw new Error('Already enrolled');
            }
        }

        const digest = await this._getFormDigestValue();

        const saveEnrollment = async (includePersonFields: boolean): Promise<any> => {
            const getFieldType = (fieldName?: string | null): string =>
                fieldName ? (refreshedContext.schema.fieldTypes?.[fieldName] || '') : '';
            const setStringField = (fieldName: string | null | undefined, value: any, label: string): void => {
                if (!fieldName || value === undefined || value === null || value === '') {
                    return;
                }

                const fieldType = getFieldType(fieldName);
                if (this._isPersonOrLookupFieldType(fieldType)) {
                    console.warn(`[Enrollments] Skipping ${label}; SharePoint resolved '${fieldName}' as a person/lookup field.`, { fieldName, fieldType });
                    return;
                }

                if (this._isNumericFieldType(fieldType)) {
                    const numericValue = Number(value);
                    if (Number.isFinite(numericValue)) {
                        payload[fieldName] = numericValue;
                    } else {
                        console.warn(`[Enrollments] Skipping ${label}; SharePoint resolved '${fieldName}' as numeric but value is non-numeric.`, {
                            fieldName,
                            fieldType,
                            value
                        });
                    }
                    return;
                }

                payload[fieldName] = String(value);
            };
            const setNumericField = (fieldName: string | null | undefined, value: any, label: string): void => {
                if (!fieldName || value === undefined || value === null || value === '') {
                    return;
                }

                const numericValue = Number(value);
                if (!Number.isFinite(numericValue)) {
                    console.warn(`[Enrollments] Skipping ${label}; value is not numeric.`, {
                        fieldName,
                        value
                    });
                    return;
                }

                const fieldType = getFieldType(fieldName);
                if (this._isPersonOrLookupFieldType(fieldType)) {
                    console.warn(`[Enrollments] Skipping ${label}; SharePoint resolved '${fieldName}' as a person/lookup field.`, { fieldName, fieldType });
                    return;
                }

                payload[fieldName] = numericValue;
            };

            const payload: Record<string, any> = {
                Title: String(enrollment.certName || 'Enrollment')
            };

            setStringField(refreshedContext.schema.userEmailField, enrollment.userEmail, 'userEmail');
            setStringField(refreshedContext.schema.userNameField, enrollment.userName, 'userName');
            setStringField(refreshedContext.schema.certNameField, enrollment.certName, 'certName');
            setStringField(refreshedContext.schema.assignedDateField, assignedDate, 'assignedDate');
            setStringField(refreshedContext.schema.examScheduledDateField, examScheduledDate, 'examScheduledDate');

            if (!options.minimalFieldsOnly) {
                setStringField(refreshedContext.schema.certCodeField, enrollment.certCode, 'certCode');
                setStringField(refreshedContext.schema.startDateField, startDate, 'startDate');
                setStringField(refreshedContext.schema.endDateField, endDate, 'endDate');
                setStringField(refreshedContext.schema.expiryDateField, expiryDate, 'expiryDate');
                setStringField(refreshedContext.schema.statusField, storedStatus, 'status');
                setNumericField(refreshedContext.schema.progressField, typeof enrollment.progress === 'number' ? enrollment.progress : Number(enrollment.progress || 0), 'progress');
                setStringField(refreshedContext.schema.certificateNameField, enrollment.certificateName || enrollment.certName, 'certificateName');
                setStringField(refreshedContext.schema.pathIdField, enrollmentPathId, 'pathId');
                setStringField(refreshedContext.schema.rescheduledDateField, rescheduledDate, 'rescheduledDate');
                setStringField(refreshedContext.schema.assignedByNameField, assignedByName, 'assignedByName');
                setStringField(refreshedContext.schema.completionDateField, completionDate, 'completionDate');
                setStringField(refreshedContext.schema.examCodeField, examCode, 'examCode');
            }

            if (enrollment.assignedByAdmin) {
                setStringField(refreshedContext.schema.assignedToEmailField, enrollment.userEmail, 'assignedToEmail');
            }

            if (certificationLookupId > 0 && refreshedContext.schema.certificationLookupField) {
                payload[`${refreshedContext.schema.certificationLookupField}Id`] = certificationLookupId;
            }

            if (includePersonFields) {
                if (assignedToId && refreshedContext.schema.assignedToField) {
                    payload[`${refreshedContext.schema.assignedToField}Id`] = Number(assignedToId);
                }

                if (assignedById && refreshedContext.schema.assignedByField) {
                    payload[`${refreshedContext.schema.assignedByField}Id`] = Number(assignedById);
                }
            }

            Object.keys(payload).forEach((key) => {
                if (payload[key] === undefined || payload[key] === null || payload[key] === '') {
                    delete payload[key];
                }
            });

            console.log('[Enrollments] Final SharePoint payload:', JSON.stringify(payload, null, 2));

            return this._getHttpClient().post(
                itemId
                    ? `${refreshedContext.siteUrl}/_api/web/lists/getbytitle('${refreshedContext.escapedListName}')/items(${itemId})`
                    : `${refreshedContext.siteUrl}/_api/web/lists/getbytitle('${refreshedContext.escapedListName}')/items`,
                SPHttpClient.configurations.v1,
                {
                    headers: this._getJsonHeaders(itemId ? {
                        'IF-MATCH': '*',
                        'X-HTTP-Method': 'MERGE',
                        'X-RequestDigest': digest
                    } : {
                        'X-RequestDigest': digest
                    }),
                    body: JSON.stringify(payload)
                }
            );
        };

        let response = await saveEnrollment(!!(assignedToId || assignedById));
        if (!response.ok && (assignedToId || assignedById)) {
            const errorText = await this._readErrorBody(response);
            console.warn('[Enrollments] Failed to save person fields. Retrying without person fields.', {
                listName: context.listName,
                userEmail: enrollment.userEmail,
                certCode: enrollment.certCode,
                responsePreview: errorText.substring(0, 300)
            });
            response = await saveEnrollment(false);
        }

        if (!response.ok) {
            const errorText = await this._readErrorBody(response);
            throw new Error(`Failed to save enrollment (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 400) || 'No error details returned.'}`);
        }

        const wasUpdate = itemId > 0;
        let savedItemId = itemId;

        if (!savedItemId) {
            const createdItem = await this._readJson<any>(response);
            savedItemId = createdItem?.Id || createdItem?.id || 0;
        }

        if (enrollment.assignedByAdmin && assignedToId && savedItemId > 0) {
            await this._ensureEnrollmentLearnerEditAccess(savedItemId, assignedToId).catch((error) => {
                console.warn('[Enrollments] Enrollment saved but learner edit access sync failed.', {
                    enrollmentId: savedItemId,
                    assignedToId,
                    userEmail: enrollment.userEmail,
                    error
                });
            });
        }

        if (enrollment.assignedByAdmin && !wasUpdate) {
            await this._syncUserAssignmentNotification(enrollment, storedStatus).catch((error) => {
                console.error('[Assignments] Enrollment saved but user notification sync failed', {
                    userEmail: enrollment.userEmail,
                    certCode: enrollment.certCode,
                    error
                });
            });

            await this.addAuditLogEntry({
                title: 'Enrollment Created',
                learnerEmail: enrollment.userEmail,
                learnerName: enrollment.userName,
                learnerUserId: assignedToId || undefined,
                action: 'Assigned',
                assignmentName: enrollment.certName,
                pathId: enrollmentPathId,
                assignmentDate: examScheduledDate || assignedDate,
                assignedById: assignedById || undefined,
                status: 'Pending'
            }).catch((error) => {
                console.error('[Audit Logs] Enrollment saved but audit log creation failed', error);
            });
        }

        if (!options.skipCertificationCountSync) {
            await this._syncCertificationAssignedLearnerCounts([
                existingItem || null,
                {
                    certificationId: certificationLookupId,
                    certName: normalizedCertName || existingItem?.certName || '',
                    certCode: normalizedCertCode || existingItem?.certCode || enrollmentPathId
                }
            ]);
        }

        this.emitEnrollmentRefreshSignal();
        return savedItemId;
    }

    public static async deleteEnrollment(id: number): Promise<void> {
        const context = await this._getEnrollmentListContext();
        const digest = await this._getFormDigestValue();
        let enrollmentToDelete: IEnrollment | null = null;

        try {
            const detailEndpoint =
                `${context.siteUrl}/_api/web/lists/getbytitle('${context.escapedListName}')/items(${id})` +
                `?$select=${this._getEnrollmentSelectFields(context.schema).join(',')}` +
                `${this._getEnrollmentExpandFields(context.schema).length > 0 ? `&$expand=${this._getEnrollmentExpandFields(context.schema).join(',')}` : ''}`;
            const enrollmentItem = await this._safeGetJson<any>(detailEndpoint, `enrollment ${id}`);
            enrollmentToDelete = this._mapEnrollmentItem(enrollmentItem, context.schema);
        } catch (error) {
            console.warn('[Enrollments] Failed to read enrollment metadata before delete.', {
                listName: context.listName,
                id,
                error
            });
        }

        await this._getHttpClient().post(
            `${context.siteUrl}/_api/web/lists/getbytitle('${context.escapedListName}')/items(${id})`,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'DELETE',
                    'X-RequestDigest': digest
                })
            }
        );

        await this._syncCertificationAssignedLearnerCounts([enrollmentToDelete]);
        this.emitEnrollmentRefreshSignal();
    }

    public static async deleteEnrollmentByEmailAndCode(email: string, code: string): Promise<void> {
        const normalizedEmail = (email || '').toString().trim().toLowerCase();
        const normalizedCode = (code || '').toString().trim().toLowerCase();
        const enrollments = await this.getEnrollments(normalizedEmail);
        const matchedEnrollments = enrollments.filter((item) =>
            (item.userEmail || '').toLowerCase() === normalizedEmail &&
            (item.certCode || item.certName || '').toLowerCase() === normalizedCode
        );

        for (const item of matchedEnrollments) {
            if (item.id) {
                await this.deleteEnrollment(item.id);
            }
        }
    }

    public static async rescheduleEnrollment(id: number, newExamDate: string, newEndDate?: string): Promise<void> {
        if (!id) {
            throw new Error('Enrollment id is required to reschedule.');
        }

        const context = await this._getEnrollmentListContext();
        const normalizedExamDate = new Date(newExamDate);
        const normalizedEndDate = new Date(newEndDate || newExamDate);

        if (Number.isNaN(normalizedExamDate.getTime())) {
            throw new Error('A valid exam date is required.');
        }

        if (Number.isNaN(normalizedEndDate.getTime())) {
            throw new Error('A valid end date is required.');
        }

        const detailEndpoint =
            `${context.siteUrl}/_api/web/lists/getbytitle('${context.escapedListName}')/items(${id})` +
            `?$select=${this._getEnrollmentSelectFields(context.schema).join(',')}` +
            `${this._getEnrollmentExpandFields(context.schema).length > 0 ? `&$expand=${this._getEnrollmentExpandFields(context.schema).join(',')}` : ''}`;
        const enrollmentItem = await this._safeGetJson<any>(detailEndpoint, `enrollment ${id}`);
        const currentEnrollment = this._mapEnrollmentItem(enrollmentItem, context.schema);

        if (!currentEnrollment.userEmail || !currentEnrollment.certName) {
            throw new Error('Failed to resolve the enrollment details before rescheduling.');
        }

        await this.addOrUpdateEnrollment(
            {
                ...currentEnrollment,
                id,
                endDate: newEndDate || newExamDate,
                examScheduledDate: newExamDate,
                rescheduledDate: newExamDate,
                status: 'rescheduled'
            },
            {
                minimalFieldsOnly: false
            }
        );
    }

    public static async getLearners(forceRefresh: boolean = false): Promise<ILearnerDirectoryUser[]> {
        const membershipSnapshot = await this.getDefaultSiteGroupMembership(forceRefresh);
        console.log('[LearnerSync] Fetched learners from SharePoint site groups', {
            owners: membershipSnapshot.owners.length,
            members: membershipSnapshot.members.length,
            visitors: membershipSnapshot.visitors.length,
            learners: membershipSnapshot.learners.length
        });
        return membershipSnapshot.learners;
    }

    public static async getAllSiteLearners(forceRefresh: boolean = false): Promise<ILearnerDirectoryUser[]> {
        const membershipSnapshot = await this.getDefaultSiteGroupMembership(forceRefresh);
        const siteUsers = await this._fetchAllSiteUsers();
        return this._enrichDirectoryUsers(this._mergeMembershipGroups([membershipSnapshot.learners, siteUsers]));
    }

    public static async getAssessmentAssignmentLearners(): Promise<ILearnerDirectoryUser[]> {
        return this.getAllSiteLearners();
    }

    public static async getLearnerDirectoryUsers(forceRefresh: boolean = false): Promise<ILearnerDirectoryUser[]> {
        for (const groupName of this._learnerGroupCandidates) {
            const groupUsers = await this._fetchNamedSiteGroupUsers(groupName);
            if (groupUsers.length > 0) {
                return this._enrichDirectoryUsers(groupUsers);
            }
        }

        return this.getAllSiteLearners(forceRefresh);
    }

    public static async getLearnerDirectoryCount(forceRefresh: boolean = false): Promise<number> {
        const learners = await this.getLearnerDirectoryUsers(forceRefresh).catch(() => [] as ILearnerDirectoryUser[]);
        const uniqueLearners = new Map<string, ILearnerDirectoryUser>();

        learners.forEach((learner, index) => {
            const dedupeKey = (
                (learner?.email || learner?.Email || learner?.login || learner?.LoginName || '').toString().trim().toLowerCase() ||
                `learner-${index}`
            );
            if (!uniqueLearners.has(dedupeKey)) {
                uniqueLearners.set(dedupeKey, learner);
            }
        });

        return uniqueLearners.size;
    }

    private static _getCertificationAssignmentNameFromItem(item: any, schema: ICertificationListSchema): string {
        return (
            this._readFieldValue(item, schema.certificationNameField) ||
            item?.Title ||
            ''
        ).toString().trim();
    }

    private static _getCertificationAssignedUser(item: any, schema: ICertificationListSchema): any {
        const assignedTo = this._readFieldValue(item, schema.assignedToField);
        return assignedTo && typeof assignedTo === 'object' ? assignedTo : null;
    }

    private static _getCertificationAssignedUserEmail(item: any, schema: ICertificationListSchema): string {
        const assignedUser = this._getCertificationAssignedUser(item, schema);
        return (
            assignedUser?.EMail ||
            assignedUser?.Email ||
            this._readFieldValue(item, schema.userEmailField) ||
            ''
        ).toString().trim();
    }

    private static _getCertificationAssignedUserName(item: any, schema: ICertificationListSchema): string {
        const assignedUser = this._getCertificationAssignedUser(item, schema);
        return (
            assignedUser?.Title ||
            this._readFieldValue(item, schema.userNameField) ||
            this._getCertificationAssignedUserEmail(item, schema)
        ).toString().trim();
    }

    private static _getAssessmentAssignedUser(item: any, schema: IAssessmentAssignmentListSchema): any {
        const assignedTo = this._readFieldValue(item, schema.assignedToField);
        return assignedTo && typeof assignedTo === 'object' ? assignedTo : null;
    }

    private static _getAssessmentAssignedUserEmail(item: any, schema: IAssessmentAssignmentListSchema): string {
        const assignedUser = this._getAssessmentAssignedUser(item, schema);
        return (
            assignedUser?.EMail ||
            assignedUser?.Email ||
            this._readFieldValue(item, schema.userEmailField) ||
            ''
        ).toString().trim();
    }

    private static _getAssessmentAssignedUserName(item: any, schema: IAssessmentAssignmentListSchema): string {
        const assignedUser = this._getAssessmentAssignedUser(item, schema);
        return (
            assignedUser?.Title ||
            this._readFieldValue(item, schema.userNameField) ||
            this._getAssessmentAssignedUserEmail(item, schema)
        ).toString().trim();
    }

    private static _mapAssessmentAssignmentListItem(item: any, schema: IAssessmentAssignmentListSchema): IAssessmentAssignmentRecord {
        return {
            id: Number(item?.Id || item?.id || 0),
            title: (item?.Title || '').toString().trim(),
            userEmail: this._getAssessmentAssignedUserEmail(item, schema),
            userName: this._getAssessmentAssignedUserName(item, schema),
            assessmentName: (
                this._readFieldValue(item, schema.assessmentNameField) ||
                item?.Title ||
                ''
            ).toString().trim(),
            orderIndex: Number(this._readFieldValue(item, schema.orderIndexField) || 0),
            assignedGroup: ((this._readFieldValue(item, schema.assignedGroupField) || 'Members').toString().trim() as 'Owners' | 'Members' | 'Visitors'),
            scheduledDate: (this._readFieldValue(item, schema.scheduledDateField) || '').toString(),
            created: (item?.Created || '').toString(),
            assessmentPayload: this._parseAssessmentPayload((this._readFieldValue(item, schema.assessmentPayloadField) || '').toString())
        };
    }

    private static async _addAuditLogEntries(entries: Array<{
        title: string;
        learnerEmail: string;
        learnerName?: string;
        learnerUserId?: number;
        action: string;
        assignmentName?: string;
        pathId?: string;
        assignmentDate?: string;
        assignedById?: number;
        status?: string;
    }>): Promise<void> {
        const normalizedEntries = (entries || [])
            .map((entry) => ({
                title: (entry?.title || '').toString().trim(),
                learnerEmail: (entry?.learnerEmail || '').toString().trim().toLowerCase(),
                learnerName: (entry?.learnerName || '').toString().trim(),
                learnerUserId: entry?.learnerUserId ? Number(entry.learnerUserId) : undefined,
                action: (entry?.action || '').toString().trim(),
                assignmentName: (entry?.assignmentName || '').toString().trim(),
                pathId: (entry?.pathId || entry?.assignmentName || '').toString().trim(),
                assignmentDate: new Date(entry?.assignmentDate || new Date().toISOString()).toISOString(),
                assignedById: entry?.assignedById ? Number(entry.assignedById) : undefined,
                status: (entry?.status || '').toString().trim()
            }))
            .filter((entry) => entry.title && entry.learnerEmail && entry.action);

        if (normalizedEntries.length === 0) {
            return;
        }

        const siteUrl = this._ensureProductionSiteUrl();
        const listName = await this._resolveAuditLogListName();
        const schema = await this._getAuditLogListSchema();
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items`;
        const digest = await this._getFormDigestValue();

        await this._processInChunks(normalizedEntries, 20, async (chunk) => {
            await Promise.all(
                chunk.map(async (entry) => {
                    const fullPayload: Record<string, string | number> = {
                        Title: entry.title
                    };

                    if (schema.learnerEmailField) {
                        fullPayload[schema.learnerEmailField] = entry.learnerEmail;
                    } else {
                        fullPayload.LearnerEmail = entry.learnerEmail;
                    }

                    if (schema.actionField) {
                        fullPayload[schema.actionField] = entry.action;
                    } else {
                        fullPayload.Action = entry.action;
                    }

                    if (schema.assignmentDateField) {
                        fullPayload[schema.assignmentDateField] = entry.assignmentDate;
                    } else {
                        fullPayload.AssignmentDate = entry.assignmentDate;
                    }

                    if (schema.timestampField && schema.timestampField !== schema.assignmentDateField) {
                        fullPayload[schema.timestampField] = entry.assignmentDate;
                    }

                    if (entry.learnerName) {
                        if (schema.learnerNameField) {
                            fullPayload[schema.learnerNameField] = entry.learnerName;
                        } else {
                            fullPayload.LearnerName = entry.learnerName;
                        }
                    }

                    if (entry.learnerUserId && schema.userField) {
                        fullPayload[`${schema.userField}Id`] = entry.learnerUserId;
                    }

                    if (entry.assignmentName) {
                        if (schema.assignmentNameField) {
                            fullPayload[schema.assignmentNameField] = entry.assignmentName;
                        } else {
                            fullPayload.AssignmentName = entry.assignmentName;
                        }
                    }

                    if (entry.pathId && schema.pathIdField) {
                        fullPayload[schema.pathIdField] = entry.pathId;
                    }

                    if (entry.assignedById && schema.assignedByField) {
                        fullPayload[`${schema.assignedByField}Id`] = entry.assignedById;
                    } else if (entry.assignedById) {
                        fullPayload.AssignedById = entry.assignedById;
                    }

                    if (entry.status) {
                        if (schema.statusField) {
                            fullPayload[schema.statusField] = entry.status;
                        } else {
                            fullPayload.Status = entry.status;
                        }
                    }

                    console.log('Audit Payload:', JSON.stringify(fullPayload, null, 2));

                    try {
                        await this._safePostJson<any>(
                            endpoint,
                            {
                                headers: this._getJsonHeaders({
                                    'X-RequestDigest': digest
                                }),
                                body: JSON.stringify(fullPayload)
                            },
                            `audit log for ${entry.learnerEmail}`
                        );
                    } catch (error) {
                        const fallbackPayload: Record<string, string | number> = {
                            Title: entry.title,
                            LearnerEmail: entry.learnerEmail,
                            Action: entry.action,
                            AssignmentDate: entry.assignmentDate
                        };

                        console.warn('[Audit Logs] Retrying with core audit fields only', {
                            learnerEmail: entry.learnerEmail,
                            error
                        });
                        console.log('Audit Payload:', JSON.stringify(fallbackPayload, null, 2));

                        await this._safePostJson<any>(
                            endpoint,
                            {
                                headers: this._getJsonHeaders({
                                    'X-RequestDigest': digest
                                }),
                                body: JSON.stringify(fallbackPayload)
                            },
                            `audit log fallback for ${entry.learnerEmail}`
                        );
                    }
                })
            );
        }, normalizedEntries.length > 20 ? 150 : 0);
    }

    private static async _getAuditLogById(id: number): Promise<IAuditLogRecord | null> {
        if (!id) {
            return null;
        }

        const siteUrl = this._ensureProductionSiteUrl();
        const listName = await this._resolveAuditLogListName();
        const schema = await this._getAuditLogListSchema();
        const selectFields = Array.from(new Set([
            'Id',
            'Title',
            'Created',
            schema.learnerEmailField || 'LearnerEmail',
            schema.learnerNameField || 'LearnerName',
            schema.actionField || 'Action',
            schema.assignmentNameField || 'AssignmentName',
            schema.assignmentDateField || 'AssignmentDate',
            schema.statusField || 'Status',
            schema.pathIdField || 'PathId',
            schema.timestampField || 'Timestamp',
            schema.userField ? `${schema.userField}/Id` : '',
            schema.userField ? `${schema.userField}/Title` : '',
            schema.userField ? `${schema.userField}/EMail` : '',
            schema.assignedByField ? `${schema.assignedByField}/Id` : ''
        ].filter((field) => !!field)));
        const expandFields = Array.from(new Set([
            schema.userField || '',
            schema.assignedByField || ''
        ].filter((field) => !!field)));
        const endpoint =
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items(${id})` +
            `?$select=${selectFields.join(',')}` +
            `${expandFields.length > 0 ? `&$expand=${expandFields.join(',')}` : ''}`;
        const data = await this._safeGetJson<any>(endpoint, `audit log ${id}`);
        return data ? this._mapAuditLogItem(data, schema) : null;
    }

    public static async addAuditLogEntry(entry: {
        title: string;
        learnerEmail: string;
        learnerName?: string;
        learnerUserId?: number;
        action: string;
        assignmentName?: string;
        pathId?: string;
        assignmentDate?: string;
        assignedById?: number;
        status?: string;
    }): Promise<void> {
        await this._addAuditLogEntries([entry]);

        if ((entry.action || '').toString().trim().toLowerCase() === 'deleted') {
            await this.deleteEnrollmentByAudit({
                id: 0,
                title: entry.title,
                learnerEmail: entry.learnerEmail,
                learnerName: entry.learnerName,
                action: entry.action,
                assignmentName: entry.assignmentName,
                assignmentDate: entry.assignmentDate || new Date().toISOString(),
                assignedById: entry.assignedById,
                status: entry.status,
                created: new Date().toISOString(),
                pathId: entry.pathId || entry.assignmentName,
                userId: entry.learnerUserId,
                timestamp: entry.assignmentDate || new Date().toISOString()
            });
        }
    }

    public static async getAuditLogs(learnerEmail?: string): Promise<IAuditLogRecord[]> {
        const siteUrl = this._ensureProductionSiteUrl();
        const listName = await this._resolveAuditLogListName();
        const schema = await this._getAuditLogListSchema();
        const normalizedLearnerEmail = (learnerEmail || '').toString().trim().toLowerCase();
        const selectFields = Array.from(new Set([
            'Id',
            'Title',
            'Created',
            schema.learnerEmailField || '',
            schema.learnerNameField || '',
            schema.actionField || '',
            schema.assignmentNameField || '',
            schema.assignmentDateField || '',
            schema.statusField || '',
            schema.pathIdField || '',
            schema.timestampField || '',
            schema.userField ? `${schema.userField}/Id` : '',
            schema.userField ? `${schema.userField}/Title` : '',
            schema.userField ? `${schema.userField}/EMail` : '',
            schema.assignedByField ? `${schema.assignedByField}/Id` : ''
        ].filter((field) => !!field)));
        const expandFields = Array.from(new Set([
            schema.userField || '',
            schema.assignedByField || ''
        ].filter((field) => !!field)));
        const primaryEndpoint =
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items?` +
            `$select=${selectFields.join(',')}` +
            `${expandFields.length > 0 ? `&$expand=${expandFields.join(',')}` : ''}` +
            `&$orderby=Created desc&$top=100`;
        const fallbackEndpoint =
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items?` +
            `$orderby=Created desc&$top=100`;

        let data: any = null;
        try {
            data = await this._safeGetJson<any>(primaryEndpoint, 'audit logs');
        } catch (error) {
            console.error('Audit Logs Error:', error);
            data = await this._safeGetJson<any>(fallbackEndpoint, 'audit logs fallback');
        }

        const logs = this._toCollection(data)
            .map((item: any): IAuditLogRecord => this._mapAuditLogItem(item, schema))
            .filter((log: IAuditLogRecord) => !normalizedLearnerEmail || (log.learnerEmail || '').toLowerCase() === normalizedLearnerEmail);

        console.log('Fetched Logs:', logs);
        return logs;
    }

    public static async getDeletedAuditLogsForUser(userEmail?: string, userId?: number): Promise<IAuditLogRecord[]> {
        const siteUrl = this._ensureProductionSiteUrl();
        const listName = await this._resolveAuditLogListName();
        const schema = await this._getAuditLogListSchema();
        const normalizedLearnerEmail = (userEmail || this.getCurrentContextUserEmail()).toString().trim().toLowerCase();
        const resolvedUserId = userId || this.getCurrentContextUserId() || await this._getSiteUserIdByEmail(normalizedLearnerEmail);
        const actionField = schema.actionField || '';
        const selectFields = Array.from(new Set([
            'Id',
            'Title',
            'Created',
            schema.learnerEmailField || '',
            schema.learnerNameField || '',
            actionField,
            schema.assignmentNameField || '',
            schema.assignmentDateField || '',
            schema.statusField || '',
            schema.pathIdField || '',
            schema.timestampField || '',
            schema.userField ? `${schema.userField}/Id` : '',
            schema.userField ? `${schema.userField}/Title` : '',
            schema.userField ? `${schema.userField}/EMail` : ''
        ].filter((field) => !!field)));
        const expandFields = Array.from(new Set([
            schema.userField || ''
        ].filter((field) => !!field)));
        const filters: string[] = [];

        if (schema.userField && resolvedUserId) {
            filters.push(`${schema.userField}Id eq ${Number(resolvedUserId)}`);
        } else if (schema.learnerEmailField && normalizedLearnerEmail) {
            filters.push(`${schema.learnerEmailField} eq '${this._escapeODataValue(normalizedLearnerEmail)}'`);
        }

        if (actionField) {
            filters.push(`${actionField} eq 'Deleted'`);
        } else {
            return this.getAuditLogs(normalizedLearnerEmail).then((logs) =>
                logs.filter((log) => {
                    const emailMatch = !normalizedLearnerEmail || (log.learnerEmail || '').toLowerCase() === normalizedLearnerEmail;
                    const userMatch = !resolvedUserId || !log.userId || log.userId === resolvedUserId;
                    return emailMatch && userMatch && (log.action || '').toLowerCase() === 'deleted';
                })
            );
        }

        const endpoint =
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items?` +
            `$select=${selectFields.join(',')}` +
            `${expandFields.length > 0 ? `&$expand=${expandFields.join(',')}` : ''}` +
            `${filters.length > 0 ? `&$filter=${filters.join(' and ')}` : ''}` +
            `&$orderby=Created desc&$top=100`;

        let logs: IAuditLogRecord[] = [];

        try {
            const data = await this._safeGetJson<any>(endpoint, 'deleted audit logs');
            logs = this._toCollection(data).map((item: any) => this._mapAuditLogItem(item, schema));
        } catch (error) {
            console.error('[Audit Logs] Failed to fetch deleted audit logs with filtered query', {
                endpoint,
                error
            });
            logs = await this.getAuditLogs(normalizedLearnerEmail);
        }

        return logs.filter((log) => {
            const emailMatch = !normalizedLearnerEmail || (log.learnerEmail || '').toLowerCase() === normalizedLearnerEmail;
            const userMatch = !resolvedUserId || !log.userId || log.userId === resolvedUserId;
            return emailMatch && userMatch && (log.action || '').toLowerCase() === 'deleted';
        });
    }

    public static async deleteEnrollmentByAudit(auditLogOrId: number | IAuditLogRecord): Promise<number> {
        const auditLog = typeof auditLogOrId === 'number'
            ? await this._getAuditLogById(auditLogOrId)
            : auditLogOrId;

        if (!auditLog) {
            return 0;
        }

        const normalizedEmail = (auditLog.learnerEmail || '').toString().trim().toLowerCase();
        const normalizedPathId = (auditLog.pathId || auditLog.assignmentName || '').toString().trim().toLowerCase();

        let enrollments = normalizedEmail
            ? await this.getEnrollments(normalizedEmail)
            : await this.getEnrollments('');

        if (auditLog.userId) {
            enrollments = enrollments.filter((enrollment) => !enrollment.userId || enrollment.userId === auditLog.userId);
        }

        const matchingEnrollments = enrollments.filter((enrollment) => {
            const pathId = this._getEnrollmentPathId(enrollment).toLowerCase();
            if (normalizedPathId) {
                return pathId === normalizedPathId;
            }

            return normalizedEmail ? (enrollment.userEmail || '').toLowerCase() === normalizedEmail : false;
        });

        if (matchingEnrollments.length === 0) {
            return 0;
        }

        await Promise.all(
            matchingEnrollments
                .filter((enrollment) => !!enrollment.id)
                .map((enrollment) => this.deleteEnrollment(Number(enrollment.id)))
        );

        return matchingEnrollments.length;
    }

    public static async syncDeletedAuditLogs(
        userEmail?: string,
        userId?: number,
        existingEnrollments?: IEnrollment[]
    ): Promise<{ deletedLogs: IAuditLogRecord[]; removedEnrollmentCount: number; }> {
        const deletedLogs = await this.getDeletedAuditLogsForUser(userEmail, userId);
        if (deletedLogs.length === 0) {
            return {
                deletedLogs: [],
                removedEnrollmentCount: 0
            };
        }

        const enrollments = existingEnrollments || await this.getEnrollments((userEmail || '').toString().trim().toLowerCase());
        const deletedPathIds = new Set(
            deletedLogs
                .map((log) => (log.pathId || log.assignmentName || '').toString().trim().toLowerCase())
                .filter((value) => !!value)
        );

        const matchedEnrollments = enrollments.filter((enrollment) => {
            const pathId = this._getEnrollmentPathId(enrollment).toLowerCase();
            return !!pathId && deletedPathIds.has(pathId);
        });

        if (matchedEnrollments.length === 0) {
            return {
                deletedLogs,
                removedEnrollmentCount: 0
            };
        }

        await Promise.all(
            matchedEnrollments
                .filter((enrollment) => !!enrollment.id)
                .map((enrollment) => this.deleteEnrollment(Number(enrollment.id)))
        );

        return {
            deletedLogs,
            removedEnrollmentCount: matchedEnrollments.length
        };
    }

    public static async deleteAuditLog(id: number): Promise<void> {
        if (!id) {
            return;
        }

        const auditLog = await this._getAuditLogById(id).catch(() => null);
        if (auditLog) {
            await this.deleteEnrollmentByAudit(auditLog).catch((error) => {
                console.error('[Audit Logs] Failed to sync enrollment delete from audit delete', {
                    auditLogId: id,
                    error
                });
            });
        }

        const siteUrl = this._ensureProductionSiteUrl();
        const listName = await this._resolveAuditLogListName();
        const digest = await this._getFormDigestValue();
        await this._spHttpClient.post(
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items(${id})`,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'DELETE',
                    'X-RequestDigest': digest
                })
            }
        );

        this.emitEnrollmentRefreshSignal();
    }

    public static async getRecentAssessmentAssignments(): Promise<IAssessmentAssignmentRecord[]> {
        const siteUrl = this._ensureProductionSiteUrl();
        const listName = await this._resolveAssessmentAssignmentListName();
        const schema = await this._getAssessmentAssignmentListSchema();
        const selectFields = Array.from(new Set([
            'Id',
            'Title',
            'Created',
            `${schema.assignedToField}/Id`,
            `${schema.assignedToField}/Title`,
            `${schema.assignedToField}/EMail`,
            schema.scheduledDateField || '',
            schema.assessmentNameField || '',
            schema.orderIndexField || '',
            schema.assignedGroupField || '',
            schema.assessmentPayloadField || ''
        ].filter((field) => !!field)));
        const endpoint =
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items` +
            `?$select=${selectFields.join(',')}` +
            `&$expand=${schema.assignedToField}` +
            `&$orderby=Created desc&$top=50`;

        const data = await this._safeGetJson<any>(endpoint, 'recent assessment assignments');
        return this._toCollection(data).map((item: any) => this._mapAssessmentAssignmentListItem(item, schema));
    }

    public static async getAllAssessmentAssignments(): Promise<IAssessmentAssignmentRecord[]> {
        const siteUrl = this._ensureProductionSiteUrl();
        const listName = await this._resolveAssessmentAssignmentListName();
        const schema = await this._getAssessmentAssignmentListSchema();
        const selectFields = Array.from(new Set([
            'Id',
            'Title',
            'Created',
            `${schema.assignedToField}/Id`,
            `${schema.assignedToField}/Title`,
            `${schema.assignedToField}/EMail`,
            schema.scheduledDateField || '',
            schema.assessmentNameField || '',
            schema.orderIndexField || '',
            schema.assignedGroupField || '',
            schema.assessmentPayloadField || ''
        ].filter((field) => !!field)));
        const endpoint =
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items` +
            `?$select=${selectFields.join(',')}` +
            `&$expand=${schema.assignedToField}` +
            `&$orderby=Created desc&$top=5000`;

        const data = await this._safeGetJson<any>(endpoint, 'all assessment assignments');
        return this._toCollection(data).map((item: any) => this._mapAssessmentAssignmentListItem(item, schema));
    }

    public static async getAssessmentTrackerItems(): Promise<IAssessmentTrackerItem[]> {
        const assignments = await this.getRecentAssessmentAssignments();
        return assignments.map((item) => ({
            id: Number(item.id || 0),
            learner: (item.userName || item.userEmail || 'Not Available').toString(),
            learnerEmail: (item.userEmail || '').toString(),
            assessment: (item.title || item.assessmentName || 'Assessment').toString(),
            created: (item.created || '').toString()
        }));
    }

    public static async deleteAssessmentAssignment(id: number): Promise<void> {
        if (!id) {
            return;
        }

        const siteUrl = this._ensureProductionSiteUrl();
        const listName = await this._resolveAssessmentAssignmentListName();
        const digest = await this._getFormDigestValue();
        await this._spHttpClient.post(
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items(${id})`,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'DELETE',
                    'X-RequestDigest': digest
                })
            }
        );
    }

    public static async deleteAssessmentTrackerItem(id: number): Promise<void> {
        await this.deleteAssessmentAssignment(id);
    }

    private static async _assignAssessmentToLearners(
        definition: IAssessmentAssignmentDefinition,
        targetLearners: ILearnerDirectoryUser[],
        selectedDate?: string
    ): Promise<{
        assignedCount: number;
        skippedCount: number;
        totalLearners: number;
    }> {
        const normalizedTitle = (definition.title || '').toString().trim();
        if (!normalizedTitle) {
            throw new Error('Assessment title is required before assigning to learners.');
        }

        const scheduledDate = new Date(selectedDate || new Date().toISOString());
        if (Number.isNaN(scheduledDate.getTime())) {
            throw new Error('Date is required');
        }

        const siteUrl = this._ensureProductionSiteUrl();
        const listName = await this._resolveAssessmentAssignmentListName();
        const schema = await this._getAssessmentAssignmentListSchema();

        const learners = this._mergeMembershipGroups([
            (targetLearners || []).filter((learner): learner is ILearnerDirectoryUser => !!learner)
        ]);

        if (learners.length === 0) {
            return {
                assignedCount: 0,
                skippedCount: 0,
                totalLearners: 0
            };
        }

        const assignmentSelectFields = Array.from(new Set([
            'Id',
            'Title',
            `${schema.assignedToField}/Id`,
            `${schema.assignedToField}/Title`,
            `${schema.assignedToField}/EMail`,
            schema.assessmentNameField || '',
            schema.orderIndexField || ''
        ].filter((field) => !!field)));
        const assignmentsEndpoint =
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items` +
            `?$select=${assignmentSelectFields.join(',')}` +
            `&$expand=${schema.assignedToField}`;

        const assignmentsData = await this._safeGetJson<any>(assignmentsEndpoint, 'assessment assignments');
        const existingAssignments = this._toCollection(assignmentsData);
        const normalizedAssignmentKeys = new Set<string>();
        const nextOrderByUser = new Map<string, number>();

        existingAssignments.forEach((item: any) => {
            const itemEmail = this._getAssessmentAssignedUserEmail(item, schema).toLowerCase();
            if (!itemEmail) {
                return;
            }

            const currentOrder = Number(this._readFieldValue(item, schema.orderIndexField) || 0);
            nextOrderByUser.set(itemEmail, Math.max(nextOrderByUser.get(itemEmail) || 0, currentOrder));
            normalizedAssignmentKeys.add(
                `${itemEmail}::${(item?.Title || '').toString().trim().toLowerCase()}`
            );
        });

        let skippedCount = 0;
        const itemsToCreate: Array<{ userEmail: string; body: Record<string, string | number> }> = [];

        for (const learner of learners) {
            const learnerEmail = (learner.email || learner.Email || '').toString().trim().toLowerCase();
            if (!learnerEmail) {
                skippedCount += 1;
                continue;
            }

            const dedupeKey = `${learnerEmail}::${normalizedTitle.toLowerCase()}`;
            if (normalizedAssignmentKeys.has(dedupeKey)) {
                skippedCount += 1;
                continue;
            }

            const nextOrder = (nextOrderByUser.get(learnerEmail) || 0) + 1;
            nextOrderByUser.set(learnerEmail, nextOrder);
            normalizedAssignmentKeys.add(dedupeKey);

            const resolvedUserId = await this._getSiteUserIdByEmail(learnerEmail) || Number(learner.Id || learner.id || 0);
            if (!resolvedUserId || Number.isNaN(Number(resolvedUserId))) {
                throw new Error('Invalid userId');
            }

            const assignedLookupField = `${schema.assignedToField}Id`;
            const payload: Record<string, string | number | undefined | null> = {
                Title: String(normalizedTitle),
                [assignedLookupField]: Number(resolvedUserId),
                [schema.scheduledDateField]: scheduledDate.toISOString()
            };

            Object.keys(payload).forEach((key) => {
                if (payload[key] === undefined || payload[key] === null) {
                    delete payload[key];
                }
            });

            if (typeof payload.Title !== 'string') {
                throw new Error('Title must be string');
            }

            if (typeof payload[assignedLookupField] !== 'number' || Number.isNaN(payload[assignedLookupField] as number)) {
                throw new Error('UserId must be number');
            }

            const scheduledDateValue = payload[schema.scheduledDateField];
            if (typeof scheduledDateValue !== 'string' || scheduledDateValue.indexOf('T') === -1) {
                throw new Error('Date must be ISO string');
            }

            Object.keys(payload).forEach((key) => {
                const value = payload[key];
                if (typeof value === 'object') {
                    throw new Error(`Invalid payload field '${key}'. Objects are not allowed.`);
                }
            });

            console.log('USER ID:', Number(resolvedUserId));
            console.log('FINAL PAYLOAD:', JSON.stringify(payload, null, 2));
            itemsToCreate.push({
                userEmail: learnerEmail,
                body: payload as Record<string, string | number>
            });
        }

        if (itemsToCreate.length === 0) {
            return {
                assignedCount: 0,
                skippedCount,
                totalLearners: learners.length
            };
        }

        const digest = await this._getFormDigestValue();
        const createEndpoint = `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items`;
        const chunkDelayMs = itemsToCreate.length > 20 ? 150 : 0;

        await this._processInChunks(itemsToCreate, 20, async (chunk) => {
            await Promise.all(
                chunk.map(async ({ userEmail, body }) => {
                    try {
                        await this._safePostJson<any>(
                            createEndpoint,
                            {
                                headers: this._getJsonHeaders({
                                    'X-RequestDigest': digest
                                }),
                                body: JSON.stringify(body)
                            },
                            `assessment assignment for ${userEmail}`
                        );
                    } catch (error) {
                        console.error('ASSIGNMENT FAILED:', error);
                        throw error;
                    }
                })
            );
        }, chunkDelayMs);

        await this._addAuditLogEntries(
            itemsToCreate.map(({ userEmail }) => ({
                title: 'Assignment Created',
                learnerEmail: userEmail,
                action: 'Assigned Assessment',
                assignmentName: normalizedTitle,
                assignmentDate: new Date().toISOString(),
                status: 'Pending'
            }))
        ).catch((error) => {
            console.error('[Audit Logs] Failed to create assignment activity records', error);
        });

        return {
            assignedCount: itemsToCreate.length,
            skippedCount,
            totalLearners: learners.length
        };
    }

    public static async assignAssessmentToSelectedLearners(
        definition: IAssessmentAssignmentDefinition,
        selectedLearners: ILearnerDirectoryUser[],
        selectedDate?: string
    ): Promise<{
        assignedCount: number;
        skippedCount: number;
        totalLearners: number;
    }> {
        return this._assignAssessmentToLearners(definition, selectedLearners, selectedDate);
    }

    public static async getAssessmentAssignmentsForUser(userEmail?: string): Promise<IAssessmentAssignmentRecord[]> {
        const normalizedEmail = (userEmail || this._context?.pageContext?.user?.email || '').toString().trim().toLowerCase();
        if (!normalizedEmail) {
            return [];
        }

        const siteUrl = this._ensureProductionSiteUrl();
        const listName = await this._resolveAssessmentAssignmentListName();
        const schema = await this._getAssessmentAssignmentListSchema();
        const selectFields = Array.from(new Set([
            'Id',
            'Title',
            'Created',
            `${schema.assignedToField}/Id`,
            `${schema.assignedToField}/Title`,
            `${schema.assignedToField}/EMail`,
            schema.userEmailField || '',
            schema.userNameField || '',
            schema.assessmentNameField || '',
            schema.orderIndexField || '',
            schema.assignedGroupField || '',
            schema.assessmentPayloadField || ''
        ].filter((field) => !!field)));
        const orderByField = schema.orderIndexField || schema.scheduledDateField || 'Created';

        const endpoint =
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items` +
            `?$select=${selectFields.join(',')}` +
            `&$expand=${schema.assignedToField}` +
            `&$orderby=${orderByField} asc`;

        const data = await this._safeGetJson<any>(endpoint, 'current user assessment assignments');
        return this._toCollection(data)
            .filter((item: any) => this._getAssessmentAssignedUserEmail(item, schema).toLowerCase() === normalizedEmail)
            .map((item: any) => this._mapAssessmentAssignmentListItem(item, schema))
            .sort((left, right) => left.orderIndex - right.orderIndex);
    }

    public static async assignCertificationToSelectedLearners(
        certificationName: string,
        selectedLearners: ILearnerDirectoryUser[],
        selectedDate?: string
    ): Promise<{
        assignedCount: number;
        skippedCount: number;
        totalLearners: number;
    }> {
        const normalizedCertificationName = (certificationName || '').toString().trim();
        if (!normalizedCertificationName) {
            throw new Error('Certification title is required before assigning to learners.');
        }

        const learners = this._mergeMembershipGroups([
            (selectedLearners || []).filter((learner): learner is ILearnerDirectoryUser => !!learner)
        ]);

        if (learners.length === 0) {
            return {
                assignedCount: 0,
                skippedCount: 0,
                totalLearners: 0
            };
        }

        const scheduledDate = new Date(selectedDate || new Date().toISOString());
        if (Number.isNaN(scheduledDate.getTime())) {
            throw new Error('Invalid scheduled date');
        }

        const certification = await this.getCertificationDetailsByTitle(normalizedCertificationName, true);
        if (!certification) {
            throw new Error('Certification not found');
        }

        const canonicalCertName = (certification.title || normalizedCertificationName).toString().trim();
        const canonicalCertCode = (certification.code || canonicalCertName).toString().trim();
        let skippedCount = 0;
        let assignedCount = 0;
        const adminName = this._getCurrentContextUserName();
        const adminId = await this._getCurrentContextUserId();
        const examScheduledDate = scheduledDate.toISOString();
        const chunkDelayMs = learners.length > 20 ? 150 : 0;

        await this._processInChunks(learners, 20, async (chunk) => {
            await Promise.all(
                chunk.map(async (learner) => {
                    const learnerEmail = (learner.email || learner.Email || '').toString().trim().toLowerCase();
                    const learnerName = (learner.name || learner.Title || learnerEmail).toString().trim();

                    if (!learnerEmail) {
                        skippedCount += 1;
                        return;
                    }

                    const alreadyEnrolled = certification.id
                        ? await this.hasEnrollmentForUserCertificationId(learnerEmail, certification.id, canonicalCertName, canonicalCertCode)
                        : await this.hasEnrollmentForUserCertification(learnerEmail, canonicalCertName, canonicalCertCode);
                    if (alreadyEnrolled) {
                        skippedCount += 1;
                        return;
                    }

                    await this.createEnrollmentForCertificationAssignment({
                        userEmail: learnerEmail,
                        userName: learnerName || learnerEmail,
                        certName: canonicalCertName,
                        certCode: canonicalCertCode,
                        pathId: canonicalCertCode,
                        assignedByName: adminName,
                        assignedById: adminId || undefined,
                        examScheduledDate
                    });

                    assignedCount += 1;
                })
            );
        }, chunkDelayMs);

        return {
            assignedCount,
            skippedCount,
            totalLearners: learners.length
        };
    }

    public static async getCertificationAssignmentsForUser(
        userEmail?: string,
        forceRefresh: boolean = false
    ): Promise<ICertificationAssignmentRecord[]> {
        const normalizedEmail = (userEmail || this._context?.pageContext?.user?.email || '').toString().trim().toLowerCase();
        if (!normalizedEmail) {
            return [];
        }

        const siteUrl = this._ensureProductionSiteUrl();
        const listName = this._certificationAssignmentListName;
        await this._ensureList(listName);
        const schema = await this._getCertificationListSchema();
        const selectFields = Array.from(new Set([
            'Id',
            'Title',
            'Created',
            `${schema.assignedToField}/Id`,
            `${schema.assignedToField}/Title`,
            `${schema.assignedToField}/EMail`,
            schema.userEmailField || '',
            schema.userNameField || '',
            schema.certificationNameField || '',
            schema.certCodeField || '',
            schema.scheduledDateField,
            schema.assignedDateField || '',
            schema.issuedDateField || '',
            schema.expiryDateField || '',
            schema.statusField || '',
            schema.orderIndexField || '',
            schema.assignedGroupField || ''
        ].filter((field) => !!field)));
        const orderByField = schema.orderIndexField || schema.scheduledDateField || 'Created';

        const endpoint =
            `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items` +
            `?$select=${selectFields.join(',')}` +
            `&$expand=${schema.assignedToField}` +
            `&$orderby=${orderByField} asc` +
            `&$top=5000` +
            `${forceRefresh ? `&_=${Date.now()}` : ''}`;
        const data = await this._safeGetJson<any>(endpoint, 'current user certifications');

        return this._toCollection(data)
            .filter((item: any) => {
                const itemEmail = this._getCertificationAssignedUserEmail(item, schema).toLowerCase();
                return itemEmail === normalizedEmail;
            })
            .map((item: any) => {
                const assignedDate = (
                    this._readFieldValue(item, schema.assignedDateField) ||
                    this._readFieldValue(item, schema.scheduledDateField) ||
                    this._readFieldValue(item, schema.issuedDateField) ||
                    item?.Created ||
                    ''
                ).toString();

                return {
                    id: Number(item?.Id || item?.id || 0),
                    title: (item?.Title || '').toString().trim(),
                    userEmail: this._getCertificationAssignedUserEmail(item, schema),
                    userName: this._getCertificationAssignedUserName(item, schema),
                    certificationName: this._getCertificationAssignmentNameFromItem(item, schema),
                    certCode: (this._readFieldValue(item, schema.certCodeField) || '').toString().trim(),
                    assignedDate,
                    issuedDate: assignedDate,
                    expiryDate: (this._readFieldValue(item, schema.expiryDateField) || '').toString(),
                    status: (this._readFieldValue(item, schema.statusField) || '').toString().trim(),
                    orderIndex: Number(this._readFieldValue(item, schema.orderIndexField) || 0),
                    assignedGroup: ((this._readFieldValue(item, schema.assignedGroupField) || 'Members').toString().trim() as 'Owners' | 'Members' | 'Visitors'),
                    created: (item?.Created || '').toString()
                };
            })
            .sort((left, right) => left.orderIndex - right.orderIndex);
    }

    public static async fetchNotifications(userEmail: string): Promise<INotification[]> {
        return this.getNotifications(userEmail);
    }

    public static async saveAssignment(enrollment: IEnrollment): Promise<number> {
        return this.addOrUpdateEnrollment(enrollment);
    }

    public static async getUserAssignmentNotifications(userEmail: string): Promise<INotification[]> {
        const trimmedEmail = (userEmail || '').trim().toLowerCase();
        if (!trimmedEmail) {
            return [];
        }

        const siteUrl = this._getSiteUrl();
        const notificationListName = await this._resolveUserNotificationListName();

        if (notificationListName === 'Notifications') {
            const notificationEndpoint =
                `${siteUrl}/_api/web/lists/getbytitle('Notifications')/items` +
                `?$filter=UserEmail eq '${this._escapeODataValue(trimmedEmail)}'` +
                `&$orderby=Created desc`;

            try {
                const data = await this._safeGetJson<any>(
                    notificationEndpoint,
                    'user notification assignments'
                );
                return this._toCollection(data).map((item: any) => ({
                    id: item.Id,
                    title: item.Description || item.Title || 'Certification Assignment',
                    text: item.Title === 'New Certification Assigned' ? 'New certification assigned by Admin' : (item.Title || 'New certification assigned by Admin'),
                    targetEmail: trimmedEmail,
                    type: 'assignment',
                    time: this._formatDisplayDate(item.AssignedDate || item.Created),
                    assignedDate: this._formatDisplayDate(item.AssignedDate || item.Created),
                    status: item.Status || 'Unread',
                    read: (item.Status || '').toLowerCase() !== 'unread',
                    sourceList: 'Notifications'
                }));
            } catch (error) {
                console.error('[Assignments] Failed to load notifications from Notifications. Falling back to legacy sources.', {
                    notificationEndpoint,
                    error
                });
            }
        }

        if (notificationListName === 'LMS_Notifications') {
            const legacyNotificationEndpoint =
                `${siteUrl}/_api/web/lists/getbytitle('LMS_Notifications')/items` +
                `?$filter=TargetEmail eq '${this._escapeODataValue(trimmedEmail)}'` +
                `&$orderby=Id desc`;

            try {
                const data = await this._safeGetJson<any>(
                    legacyNotificationEndpoint,
                    'legacy user notification assignments'
                );
                const notifications = this._toCollection(data)
                    .map((item: any) => ({
                        id: item.Id,
                        title: item.Title || 'Certification Assignment',
                        text: item.NotificationText || 'New certification assigned by Admin',
                        targetEmail: trimmedEmail,
                        type: item.NotificationType || 'assignment',
                        time: this._formatDisplayDate(item.Time || item.Created),
                        assignedDate: this._formatDisplayDate(item.Time || item.Created),
                        status: item.IsRead ? 'Viewed' : 'Unread',
                        read: !!item.IsRead,
                        sourceList: 'LMS_Notifications'
                    }))
                    .filter((item: INotification) => item.type === 'assignment' || item.type === 'priority');

                if (notifications.length > 0) {
                    return notifications;
                }
            } catch (error) {
                console.error('[Assignments] Failed to load notifications from LMS_Notifications. Falling back to enrollment history.', {
                    legacyNotificationEndpoint,
                    error
                });
            }
        }

        const listName = await this._resolveAssignmentNotificationListName();

        if (
            this._enrollmentListCandidates
                .map((candidate) => this._normalizeListTitle(candidate))
                .indexOf(this._normalizeListTitle(listName)) !== -1
        ) {
            try {
                const enrollments = await this.getEnrollments(trimmedEmail);
                return enrollments
                    .filter((item) => (item.userEmail || '').toLowerCase() === trimmedEmail)
                    .map((item) => ({
                        id: item.id || 0,
                        title: item.certName || item.certificateName || 'Certification Assignment',
                        text: 'New certification assigned by Admin',
                        targetEmail: trimmedEmail,
                        type: 'assignment',
                        time: this._formatDisplayDate(item.assignedDate || item.examScheduledDate || item.startDate),
                        assignedDate: this._formatDisplayDate(item.assignedDate || item.examScheduledDate || item.startDate),
                        status: item.listStatus || item.status || 'New',
                        read: ((item.listStatus || item.status || 'New').toString().toLowerCase() !== 'new'),
                        sourceList: listName
                    }));
            } catch (error) {
                console.error('[Assignments] Failed to load notifications from the Enrollment list. Falling back to LMS_Enrollments.', {
                    listName,
                    error
                });
            }
        }

        await this._ensureList('LMS_Enrollments');

        const fallbackEndpoint =
            `${siteUrl}/_api/web/lists/getbytitle('LMS_Enrollments')/items` +
            `?$select=Id,Title,AssignedDate,AssignedToEmail,Status,Created` +
            `&$filter=AssignedToEmail eq '${this._escapeODataValue(trimmedEmail)}'` +
            `&$orderby=Created desc`;

        const fallbackData = await this._safeGetJson<any>(
            fallbackEndpoint,
            'legacy enrollment notifications'
        );
        return this._toCollection(fallbackData).map((item: any) => ({
            id: item.Id,
            title: item.Title || 'Certification Assignment',
            text: 'New certification assigned by Admin',
            targetEmail: trimmedEmail,
            type: 'assignment',
            time: this._formatDisplayDate(item.AssignedDate || item.Created),
            assignedDate: this._formatDisplayDate(item.AssignedDate || item.Created),
            status: item.Status || 'New',
            read: (item.Status || '').toLowerCase() !== 'new',
            sourceList: 'LMS_Enrollments'
        }));
    }

    public static async markAssignmentNotificationAsRead(id: number, sourceList?: string): Promise<void> {
        const listName = sourceList || await this._resolveAssignmentNotificationListName();

        if (listName === 'Notifications') {
            await this._updateListItemStatus(listName, id, 'Viewed');
            return;
        }

        if (listName === 'LMS_Notifications') {
            await this.markNotificationAsRead(id);
            return;
        }

        await this._updateListItemStatus(listName, id, 'Viewed');
    }

    public static async getNotifications(userEmail: string): Promise<INotification[]> {
        const listName = 'LMS_Notifications';
        const siteUrl = this._getSiteUrl();
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$orderby=Id desc`;
        await this._ensureList(listName);
        try {
            const data = await this._safeGetJson<any>(
                endpoint,
                'admin notifications'
            );
            if (!data) {
                return [];
            }

            return this._toCollection(data)
                .map((item: any) => ({
                    id: item.Id,
                    title: item.Title,
                    text: item.NotificationText,
                    targetEmail: item.TargetEmail,
                    type: item.NotificationType,
                    time: item.Time,
                    read: item.IsRead
                }))
                .filter((n: any) =>
                    (n.targetEmail && n.targetEmail.toLowerCase() === userEmail.toLowerCase()) ||
                    n.targetEmail === 'Admin'
                );
        } catch (error) {
            console.error('[Notifications] Failed to fetch SharePoint notifications', {
                endpoint,
                error
            });
            throw error;
        }
    }

    public static async getFilesFromSharedDocuments(): Promise<any[]> {
        const documentLibrary = await this._getDocumentLibrary();
        const serverRelativeUrl = documentLibrary?.RootFolder?.ServerRelativeUrl || '/Shared Documents';
        const endpoint =
            `${this._getSiteUrl()}/_api/web/GetFolderByServerRelativeUrl('${this._escapeODataValue(serverRelativeUrl)}')/Files`;
        const data = await this._safeGetJson<any>(
            endpoint,
            'shared documents files'
        );
        return this._toCollection(data);
    }

    public static async addNotification(notif: INotification): Promise<void> {
        const listName = 'LMS_Notifications';
        const siteUrl = this._getSiteUrl();
        await this._ensureList(listName);
        const digest = await this._getFormDigestValue();
        
        await this._spHttpClient.post(
            `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'X-RequestDigest': digest
                }),
                body: JSON.stringify({
                    Title: notif.title,
                    NotificationText: notif.text,
                    TargetEmail: notif.targetEmail,
                    NotificationType: notif.type,
                    Time: notif.time || new Date().toISOString(),
                    IsRead: false
                })
            }
        );
    }

    public static async markNotificationAsRead(id: number): Promise<void> {
        const listName = 'LMS_Notifications';
        const siteUrl = this._getSiteUrl();
        const digest = await this._getFormDigestValue();
        
        await this._spHttpClient.post(
            `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`,
            SPHttpClient.configurations.v1,
            {
                headers: this._getJsonHeaders({
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'MERGE',
                    'X-RequestDigest': digest
                }),
                body: JSON.stringify({
                    IsRead: true
                })
            }
        );
    }

    /* --- Taxonomy Management --- */
    public static async getTaxonomy(): Promise<any> {
        const siteUrl = this._getSiteUrl();
        const listName = await this._resolveTaxonomyListName();
        await this._ensureList(listName);
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this._escapeODataValue(listName)}')/items`;

        try {
            const response = await this._spHttpClient.get(
                endpoint,
                SPHttpClient.configurations.v1,
                {
                    headers: this._getJsonHeaders()
                }
            );

            if (!response.ok) {
                const errorText = await this._readErrorBody(response);
                const isHtmlError = errorText.trim().toLowerCase().indexOf('<html') !== -1;

                console.error('[Taxonomy] SharePoint taxonomy request failed', {
                    listName,
                    status: response.status,
                    statusText: response.statusText,
                    responsePreview: errorText.substring(0, 500)
                });

                if (isHtmlError) {
                    throw new Error(`SharePoint returned HTML for taxonomy sync (HTTP ${response.status}). Verify the current site authentication and the taxonomy list title.`);
                }

                throw new Error(`Failed to fetch taxonomy items (HTTP ${response.status} ${response.statusText}): ${errorText.substring(0, 300) || 'No error details returned.'}`);
            }

            const data = await this._readJson<any>(response);
            const taxonomy: any = {
                departments: [],
                businessUnits: [],
                roles: [],
                locations: [],
                groups: [],
                categories: []
            };

            (data?.value || []).forEach((item: any) => {
                const category = (item.Category || '').toLowerCase();
                const items = JSON.parse(item.SchemaData || '[]');

                if (category === 'departments') taxonomy.departments = items;
                if (category === 'businessunits') taxonomy.businessUnits = items;
                if (category === 'roles') taxonomy.roles = items;
                if (category === 'locations') taxonomy.locations = items;
                if (category === 'groups') taxonomy.groups = items;
                if (category === 'categories') taxonomy.categories = items;
            });

            return taxonomy;
        } catch (error) {
            console.error('[Taxonomy] SharePoint taxonomy sync failed:', error);
            return null;
        }
    }

    public static async updateTaxonomy(category: string, items: string[]): Promise<void> {
        const siteUrl = this._getSiteUrl();
        const listName = await this._resolveTaxonomyListName();
        await this._ensureList(listName);
        const digest = await this._getFormDigestValue();
        
        // Find existing to update or create
        const checkResp = await this._spHttpClient.get(
            `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=Category eq '${this._escapeODataValue(category)}'`,
            SPHttpClient.configurations.v1,
            { headers: this._getJsonHeaders() }
        );
        
        const checkData = await this._readJson<any>(checkResp);
        const existingItems = this._toCollection(checkData);
        
        const body = JSON.stringify({
            Title: category,
            Category: category,
            SchemaData: JSON.stringify(items)
        });

        if (existingItems && existingItems.length > 0) {
            const itemId = existingItems[0].Id;
            await this._spHttpClient.post(
                `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`,
                SPHttpClient.configurations.v1,
                {
                    headers: this._getJsonHeaders({
                        'IF-MATCH': '*',
                        'X-HTTP-Method': 'MERGE',
                        'X-RequestDigest': digest
                    }),
                    body: body
                }
            );
        } else {
            await this._spHttpClient.post(
                `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`,
                SPHttpClient.configurations.v1,
                {
                    headers: this._getJsonHeaders({
                        'X-RequestDigest': digest
                    }),
                    body: body
                }
            );
        }
    }

}

