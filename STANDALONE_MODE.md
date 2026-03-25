# SPFx Learning Center - Standalone Mode (No Backend Server Required)

## Overview

The SPFx Learning Center web parts are now fully operational **without any backend server dependency**. All data is stored and managed directly through **SharePoint REST APIs**.

### Key Architecture Changes

✅ **Removed Backend Dependency**
- No longer requires Node.js Express backend server (port 5000)
- No database connection needed
- No JWT authentication layer required for the SPFx components

✅ **Direct SharePoint Integration**
- All enrollment data stored in `LMS_Enrollments` SharePoint list
- Notifications stored in `LMS_Notifications` list
- Taxonomy data stored in `LMS_Taxonomy` list
- Admin certifications stored in `LMS_AdminCerts` list
- Content assets stored in SharePoint document library

✅ **Offline-First Architecture**
- Local browser storage (localStorage) acts as a sync cache
- Changes sync immediately to SharePoint
- Automatic retry when connectivity returns
- Works even if SharePoint is temporarily unavailable

## Prerequisites

### SharePoint Site Requirements

1. **SharePoint Online** site collection (Office 365)
2. **App/Admin Catalog** configured in the tenant
3. **Standard Lists** to be auto-created by SharePointService:
   - `LMS_Enrollments`
   - `LMS_Notifications`
   - `LMS_AdminCerts`
   - `LMS_Taxonomy`
   - `LMS_ContentLibrary`

4. **Document Library**: Standard "Documents" or "Shared Documents" library

## Deployment Instructions

### 1. Build the SPFx Package

From the `spfx-learning-center` directory:

```bash
# Install dependencies (if not already done)
npm install

# Build the solution (creates .sppkg file)
npm run build

# Or build with production optimizations
gulp bundle --ship
```

### 2. Upload to App Catalog

1. Navigate to **SharePoint App Catalog** (`https://[tenant]-admin.sharepoint.com/sites/appcatalog`)
2. Go to **Apps for SharePoint**
3. Upload the generated `.sppkg` file from:
   - `spfx-learning-center/sharepoint/solution/spfx-learning-center.sppkg`
4. **Check "Make this a tenant-wide deployed app" option** if deploying site-wide
5. Click **Deploy**
6. **Approve** the permissionrequests

### 3. Add Web Parts to SharePoint Site

1. Go to your **Learning Center SharePoint site**
2. Create or edit a page
3. Click **+ Add a new web part**
4. Search for:
   - **"Learning Center"** (learner view)
   - **"Admin Access"** (admin portal)
5. Add and configure as needed

### 4. Initialize Data (First Run)

On first load, the web parts will:

1. Auto-create required SharePoint lists with fields
2. Initialize taxonomies with defaults (Departments, Roles, etc.)
3. Sync any browser local storage cache to SharePoint
4. Display wizard if SharePoint initialization needs manual approval

## Features by Web Part

### Admin Access Web Part

**Portal URL**: `#/admin` (when in admin mode)

**Features**:
- Dashboard with real-time enrollment stats (from SharePoint data)
- Certification Path Management
- Learner Management & User Syncing
- Content Library (upload to SharePoint Documents)
- Assessment Builder
- Taxonomy Management
- Enrollment Tracking
- Audit Logs
- System Configuration

**Data Storage**: All via SharePointService

### Learning Center Web Part

**Portal URL**: Root view for learners

**Features**:
- View assigned certifications from `LMS_Enrollments`
- Complete self-paced learning paths
- Take assessments (stored in localStorage → syncs to SharePoint)
- Track progress
- Receive notifications from `LMS_Notifications`

## Development

### Running Locally

```bash
# Terminal 1: Start Gulp serve
gulp serve --nobrowser

# Terminal 2: Open browser to local dev URL
# SharePoint dev workbench shows web parts in isolation mode
```

### Disabling Backend (Confirming Standalone Mode)

**The following services are no longer active:**
- ❌ No Node.js Express backend  (`http://localhost:5000`)
- ❌ No PostgreSQL database connection
- ❌ No GraphAPI (Microsoft Graph) calls from SPFx layer 

**All data flows through SharePoint REST APIs only.**

## Configuration

### Environment Variables (If Needed)

The SPFx components read configuration from:

1. **localStorage** keys:
   - `lmsAllUsers` - cached user directory
   - `scheduledCerts` - cached enrollments
   - `lmsAdminAssessments` - admin-created assessments
   - `lmsTaxonomyData` - taxonomy entries
   - `lmsContentLibrary` - content assets

2. **SharePoint Lists** (primary source of truth):
   - `LMS_Enrollments`
   - `LMS_Notifications`
   - `LMS_AdminCerts`
   - `LMS_Taxonomy`
   - `LMS_ContentLibrary`

### Customization

To add a new data type:

1. Define it in `SharePointService.ts`:
   ```typescript
   private static _getListFieldDefinitions(listName: string): Array<{ name: string; type: number }> {
       if (listName === 'YOUR_LIST') {
           return [
               { name: 'Field1', type: 2 }, // Text
               { name: 'Field2', type: 3 }, // Note (multiline text)
           ];
       }
   }
   ```

2. Add CRUD methods:
   ```typescript
   public static async addYourItem(data: any): Promise<number> {
       return this._ensureAndPostToList('YOUR_LIST', data);
   }
   ```

3. Use in components via `SharePointService.addYourItem()`

## Troubleshooting

### Lists Not Creating

**Issue**: "STORAGE NOT FOUND" error message

**Solution**:
1. Ensure site admin role on the SharePoint site
2. Check that "Documents" library exists
3. Manually create lists in SharePoint:
   - Go to **Site Settings** → **List and Libraries**
   - Create custom lists matching field definitions in SharePointService
   - Ensure user has Contribute+ permissions

### Permissions Errors

**Issue**: "Access Denied" or "Unsupported operation"

**Solution**:
1. Grant **Site Owner** or **Site Admin** role
2. Approve any pending **API permission** requests
3. Clear browser cache and re-authenticate:
   ```javascript
   // In browser console:
   localStorage.clear();
   sessionStorage.clear();
   window.location.reload();
   ```

### Data Not Syncing

**Issue**: Changes made in Admin Portal don't appear for learners

**Solution**:
1. Check SharePoint list directly:
   - Go to **LMS_Enrollments**  → Verify records exist
2. Force sync:
   ```javascript
   // In browser console:
   window.dispatchEvent(new StorageEvent('storage', { key: 'scheduledCerts' }));
   ```
3. Hard refresh both web parts (Ctrl+Shift+R)

## Monitoring & Auditing

### Viewing Audit Logs

Admin Portal → **Audit & Assignment Control** → **System Logs**

Logs are stored in localStorage but synced to SharePoint. All admin actions are logged:
- User assignments
- Enrollment revocations
- Taxonomy changes
- Content uploads

### Checking SharePoint Lists

```
Site URL: https://[tenant].sharepoint.com/sites/YourSite
Go to: Site Contents → [List Name]
Example: https://[tenant].sharepoint.com/sites/LearningCenter/lists/LMS_Enrollments
```

## Performance Optimization

### Caching Strategy

The system uses **3-level caching**:

1. **In-Memory** (React state) - Fastest
2. **localStorage** - Offline capable
3. **SharePoint** - Persistent, auditable

### Polling Intervals

- Admin Dashboard: 5 seconds
- Enrollment data: 30 seconds
- Learner notifications: 10 seconds

Adjust in component `useEffect` hooks if needed.

## Security Considerations

✅ Uses **SharePoint's built-in security** model
✅ No credentials stored in browser (uses session auth)
✅ All data at rest encrypted by SharePoint Online
✅ Audit trails maintained for compliance

### Best Practices

1. **Restrict Admin Web Part** to site owners/admins only
2. **Use Microsoft 365 Groups** for enrollment batches
3. **Enable audit logging** in tenant admin center
4. **Review audit logs** regularly in the portal

## Migration from Backend Mode

If you had data in the old backend:

1. Export from PostgreSQL database
2. Import into SharePoint lists via:
   - Admin Portal → **Bulk Import**
   - Or Power Automate connector
   - Or manual CSV upload

## Support & Troubleshooting

### Diagnostic Commands (Browser Console)

```javascript
// Check SharePoint service status
console.log(SharePointService._siteUrl);

// Examine cached enrollment data
console.log(JSON.parse(localStorage.getItem('scheduledCerts')));

// Clear all cached data (WARNING: Deletes offline cache)
localStorage.clear();

// Force refresh from SharePoint
location.reload(true);
```

### Common Issues Dashboard

| Issue | Symptom | Fix |
|-------|---------|-----|
| No enrollments showing | Blank list | Check `scheduledCerts` in localStorage & `LMS_Enrollments` list |
| Notifications not appearing | No bell alerts | Verify `LMS_Notifications` list has records |
| Can't upload content | Upload fails | Ensure "Documents" library exists and has contribute rights |
| Admin portal won't load | Blank screen | Check browser console for errors; run site sync |

---

**Version**: 2.0 (Standalone, No Backend)
**Last Updated**: March 2026
**Requirement**: SharePoint Online (SPO) - Not compatible with on-premises SharePoint 2019
