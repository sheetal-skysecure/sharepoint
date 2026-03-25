# SPFx Learning Center - Testing & Verification Guide

## Pre-Launch Testing Checklist

Before deploying to production, verify all components work correctly in standalone mode.

## Test 1: Build Verification

### Objective
Ensure the SPFx package builds without errors and creates the `.sppkg` file.

### Steps

1. **Navigate to SPFx directory**
   ```bash
   cd spfx-learning-center/
   ```

2. **Clean previous builds**
   ```bash
   npm run clean
   # or
   gulp clean
   ```

3. **Install dependencies**
   ```bash
   npm install
   ```

4. **Build the solution**
   ```bash
   npm run build
   # Expected output: lib/ folder created with compiled files
   ```

5. **Verify no errors**
   - Look for "BUILD SUCCESSFUL" or similar message
   - Check `lib/` folder exists with files

6. **Package for deployment**
   ```bash
   gulp package-solution --ship
   ```

7. **Verify .sppkg file**
   ```bash
   ls sharepoint/solution/*.sppkg
   # Should list: spfx-learning-center.sppkg
   ```

**Expected Result**: ✅ Build completes, `.sppkg` file created

**If Failed**: 
- Check Node version: `node --version` (should be 18.17.1 - 18.x)
- See [BUILD_AND_DEPLOY.md](./BUILD_AND_DEPLOY.md) troubleshooting section

---

## Test 2: Admin Portal Dashboard Load

### Objective
Verify Admin web part loads without backend server and displays calculated stats.

### Setup
1. Upload `.sppkg` to SharePoint App Catalog
2. Deploy to test site
3. Create test page and add **Admin Access** web part

### Steps

1. **Open Admin Portal**
   - Navigate to page with Admin Access web part
   - Click web part to go to admin dashboard

2. **Verify Dashboard Stats**
   - Should show real-time stats calculated from SharePoint data:
     - ✅ "0 Certifications" initially (increase as you add)
     - ✅ "0 Enrolled" initially
     - ✅ "0 In Progress" initially
     - ✅ "0 Blocked" (at capacity)

3. **Check Report View**
   - Click **Detailed Reports** (Reports view)
   - Should NOT show "Backend request failed" error
   - Should display charts with local data
   - Top paths should be empty or show actual enrollments if data exists

4. **Verify no console errors**
   - Open browser Developer Tools (F12)
   - Go to **Console** tab
   - Should NOT see errors about:
     - "BackendService is not defined"
     - "localhost:5000"
     - "Cannot fetch from http://localhost:5000"

**Expected Result**: ✅ Dashboard loads, no backend errors, shows local stats

**If Failed**:
- Check if `SharePointService` initialization ran
- Verify SharePoint lists exist
- Look for permission errors in console

---

## Test 3: SharePoint Lists Auto-Creation

### Objective
Verify required SharePoint lists are created automatically on first use.

### Steps

1. **First Admin Access**
   - If this is first load, admin portal should trigger list creation
   - Watch for toast notification: "Initializing SharePoint storage..."

2. **Verify Lists Created**
   - Go to SharePoint site
   - Click **Site Contents**
   - Should see new lists:
     - ✅ `LMS_Enrollments`
     - ✅ `LMS_Notifications`
     - ✅ `LMS_AdminCerts`
     - ✅ `LMS_Taxonomy`
     - ✅ `LMS_ContentLibrary` (may be optional)

3. **Manual Verification (if auto-create fails)**
   ```
   Site Settings → Lists and Libraries → Create custom list
   Name: LMS_Enrollments
   Columns: UserEmail, UserName, CertCode, CertName, StartDate, EndDate, Status, Progress, CertificateName (all Text)
   ```

**Expected Result**: ✅ All required SharePoint lists exist

**If Failed**:
- May need Site Owner/Admin permissions
- Check browser console for API errors
- Manually create lists (see [STANDALONE_MODE.md](./STANDALONE_MODE.md) Troubleshooting)

---

## Test 4: Data Creation & Storage

### Objective
Verify data created admin portal correctly stores in SharePoint.

### Steps

1. **Create Test Certification Path**
   - Admin Portal → Certification Management → Create New Path
   - Fill in:
     - Name: "Test Azure Fundamentals"
     - Code: "AZ-TEST"
     - Description: "Test certification"
     - Provider: "Microsoft"
   - Click **Deploy to Portal**

2. **Verify SharePoint Storage**
   - Go to Site Contents → `LMS_AdminCerts` list
   - Should see new entry:
     - Title: "Test Azure Fundamentals"
     - CertCode: "AZ-TEST"

3. **Create Test User**
   - Admin Portal → Learner Management → Create User
   - Fill details and save
   - Verify in `LMS_AllUsers` localStorage or direct list

4. **Assign Certification**
   - Admin Portal → Learner Management
   - Select test user → Assign Certification
   - Choose test path created above
   - Click **Confirm Enrollment**

5. **Verify Enrollment Created**
   - Go to Site Contents → `LMS_Enrollments` list
   - Should see:
     - UserEmail: [test user email]
     - CertCode: "AZ-TEST"
     - Status: "scheduled"

**Expected Result**: ✅ Data persists in SharePoint without backend server

**If Failed**:
- Check SharePoint list permissions (need Contribute)
- Verify SPHttpClient has context from web part
- Check browser console for CORS or permission errors

---

## Test 5: Learner Portal Functionality

### Objective
Verify learners see assigned content and can progress.

### Steps

1. **Open Learning Center Web Part**
   - Navigate to page with Learning Center web part
   - Should load without backend errors

2. **View Assigned Certifications**
   - Should display certifications assigned via admin portal
   - Each shows: Name, Code, Progress bar

3. **Navigate Certification Content**
   - Click on assigned certification
   - Should show course modules/lessons
   - No backend API calls needed

4. **Progress Tracking**
   - View learner's progress bar
   - Data should pull from `LMS_Enrollments` SharePoint list

5. **View Notifications**
   - Learner should see notifications
   - Assignments and important updates should appear
   - Check notifications from `LMS_Notifications` list

**Expected Result**: ✅ Learner sees data, can navigate, no backend required

**If Failed**:
- Verify assignments were created in admin portal
- Check `LMS_Enrollments` has data
- Check learner has view permissions on lists

---

## Test 6: Offline Functionality

### Objective
Verify system works when offline and syncs when back online.

### Steps

1. **Create State While Online**
   - Admin Portal → Create certification
   - Create user
   - Assign certification
   - Verify all appears in SharePoint lists

2. **Go Offline**
   - Browser DevTools → Network tab → Set to "Offline"

3. **Make Changes**
   - Try to create another certification (might cache locally)
   - Try to view enrollments
   - Should fall back to localStorage cache

4. **Go Back Online**
   - Set network back to normal
   - Refresh page
   - Should sync cached data back to SharePoint
   - Verify all changes persisted

**Expected Result**: ✅ Graceful offline mode, automatic sync on reconnect

**If Failed**:
- May indicate issue with localStorage cache mechanism
- Check browser console for sync errors

---

## Test 7: Content Upload & Storage

### Objective
Verify file uploads go directly to SharePoint Document Library.

### Steps

1. **Upload Content**
   - Admin Portal → Content Library → Click "Choose File"
   - Select test document (PDF, Word, etc.)
   - Upload

2. **Verify Upload Destination**
   - Go to SharePoint Site → Documents library
   - File should appear there (no backend server copy needed)

3. **Storage Verification**
   - Check `LMS_ContentLibrary` list has entry (if using that approach)
   - Or verify file URL in Documents library

4. **Access from Admin Portal**
   - Admin Portal → Content Library
   - Should list uploaded file
   - URL should point to SharePoint

**Expected Result**: ✅ Files stored in SharePoint, accessible via portal

**If Failed**:
- Verify "Documents" library exists on site
- Check permissions (need Contribute to Documents library)
- If custom library used, ensure it's properly configured

---

## Test 8: No Backend Server Dependency

### Objective
Confirm system works WITHOUT Node.js backend running.

### Steps

1. **Verify Backend NOT Running**
   - Open browser console (F12)
   - Type: `fetch('http://localhost:5000/').catch(() => console.log('Backend NOT running'))`
   - Should see "Backend NOT running" message
   - Should NOT see successful response

2. **Test Portal Functions**
   - Refresh Admin Portal page
   - Should still work normally
   - Try: Create cert, Assign user, View enrollments
   - All should work despite backend being unavailable

3. **Check Error Logs**
   - Browser console should NOT show:
     - "Failed to connect to backend"
     - "BackendService is not available"
     - "http://localhost:5000" connection errors
   - Only legitimate SharePoint API errors (if any)

4. **Monitor Network Tab**
   - Open DevTools → Network tab
   - Perform actions (create, assign, etc.)
   - Should see SharePoint API calls:
   - Should NOT see any calls to `localhost:5000`
   - Should NOT see any calls to backend domain

**Expected Result**: ✅ All features work independently of backend server

**If Failed**:
- Indicates code still depends on backend
- Check for remaining `BackendService` calls
- Review code changes in AdminPortal.tsx and CertificationsList.tsx

---

## Test 9: Administrator Features

### Objective
Verify all admin-exclusive features work without backend.

### Steps

**Certification Management**:
- ✅ Create new certification path
- ✅ Edit existing certification
- ✅ Delete certification
- ✅ View enrollment capacity

**User Management**:
- ✅ Sync SharePoint users
- ✅ Add manual learner entry
- ✅ Edit learner profile
- ✅ Assign certification to learner
- ✅ Revoke certification

**Content Management**:
- ✅ Upload content file
- ✅ View content library
- ✅ Delete content

**Taxonomy Management**:
- ✅ Add department
- ✅ Add role
- ✅ Add location
- ✅ Add business unit
- ✅ Delete taxonomy item

**Assessment Management**:
- ✅ Create assessment
- ✅ Auto-generate questions (AI)
- ✅ Bulk upload questions (CSV)
- ✅ Publish assessment
- ✅ Push assessment to learner
- ✅ View learner results

**Audit & Compliance**:
- ✅ View audit logs
- ✅ View active assignments
- ✅ Revoke single assignment
- ✅ Bulk revoke assignments

**Expected Result**: ✅ All admin features function without backend

**If Failed**:
- Check individual feature for backend calls
- Review browser console for API errors

---

## Test 10: Performance & Scalability

### Objective
Verify system performs well with realistic data volumes.

### Steps

1. **Bulk Create Data**
   - Create 50+ certifications
   - Create 100+ learner users
   - Assign 500+ enrollments
   - Monitor performance

2. **Dashboard Load Time**
   - Measure dashboard load time (should be <3 seconds)
   - Stats should calculate within <1 second
   - No timeout errors

3. **List Operations**
   - List all enrollments: <2 second load
   - Filter/search: <1 second
   - Sort: <1 second

4. **Memory Usage**
   - Check browser memory in DevTools
   - Should not continuously grow (leak detection)
   - Clear cache between operations

**Expected Result**: ✅ Fast load times, efficient use of browser memory

**If Failed**:
- May need to optimize list queries
- Implement pagination for large lists
- Consider reducing polling frequency

---

## Test 11: SharePoint Site Template Compatibility

### Objective
Verify works across different SharePoint site templates.

### Test on:
- ✅ Communication Site
- ✅ Team Site (public)
- ✅ Team-connected Site
- ✅ Hub Site

**Expected Result**: ✅ Web parts function on all site types

**If Failed**:
- Some templates may have permission restrictions
- Verify app catalog deployment permissions

---

## Post-Testing Checklist

- [ ] Build completes successfully
- [ ] Admin portal loads without backend errors
- [ ] SharePoint lists create automatically
- [ ] Data persists in SharePoint lists
- [ ] Learner portal displays assigned content
- [ ] Offline mode works with sync
- [ ] File uploads go to SharePoint
- [ ] Backend server NOT running/required
- [ ] All admin features work
- [ ] Performance acceptable
- [ ] Works on multiple site types
- [ ] No console errors related to backend
- [ ] Production `.sppkg` ready for deployment

## Reporting Issues

If any test fails:

1. **Collect Information**
   ```javascript
   // In browser console:
   console.log(SharePointService._siteUrl);
   console.log('Local enrollments:', localStorage.getItem('scheduledCerts'));
   ```

2. **Copy browser DevTools output**
   - Console errors
   - Network tab requests
   - Application → LocalStorage contents

3. **Document steps to reproduce**

4. **Share findings**

---

**Testing Status**: Ready for deployment after passing all tests
**Version**: 2.0 Standalone
**Last Updated**: March 2026
