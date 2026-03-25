# SPFx Learning Center - Build & Deployment Guide

## Quick Start: Build and Deploy

### Prerequisites

- **Node.js** 18.17.1 - 18.x (not 19.x or 20.x)
- **npm** 8.x or higher
- **Global tools**:
  ```bash
  npm install -g @microsoft/sharepoint-cli yo @microsoft/generator-sharepoint gulp
  ```

### Step 1: Build the Solution

From `spfx-learning-center/` directory:

```bash
# Install dependencies
npm install

# Build bundle (generates .js and .d.ts in lib/)
npm run build

# Or with production optimizations
gulp bundle --ship
```

### Step 2: Create Package (.sppkg)

```bash
# Package for App Catalog
gulp package-solution --ship
```

**Output**: `sharepoint/solution/spfx-learning-center.sppkg`

### Step 3: Upload to SharePoint App Catalog

1. Navigate to: `https://[tenant]-admin.sharepoint.com/sites/appcatalog`

2. Go to **Apps for SharePoint** library

3. Upload the `.sppkg` file:
   - Click **Upload** → Select from `sharepoint/solution/` folder
   - Choose **"spfx-learning-center.sppkg"**

4. **Important**: Check **"Make this a tenant-wide app deployment"** for organization-wide availability

5. Click **Deploy**

6. **Approve** API permission requests if prompted (usually automatic for SharePoint API)

### Step 4: Deploy to Your Site

**Option A: Automatic (Tenant-Wide)**

If you checked "Tenant-wide deployment", the app is immediately available in all site App Catalogs.

**Option B: Manual Site Deployment**

1. Go to your **SharePoint site**

2. Go to **Site Settings** → **Add an App**

3. Click **From Your Organization**

4. Find and install:
   - **Learning Center** web part (learner view)
   - **Admin Access** web part (admin portal)

5. Go to a site page and click **+ Add a web part** to insert them

### Step 5: Verify Installation

1. Create/edit a **Modern Page** in the site

2. Click **+ Add a web part**

3. Search for:
   - "Learning Center" ✅
   - "Admin Access" ✅

4. Add one to test the installation

## Build Troubleshooting

### Issue: "Node version not supported"

```
Error: The target of the property "node" in the package.json file is ~18.17.1 which includes version 19.x.x or 20.x.x
```

**Fix**: Downgrade Node.js
```bash
# Using nvm (Node Version Manager)
nvm install 18.17.1
nvm use 18.17.1
node --version  # Should show v18.17.1

# Or download from https://nodejs.org/ (LTS 18.x)
```

### Issue: "Cannot find tsconfig.json"

```
Error: ENOENT: no such file or directory, open '.../tsconfig.json'
```

**Fix**: Ensure you're in the correct directory
```bash
cd spfx-learning-center/
ls tsconfig.json  # Should exist
npm run build
```

### Issue: "Port 4321 already in use"

```
Error: EADDRINUSE: address already in use :::4321
```

**Fix**: Kill existing process or use different port
```bash
# Kill on port 4321
npx kill-port 4321

# Or use custom port
gulp serve --port 5500
```

## Local Development

### Running Dev Server

```bash
cd spfx-learning-center/
gulp serve --nobrowser
```

**Workbench URL**: `http://localhost:4321/temp/workbench.html`

Open in browser to load web parts in isolation for testing.

### Making Changes

1. Edit TypeScript/React files in `src/` directory
2. Gulp automatically watches and recompiles
3. Refresh browser to see changes (usually Hot Module Reload works)

### Common Dev Changes

**Change web part name**: `src/webparts/[name]/[Name]WebPart.manifest.json`

**Change web part title**: Look for `Title` and `Description` fields in manifest

**Add new component**: Create in `src/webparts/[name]/components/` and import

## Production Build Checklist

Before deploying to production:

- [ ] Run `npm install` to get latest dependencies
- [ ] Run `npm run build` successfully with no errors
- [ ] Run `gulp clean && gulp bundle --ship` for optimized build
- [ ] Check `lib/` folder was created
- [ ] Run `gulp package-solution --ship` to generate `.sppkg`
- [ ] Test in dev/staging SharePoint site first
- [ ] Document any custom changes made
- [ ] Have backup of current solution if replacing existing

## Post-Deployment Verification

After deploying to SharePoint:

### 1. Check App Installation

Go to site **Site Settings** → **Manage apps**
- Should see "spfx-learning-center" listed
- Status should be "OK"

### 2. Test Web Parts

**Learning Center Web Part**:
- Page loads without errors
- Current user sees their enrollments (if any)
- Can navigate certification paths

**Admin Portal (Admin Access Web Part)**:
- Admin can access the portal
- Can see dashboard
- Can create test certification paths
- Can sync SharePoint users

### 3. Verify Lists Creation

First admin use should auto-create SharePoint lists:
- Navigate to Site Contents
- Should see auto-created lists:
  - `LMS_Enrollments`
  - `LMS_Notifications`
  - `LMS_AdminCerts`
  - `LMS_Taxonomy`
  - `LMS_ContentLibrary` (optional, may sync directly to Documents)

### 4. Test Data Flow

**Admin**: Create a test certification
- Goes to Admin Access → Certification Management → Create New Path
- Fill details and save
- Verify appears in `LMS_AdminCerts` list

**Learner**: Assign certification
- Admin assigns to test learner in Users Management
- Verify appears in `LMS_Enrollments` list
- Learner should see it in Learning Center

## Updating/Upgrading

### To Update Existing Installation

```bash
# 1. Make your code changes in src/
# 2. Rebuild
npm run build

# 3. Re-package
gulp package-solution --ship

# 4. Upload newer .sppkg to App Catalog
# - SharePoint will auto-upgrade on all sites

# 5. Browser cache clearing may be needed
```

### Version Management

Track versions in `package.json`:
```json
{
  "version": "2.0.0",
  "description": "SPFx Learning Center - Standalone (No Backend)"
}
```

## Troubleshooting Deployment Issues

### Issue: "App installation fails with blank error"

**Solution**:
1. Check browser console for JavaScript errors
2. Test in different browser
3. Try in different SharePoint site
4. Check **App Catalog** → verify package is there

### Issue: "Lists don't auto-create on first use"

**Solution**:
1. Manual create lists in Site Settings
2. Or run admin portal once with sufficient permissions
3. Check browser console for errors during load

### Issue: "Web part loads but shows empty"

**Solution**:
1. Ensure logged in with appropriate permissions
2. Clear browser cache: `Ctrl+Shift+Delete`
3. Hard refresh page: `Ctrl+Shift+R`
4. Check browser console for errors
5. Verify SharePoint lists exist and have data

### Issue: "Cannot read properties of undefined (reading 'pageContext')"

**Solution**: Web part requires proper SharePoint context
1. Ensure it's added to SharePoint page (not dev workbench)
2. Check that `context` prop is being passed in web part manifest
3. Verify `ILearningCenterProps` or `IAdminDashboardProps` interface in web part code

## CI/CD Integration (Optional)

### GitHub Actions Example

Create `.github/workflows/build-and-deploy.yml`:

```yaml
name: Build SPFx

on:
  push:
    branches: [ main ]
    paths:
      - 'spfx-learning-center/**'

jobs:
  build:
    runs-on: windows-latest
    
    strategy:
      matrix:
        node-version: [18.17.1]
    
    steps:
    - uses: actions/checkout@v2
    
    - name: Use Node.js ${{ matrix.node-version }}
      uses: actions/setup-node@v2
      with:
        node-version: ${{ matrix.node-version }}
    
    - name: Install dependencies
      run: |
        cd spfx-learning-center
        npm install
    
    - name: Build solution
      run: |
        cd spfx-learning-center
        npm run build
    
    - name: Package solution
      run: |
        cd spfx-learning-center
        gulp package-solution --ship
    
    - name: Upload artifact
      uses: actions/upload-artifact@v2
      with:
        name: sharepoint-package
        path: spfx-learning-center/sharepoint/solution/*.sppkg
```

Then manually upload `.sppkg` from artifacts to SharePoint App Catalog.

## File Structure

After running `npm run build`:

```
spfx-learning-center/
├── lib/                          # 📁 Compiled output (created by build)
│   ├── index.js                 # Bundle
│   ├── index.d.ts               # Types
│   └── webparts/
│       ├── adminAccess/         # Admin portal compiled
│       └── learningCenter/       # Learner portal compiled
├── src/                          # 📁 TypeScript source
│   ├── webparts/                # Web part components
│   └── declarations.d.ts        # Global type defs
├── sharepoint/                   # 📁 Deployment files
│   └── solution/
│       └── spfx-learning-center.sppkg  # ⚙️ Upload this!
├── config/
│   ├── config.json              # SPFx configuration
│   ├── package-solution.json    # Package metadata
│   └── serve.json               # Dev server config
├── package.json                  # npm scripts and dependencies
├── tsconfig.json                # TypeScript config
├── gulpfile.js                  # Build tasks
└── eslintrc.js                  # Linting rules
```

## Support

For build issues:
- Check [Microsoft SPFx Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)
- See error logs: `gulp build --verbose`
- SharePoint modern web parts require modern browsers (Edge, Chrome, Safari)

---

**Last Updated**: March 2026
**SPFx Version**: 1.20.0+
**Node Requirement**: v18.17.1 - v18.x.x
