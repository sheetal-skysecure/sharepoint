# spfx-learning-center

## Summary

**SharePoint Framework Learning Center Web Parts** - A complete Learning Management System (LMS) built on Microsoft 365 that requires NO external backend server. All data is stored directly in SharePoint Online.

**Key Features:**
- ✅ Standalone SharePoint-based architecture (no backend required)
- ✅ Admin portal for certification management
- ✅ Learner portal for course tracking and assessments
- ✅ Direct content upload to SharePoint
- ✅ Offline capability with automatic sync
- ✅ Real-time dashboard with engagement metrics

## ⭐ **NOW BACKEND-FREE!**

We've removed all dependencies on the Node.js Express backend server. The system now:
- ✅ Stores all data in SharePoint lists
- ✅ Works without any running backend server
- ✅ Uploads content directly to SharePoint
- ✅ Syncs data via localStorage cache
- ✅ Simplified deployment and maintenance

**→ [START HERE: QUICK_START.md](QUICK_START.md)** for deployment instructions

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.20.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

- **Node.js**: 18.17.1 (18.x only - NOT 19.0+ or 20.0+)
- **npm**: 9.x or higher
- **SharePoint Online**: Modern experiences enabled
- **Modern browser**: Edge, Chrome, or Firefox
- **Permissions**: Site owner or admin to create lists and deploy web part

## Architecture

### Data Storage (No Backend!)
```
SharePoint Lists (Primary Storage)
├── LMS_Enrollments     → Track learner progress
├── LMS_Notifications   → User notifications
├── LMS_AdminCerts      → Certification definitions
├── LMS_Taxonomy        → Departments & roles
└── LMS_ContentLibrary  → Content references

LocalStorage (Cache/Sync Layer)
└── Auto-syncs with SharePoint every 30 seconds

SharePoint Documents (File Storage)
└── Direct upload location for course content
```

## Documentation

| Document | Purpose |
|----------|---------|
| [QUICK_START.md](QUICK_START.md) | ⭐ Start here - 5-minute deployment guide |
| [BUILD_AND_DEPLOY.md](BUILD_AND_DEPLOY.md) | Detailed build & deployment instructions |
| [STANDALONE_MODE.md](STANDALONE_MODE.md) | Complete technical architecture & troubleshooting |
| [TEST_AND_VERIFY.md](TEST_AND_VERIFY.md) | Pre-production testing checklist |
| [BACKEND_REMOVAL_SUMMARY.md](BACKEND_REMOVAL_SUMMARY.md) | What changed & why |

## Solution

| Component | Details |
| --------- | ------- |
| **Web Parts** | Admin Access Portal + Learning Center Portal |
| **Storage** | SharePoint Online Lists + Documents |
| **Authentication** | Azure AD (via SharePoint context) |
| **Framework** | React + TypeScript |
| **Build Tool** | Gulp + Webpack |
| **Version** | 2.0 (Backend Removed) |

## Version history

| Version | Date | Comments |
|---------|------|----------|
| 2.0 | March 2024 | Removed Node.js backend dependency - now fully standalone |
| 1.5 | Feb 2024 | Optimized SharePoint integration |
| 1.0 | Jan 2024 | Initial release with backend server |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Getting Started - Minimal Path to Awesome

### Option 1: Quick Deployment (Recommended)
```bash
cd spfx-learning-center
npm install
npm run build
gulp package-solution --ship
# Upload sharepoint/solution/spfx-learning-center.sppkg to App Catalog
```
→ See [QUICK_START.md](QUICK_START.md) for detailed steps

### Option 2: Local Development
```bash
cd spfx-learning-center
npm install
gulp serve
# Open https://[tenant].sharepoint.com/sites/[site] and add web part
```

### First Run (Automatic Setup)
1. Add Admin Access web part to a page
2. Admin portal opens
3. SharePoint lists auto-create automatically
4. No backend server needed ✅

---

## Key Features

### Admin Portal
- ✅ Dashboard with real-time stats
- ✅ Certification path management
- ✅ Learner synchronization from SharePoint
- ✅ Content upload directly to SharePoint
- ✅ Assessment creation & deployment
- ✅ Taxonomy management
- ✅ Enrollment tracking
- ✅ Audit logging

### Learner Portal
- ✅ View assigned certifications
- ✅ Track progress in real-time
- ✅ Complete self-paced content
- ✅ Take assessments
- ✅ Receive notifications
- ✅ Update user profile

### System Features
- ✅ Works offline with localStorage cache
- ✅ Auto-syncs when online
- ✅ No backend server required
- ✅ SharePoint-native compliance
- ✅ Azure AD authentication
- ✅ Responsive design

---

## Troubleshooting

**Build fails with Node version error?**
- Ensure you're on Node 18.17.1: `node --version`
- If wrong: Install nvm-windows and switch to 18.x

**Lists not creating in SharePoint?**
- See [STANDALONE_MODE.md](STANDALONE_MODE.md) > Troubleshooting

**Data not syncing?**
- Check browser console for errors
- Verify SharePoint site permissions
- See [TEST_AND_VERIFY.md](TEST_AND_VERIFY.md)

**Need more help?**
- See [BACKEND_REMOVAL_SUMMARY.md](BACKEND_REMOVAL_SUMMARY.md) for what changed
- See [BUILD_AND_DEPLOY.md](BUILD_AND_DEPLOY.md) for build issues

---

## Support & Questions

📖 **Complete Documentation**: See files listed above

🚀 **Ready to deploy?** → Start with [QUICK_START.md](QUICK_START.md)

✅ **Status**: Production Ready - No backend server needed!

> Include any additional steps as needed.

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
