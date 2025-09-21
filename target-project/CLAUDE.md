# Staff Rota Manager - Claude Development Notes

## 🔒 **CRITICAL MENU PROTECTION RULES**

### ⚠️ **NEVER MODIFY THESE - THEY MAKE THE MENU WORK**

The "Staff Rota" dropdown menu in Google Sheets works because of these exact configurations:

#### 1. **onOpen Function** (Code.js)
```javascript
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Staff Rota')  // ⚠️ CRITICAL: Uses createMenu() NOT createAddonMenu()
      .addItem('🔄 Generate Monthly Rota', 'showRotaGenerator')
      .addSeparator()
      .addItem('🏥 Log New Absence', 'showAbsenceLogger')
      .addItem('📊 View Absence Reports', 'showAbsenceReports')
      .addSeparator()
      .addItem('⚙️ Settings', 'showSettings')
      .addSeparator()
      .addItem('❓ Help & Documentation', 'showHelp')
      .addItem('🔒 Privacy Policy', 'showPrivacyPolicy')
      .addToUi();
```

#### 2. **appsscript.json Configuration**
- ✅ Has proper `oauthScopes`
- ✅ Has `addOns` section with both `common` and `sheets` triggers
- ✅ Uses `onFileScopeGrantedTrigger` pointing to `onAuthorizationRequired`

#### 3. **Required Functions**
- `onOpen()` - Creates the menu (NEVER REMOVE)
- `onHomepage()` - For add-on homepage
- `onAuthorizationRequired()` - For file scope granted trigger

### 🚨 **FORBIDDEN CHANGES:**

1. **NEVER change `createMenu()` to `createAddonMenu()`** - This breaks the menu
2. **NEVER remove the `onOpen()` function**
3. **NEVER modify the appsscript.json `addOns` section**
4. **NEVER add complex trigger setups like `setupTriggers()`**
5. **NEVER add try/catch blocks or console.log to onOpen()** - Keep it simple

### ✅ **Safe Changes:**
- Menu item labels and emojis
- Function names the menu items call
- HTML file contents
- Add new functions (but don't modify core menu functions)
- Backend functionality

### 📍 **Current Working State:**
- Commit: `9191509c5f319f55139c8a2c24de6237236de8b7`
- Menu appears as "Staff Rota" in Google Sheets menu bar
- Container-bound script with proper triggers and permissions

## 🚀 **Deployment & Version Control**

### 📂 **Working Directory:**
Always work from: `/mnt/c/Users/leigh/OneDrive/Coding Projects/Rota-Manager-main/target-project`

### 🔧 **Google Apps Script Deployment:**
```bash
# Navigate to target-project directory
cd "/mnt/c/Users/leigh/OneDrive/Coding Projects/Rota-Manager-main/target-project"

# Push to Google Apps Script
cmd.exe /c "clasp push --force"

# Pull from Google Apps Script (if needed)
cmd.exe /c "clasp pull"
```

**Apps Script Project URL:**
https://script.google.com/u/0/home/projects/1uhijHgBKzTFLhBMU72KlR4cFBGmi1IDSlklU7yz0MmRqJ5_hX9K-8qrL/edit

### 📝 **GitHub Repository Management:**
```bash
# Navigate to main project directory  
cd "/mnt/c/Users/leigh/OneDrive/Coding Projects/Rota-Manager-main"

# Pull latest changes from GitHub
git pull origin main

# Stage all changes
git add .

# Commit changes
git commit -m "Your commit message

🤖 Generated with Claude Code

Co-Authored-By: Claude <noreply@anthropic.com>"

# Push to GitHub (create new branch for features)
git checkout -b feature-branch-name
git push origin feature-branch-name

# Or push to main branch
git checkout main
git push origin main
```

**GitHub Repository:**
https://github.com/Leighann807/Rota-Manager.git

### ⚙️ **Git Configuration:**
The repository is configured with Windows Git credential manager:
```bash
git config credential.helper "/mnt/c/Program Files/Git/mingw64/bin/git-credential-manager.exe"
```

### 🔄 **Standard Workflow:**
1. **Make changes** in `/target-project/` directory
2. **Test changes** in Google Sheets
3. **Deploy to Apps Script:** `clasp push --force`
4. **Commit to Git:** Add, commit, and push to GitHub
5. **Create Pull Request** if using feature branches

## 📂 **Project Structure**

### HTML Files:
- `Settings.html` - Staff management interface
- `RotaGenerator.html` - Monthly rota creation
- `AbsenceLogger.html` - Log staff absences
- `AbsenceReports.html` - View absence statistics
- `Help.html` - Documentation
- `PrivacyPolicy.html` - Privacy information
- `Homepage.html` - Add-on homepage
- `SettingsDialog.html` - Settings dialog

### Core Files:
- `Code.js` - Main backend functions
- `appsscript.json` - Add-on manifest
- `.clasp.json` - Deployment configuration

## 🎯 **Key Features**
- Monthly rota generation with shift patterns
- Staff absence tracking and reporting
- Settings management for staff and preferences
- Privacy policy and data management
- Comprehensive help documentation

---
**Last Updated:** After successful revert to working commit
**Status:** ✅ Menu working correctly