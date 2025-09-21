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

### 🔧 **Deployment Commands:**
```bash
cd "/mnt/c/Users/leigh/OneDrive/Coding Projects/Rota-Manager-main/target-project"
cmd.exe /c "clasp push --force"
```

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