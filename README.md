# Microsoft Intune SDK

[Full Documentation](https://pliancy.github.io/intune-node/) 

## Getting Started

You can install the package with the following command:

```javascript
npm install microsoft-intune
```

```javascript
yarn add microsoft-intune
```

Import the package 

```javascript
import { Intune } from 'microsoft-intune'
```

Initialize with Client ID and Client Secret Auth:


```javascript
  const intune = Intune.init({
    authentication: {
      clientId: '',
      clientSecret: ''
    },
    tenantId: ''
  })
```



## Example Usage

### Get all Intune Devices

```javascript
await intune.devices.list()
```

### Get all AzureAd Devices

```javascript
await intune.devices.listAzureAdDevices()
```

### Get all Device Configurations

```javascript
await intune.deviceConfigurations.list()
```

### Create Device Configuration

```javascript
const deviceConfiguration = {
  "@odata.type": "#microsoft.graph.windows10GeneralConfiguration",
  "description": "Disables sleep when lid is closed",
  "displayName": "Power - Disable Sleep for Lid Close",
  "powerLidCloseActionPluggedIn": "noAction"
}

await intune.deviceConfigurations.create(deviceConfiguration)
```

### Create Office Suite App

```javascript
const officeApp = { 
  "@odata.type": "#microsoft.graph.officeSuiteApp",
  "displayName": "Office 365",
  "description": "Office 365 for Windows 10",
  "publisher": "Microsoft",
  "largeIcon": null,
  "isFeatured": true,
  "privacyInformationUrl": "https://privacy.microsoft.com/en-US/privacystatement",
  "informationUrl": "https://products.office.com/en-us/explore-office-for-home",
  "owner": "Microsoft",
  "developer": "Microsoft",
  "notes": "",
  "roleScopeTagIds": [],
  "autoAcceptEula": true,
  "productIds": [
    "o365ProPlusRetail"
  ],
  "useSharedComputerActivation": false,
  "updateChannel": "deferred",
  "officePlatformArchitecture": "x64",
  "localesToInstall": [],
  "installProgressDisplayLevel": "none",
  "shouldUninstallOlderVersionsOfOffice": true,
  "targetVersion": "",
  "updateVersion": "",
  "officeConfigurationXml": null,
  "excludedApps": {
    "access": true,
    "excel": false,
    "groove": true,
    "infoPath": true,
    "lync": true,
    "oneDrive": true,
    "oneNote": false,
    "outlook": false,
    "powerPoint": false,
    "publisher": true,
    "sharePointDesigner": true,
    "teams": true,
    "visio": true,
    "word": false
  }
}

await intune.mobileApps.create(officeApp)
```

### Create and Upload Win32 App from Stream

This function requires the mobileApp Info, fileEncryptionInfo, mobileAppContentFile,  and the unencrypted .intunewin file . Some info for these objects is found in the detection.xml that's located in the extracted .intunewin file.

```javascript
const mobileApp = {
  '@odata.type': '#microsoft.graph.win32LobApp',
  displayName: 'App',
  description: '',
  publisher: 'Publisher',
  isFeatured: true,
  privacyInformationUrl: '',
  informationUrl: null,
  owner: '',
  developer: '',
  notes: '',
  fileName: 'app.intunewin',
  installCommandLine: 'install.cmd',
  uninstallCommandLine: 'uninstall.cmd',
  applicableArchitectures: 'x64',
  minimumFreeDiskSpaceInMB: null,
  minimumMemoryInMB: null,
  minimumNumberOfProcessors: null,
  minimumCpuSpeedInMHz: null,
  msiInformation: null,
  setupFilePath: 'app.exe',
  largeIcon: {
    type: 'image/png',
    value: 'keejejejejenenbejdejdn...'
  },
  minimumSupportedOperatingSystem: {
    v8_0: false,
    v8_1: false,
    v10_0: false,
    v10_1607: true,
    v10_1703: false,
    v10_1709: false,
    v10_1803: false,
    v10_1809: false,
    v10_1903: false
  },
  detectionRules: [
    {
      '@odata.type': '#microsoft.graph.win32LobAppFileSystemDetection',
      path: '%ProgramFiles%\\App',
      fileOrFolderName: 'App.exe',
      check32BitOn64System: false,
      detectionType: 'exists',
      operator: 'notConfigured',
      detectionValue: null
    }
  ],
  requirementRules: [],
  installExperience: {
    runAsAccount: 'system',
    deviceRestartBehavior: 'suppress'
  },
  returnCodes: [
    {
      returnCode: 0,
      type: 'success'
    },
    {
      returnCode: 1,
      type: 'failed'
    }
  ]
}

const fileEncryptionInfo = {
    fileDigestAlgorithm: 'SHA256',
    encryptionKey: 'BKu4^YNmrrfG74yT3R&qAly',
    initializationVector: 'BKu4^YNmrrfG74yT3R&qAly',
    fileDigest: 'BKu4^YNmrrfG74yT3R&qAly',
    mac: 'BKu4^YNmrrfG74yT3R&qAly',
    profileIdentifier: 'ProfileVersion1',
    macKey: 'BKu4^YNmrrfG74yT3R&qAly'
  }

const mobileAppContentFile = {
  '@odata.type': '#microsoft.graph.mobileAppContentFile',
  manifest: null,
  size: 3332,
  name: 'app.intunewin',
  sizeEncrypted: 3993,
  isDependency: false
}

await intune.createWin32LobApp(mobileApp, fileEncryptionInfo, mobileAppContentFile, unencryptedFile) 
```

### Custom Request

```javascript
await intune.customRequest.get('/endpoint')
```

### 