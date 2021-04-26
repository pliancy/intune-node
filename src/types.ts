export interface IntuneConfig {
  tenantId: string
  authentication: ClientAuth | BearerAuth
}

export interface ClientAuth {
  clientId: string
  clientSecret: string
}

export interface BearerAuth {
  bearerToken: string
}
export interface IOAuthResponse {
  token_type: string
  expires_in: string
  ext_expires_in: string
  expires_on: string
  not_before: string
  resource: string
  access_token: string
}

export interface IntuneScript {
  displayName: string
  description: string
  scriptContent: string
  runAsAccount: 'system' | 'user'
  enforceSignatureCheck: boolean
  fileName: string
  runAs32Bit: boolean
}

export interface AutoPilotUpload {
  serialNumber?: string
  groupTag?: string
  productKey?: string
  hardwareIdentifier?: string
  assignedUser?: string
}

export interface IntuneDeviceResponse {
  id: string
  userId: string
  deviceName: string
  ownerType: string
  managedDeviceOwnerType: string
  managementState: string
  enrolledDateTime: string
  lastSyncDateTime: string
  chassisType: string
  operatingSystem: string
  deviceType: string
  complianceState: string
  jailBroken: string
  managementAgent: string
  osVersion: string
  easActivated: boolean
  easDeviceId: string
  easActivationDateTime: string
  aadRegistered: boolean
  azureADRegistered: boolean
  deviceEnrollmentType: string
  lostModeState: string
  activationLockBypassCode?: unknown
  emailAddress: string
  azureActiveDirectoryDeviceId: string
  azureADDeviceId: string
  deviceRegistrationState: string
  deviceCategoryDisplayName: string
  isSupervised: boolean
  exchangeLastSuccessfulSyncDateTime: string
  exchangeAccessState: string
  exchangeAccessStateReason: string
  remoteAssistanceSessionUrl?: unknown
  remoteAssistanceSessionErrorDetails?: unknown
  isEncrypted: boolean
  userPrincipalName: string
  model: string
  manufacturer: string
  imei: string
  complianceGracePeriodExpirationDateTime: string
  serialNumber: string
  phoneNumber: string
  androidSecurityPatchLevel: string
  userDisplayName: string
  configurationManagerClientEnabledFeatures?: unknown
  wiFiMacAddress: string
  deviceHealthAttestationState?: unknown
  subscriberCarrier: string
  meid: string
  totalStorageSpaceInBytes: number
  freeStorageSpaceInBytes: number
  managedDeviceName: string
  partnerReportedThreatState: string
  retireAfterDateTime: string
  preferMdmOverGroupPolicyAppliedDateTime: string
  autopilotEnrolled: boolean
  requireUserEnrollmentApproval?: unknown
  managementCertificateExpirationDate: string
  iccid?: unknown
  udid?: unknown
  roleScopeTagIds: unknown[]
  windowsActiveMalwareCount: number
  windowsRemediatedMalwareCount: number
  notes?: unknown
  configurationManagerClientHealthState?: unknown
  configurationManagerClientInformation?: unknown
  ethernetMacAddress?: unknown
  physicalMemoryInBytes: number
  processorArchitecture: string
  specificationVersion?: unknown
  joinType: string
  skuFamily: string
  skuNumber: number
  managementFeatures: string
  hardwareInformation: {
    serialNumber: string
    totalStorageSpace: number
    freeStorageSpace: number
    imei: string
    meid?: unknown
    manufacturer?: unknown
    model?: unknown
    phoneNumber?: unknown
    subscriberCarrier?: unknown
    cellularTechnology?: unknown
    wifiMac?: unknown
    operatingSystemLanguage?: unknown
    isSupervised: boolean
    isEncrypted: boolean
    batterySerialNumber?: unknown
    batteryHealthPercentage: number
    batteryChargeCycles: number
    isSharedDevice: boolean
    tpmSpecificationVersion?: unknown
    operatingSystemEdition?: unknown
    deviceFullQualifiedDomainName?: unknown
    deviceGuardVirtualizationBasedSecurityHardwareRequirementState: string
    deviceGuardVirtualizationBasedSecurityState: string
    deviceGuardLocalSystemAuthorityCredentialGuardState: string
    osBuildNumber?: unknown
    operatingSystemProductType: number
    ipAddressV4?: unknown
    subnetAddress?: unknown
    sharedDeviceCachedUsers: []
  }
  deviceActionResults: unknown[]
  usersLoggedOn: unknown[]
  trustType?: string
}

export interface AzureDeviceResponse {
  id: string
  deletedDateTime: Date | string
  accountEnabled: boolean
  approximateLastSignInDateTime: Date | string
  complianceExpirationDateTime: Date | string
  createdDateTime: Date | string
  deviceCategory?: string
  deviceId: string
  deviceMetadata?: unknown
  deviceOwnership: string
  deviceVersion: number
  displayName: string
  domainName?: string
  enrollmentProfileName?: string
  enrollmentType: string
  externalSourceName?: string
  isCompliant: boolean
  isManaged: boolean
  isManagementRestricted?: boolean
  isRooted: boolean
  managementType: string
  manufacturer: string
  mdmAppId: string
  model: string
  onPremisesLastSyncDateTime?: Date | string
  onPremisesSyncEnabled?: boolean
  operatingSystem: string
  operatingSystemVersion: string
  hostnames: string[]
  physicalIds: string[]
  profileType: string
  registrationDateTime: Date | string
  sourceType?: string
  systemLabels: string[]
  trustType: string
  alternativeSecurityIds: string[]
  extensionAttributes: {
    extensionAttribute1?: string
    extensionAttribute2?: string
    extensionAttribute3?: string
    extensionAttribute4?: string
    extensionAttribute5?: string
    extensionAttribute6?: string
    extensionAttribute7?: string
    extensionAttribute8?: string
    extensionAttribute9?: string
    extensionAttribute10?: string
    extensionAttribute11?: string
    extensionAttribute12?: string
    extensionAttribute13?: string
    extensionAttribute14?: string
    extensionAttribute15?: string
  }
}
