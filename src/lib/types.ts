import * as Graph from '@microsoft/microsoft-graph-types-beta'

export interface Config {
    tenantId: string
    authentication: ClientAuth
}

export interface ClientAuth {
    clientId: string
    clientSecret: string
}

export interface AutoPilotUpload {
    serialNumber?: string
    groupTag?: string
    productKey?: string
    hardwareIdentifier?: string
    assignedUser?: string
}

export interface DeviceConfiguration extends Graph.DeviceConfiguration {
    '@odata.type': string
}

export interface DeviceManagementScript extends Graph.DeviceManagementScript {
    '@odata.type': string
}

export interface MobileAppContentFile extends Graph.MobileAppContentFile {
    '@odata.type': '#microsoft.graph.mobileAppContentFile'
    size: number
    sizeEncrypted: number
}

export type MobileApp =
    | AndroidForWorkApp
    | AndroidManagedStoreApp
    | AndroidStoreApp
    | MacOSMicrosoftEdgeApp
    | MacOsVppApp
    | MacOSMicrosoftEdgeApp
    | ManagedIOSStoreApp
    | IosStoreApp
    | Win32LobApp
    | MicrosoftStoreForBusinessApp
    | OfficeSuiteApp
    | WebApp
    | WindowsMicrosoftEdgeApp
    | WindowsStoreApp
    | IosLobApp
    | MacOSLobApp
    | WindowsAppX
    | WindowsUniversalAppX
    | AndroidLobApp
    | WindowsMobileMSI

export interface AndroidForWorkApp extends Graph.AndroidForWorkApp {
    '@odata.type': '#microsoft.graph.androidForWorkApp'
}

export interface AndroidManagedStoreApp extends Graph.AndroidManagedStoreApp {
    '@odata.type': '#microsoft.graph.managedAndroidStoreApp'
}

export interface AndroidStoreApp extends Graph.AndroidStoreApp {
    '@odata.type': '#microsoft.graph.androidStoreApp'
}

export interface IosStoreApp extends Graph.IosStoreApp {
    '@odata.type': '#microsoft.graph.managedIOSStoreApp'
}

export interface IosVppApp extends Graph.IosVppApp {
    '@odata.type': '#microsoft.graph.iosVppApp'
}

export interface ManagedIOSStoreApp extends Graph.ManagedIOSStoreApp {
    '@odata.type': '#microsoft.graph.managedIOSStoreApp'
}

export interface MacOSMicrosoftEdgeApp extends Graph.MacOSMicrosoftEdgeApp {
    '@odata.type': '#microsoft.graph.macOSMicrosoftEdgeApp'
}

export interface MacOsVppApp extends Graph.MacOsVppApp {
    '@odata.type': '#microsoft.graph.macOsVppApp'
}

export interface MicrosoftStoreForBusinessApp extends Graph.MicrosoftStoreForBusinessApp {
    '@odata.type': '#microsoft.graph.microsoftStoreForBusinessApp'
}

export interface OfficeSuiteApp extends Graph.OfficeSuiteApp {
    '@odata.type': '#microsoft.graph.officeSuiteApp'
}

export interface WebApp extends Graph.WebApp {
    '@odata.type': '#microsoft.graph.webApp'
}

export interface WindowsMicrosoftEdgeApp extends Graph.WindowsMicrosoftEdgeApp {
    '@odata.type': '#microsoft.graph.windowsMicrosoftEdgeApp'
}

export interface WindowsStoreApp extends Graph.WindowsStoreApp {
    '@odata.type': '#microsoft.graph.windowsStoreApp'
}
export interface Win32LobApp extends Graph.Win32LobApp {
    '@odata.type': '#microsoft.graph.win32LobApp'
}

export interface IosLobApp extends Graph.IosLobApp {
    '@odata.type': '#microsoft.graph.iosLobApp'
}

export interface MacOSLobApp extends Graph.MacOSLobApp {
    '@odata.type': '#microsoft.graph.macOSLobApp'
}

export interface WindowsAppX extends Graph.WindowsAppX {
    '@odata.type': '#microsoft.graph.windowsAppX'
}

export interface WindowsUniversalAppX extends Graph.WindowsUniversalAppX {
    '@odata.type': '#microsoft.graph.windowsUniversalAppX'
}

export interface AndroidLobApp extends Graph.AndroidLobApp {
    '@odata.type': '#microsoft.graph.androidLobApp'
}

export interface WindowsMobileMSI extends Graph.WindowsMobileMSI {
    '@odata.type': '#microsoft.graph.windowsMobileMSI'
}

export interface ReadStream extends NodeJS.ReadStream {}

export interface MobileAppDependency extends Graph.MobileAppDependency {
    '@odata.type': '#microsoft.graph.mobileAppDependency'
}

export interface MobileAppSupersedence extends Graph.MobileAppSupersedence {
    '@odata.type': '#microsoft.graph.mobileAppSupersedence'
}

export type MobileAppRelationship = MobileAppDependency | MobileAppSupersedence

export interface MobileAppAssignment extends Graph.MobileAppAssignment {
    '@odata.type': '#microsoft.graph.mobileAppAssignment'
}

export interface DeviceManagementTemplate extends Graph.DeviceManagementTemplate {
    '@odata.type': '#microsoft.graph.deviceManagementTemplate'
}

export interface DeviceManagementSettingInstance extends Graph.DeviceManagementSettingInstance {
    '@odata.type': '#microsoft.graph.deviceManagementSettingInstance'
}

export interface CreateTemplateInstance {
    displayName: string
    description: string
    settingsDelta: DeviceManagementSettingInstance[]
    roleScopeTagIds: string[]
}

export interface DeviceManagementIntent extends Graph.DeviceManagementIntent {
    '@odata.type': '#microsoft.graph.deviceManagementIntent'
}
