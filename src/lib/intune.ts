import { Config } from './types'
import { Client } from '@microsoft/microsoft-graph-client'
import { Devices } from './devices/devices'
import { DeviceConfigurations } from './deviceConfigurations/deviceConfigurations'
import { DeviceManagementScripts } from './deviceManagementScripts/deviceManagementScripts'
import { MobileApps } from './mobileApps/mobileApps'
import { Groups } from './groups/groups'
import { Users } from './users/users'
import { GroupPolicyConfigurations } from './groupPolicyConfigurations/groupPolicyConfigurations'
import { CustomRequest } from './customRequest/customRequest'
import { Autopilot } from './autopilot/autopilot'
import { DeviceManagementTemplates } from './deviceManagementTemplates/deviceManagementTemplates'
import { DeviceManagementIntents } from './deviceManagementIntents/deviceManagementIntents'
import { DeviceHealthScripts } from './deviceHealthScripts/deviceHealthScripts'
import { AuthProvider } from './utils/auth-provider'
import { DeviceConfigurationPolicies } from './deviceConfigurationPolicies/deviceConfigurationPolicies'
import { DeviceShellScripts } from './deviceShellScripts/deviceShellScripts'
require('isomorphic-fetch')

export class Intune {
    readonly graphclient: Client

    readonly autopilot: Autopilot

    readonly devices: Devices

    readonly deviceConfigurations: DeviceConfigurations

    readonly deviceHealthScripts: DeviceHealthScripts

    readonly deviceManagementScripts: DeviceManagementScripts

    readonly mobileApps: MobileApps

    readonly groups: Groups

    readonly groupPolicyConfigurations: GroupPolicyConfigurations

    readonly users: Users

    readonly customRequest: CustomRequest

    readonly deviceManagementTemplates: DeviceManagementTemplates

    readonly deviceManagementIntents: DeviceManagementIntents

    readonly deviceConfigurationPolicies: DeviceConfigurationPolicies

    readonly deviceShellScripts: DeviceShellScripts

    private readonly authProvider: AuthProvider

    constructor(private readonly config: Config) {
        this.authProvider = new AuthProvider(this.config)

        this.graphclient = Client.initWithMiddleware({
            authProvider: this.authProvider,
            defaultVersion: 'beta',
        })

        this.devices = new Devices(this.graphclient)
        this.deviceConfigurations = new DeviceConfigurations(this.graphclient)
        this.deviceManagementScripts = new DeviceManagementScripts(this.graphclient)
        this.mobileApps = new MobileApps(this.graphclient)
        this.groups = new Groups(this.graphclient)
        this.groupPolicyConfigurations = new GroupPolicyConfigurations(this.graphclient)
        this.users = new Users(this.graphclient)
        this.customRequest = new CustomRequest(this.graphclient)
        this.autopilot = new Autopilot(this.graphclient)
        this.deviceManagementTemplates = new DeviceManagementTemplates(this.graphclient)
        this.deviceManagementIntents = new DeviceManagementIntents(this.graphclient)
        this.deviceHealthScripts = new DeviceHealthScripts(this.graphclient)
        this.deviceConfigurationPolicies = new DeviceConfigurationPolicies(this.graphclient)
        this.deviceShellScripts = new DeviceShellScripts(this.graphclient)
    }

    /**
     * Get the current refresh token
     * @returns {Promise<string>} The current refresh token
     */
    get refreshToken(): string | undefined {
        return this.authProvider.refreshToken
    }
}
