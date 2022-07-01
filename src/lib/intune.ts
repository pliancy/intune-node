import { Config } from './types'
import { Client } from '@microsoft/microsoft-graph-client'
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials'
import { ClientSecretCredential } from '@azure/identity'
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
require('isomorphic-fetch')

export class Intune {
    readonly graphclient: Client

    readonly autopilot: Autopilot

    readonly devices: Devices

    readonly deviceConfigurations: DeviceConfigurations

    readonly deviceManagementScripts: DeviceManagementScripts

    readonly mobileApps: MobileApps

    readonly groups: Groups

    readonly groupPolicyConfigurations: GroupPolicyConfigurations

    readonly users: Users

    readonly customRequest: CustomRequest

    readonly deviceManagementTemplates: DeviceManagementTemplates

    constructor(private readonly config: Config) {
        const credential = new ClientSecretCredential(
            this.config.tenantId,
            this.config.authentication.clientId,
            this.config.authentication.clientSecret,
        )
        const authProvider = new TokenCredentialAuthenticationProvider(credential, {
            scopes: ['.default'],
        })

        this.graphclient = Client.initWithMiddleware({
            authProvider,
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
    }
}
