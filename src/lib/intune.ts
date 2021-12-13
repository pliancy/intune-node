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

    private constructor(private readonly authProvider: TokenCredentialAuthenticationProvider) {
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
    }

    static init(config: Config) {
        const credential = new ClientSecretCredential(
            config.tenantId,
            config.authentication.clientId,
            config.authentication.clientSecret,
        )
        const authProvider = new TokenCredentialAuthenticationProvider(credential, {
            scopes: ['.default'],
        })
        return new Intune(authProvider)
    }
}
