import { Devices } from './devices/devices'
import { DeviceConfigurations } from './deviceConfigurations/deviceConfigurations'
import { DeviceManagementScripts } from './deviceManagementScripts/deviceManagementScripts'
import { MobileApps } from './mobileApps/mobileApps'
import { Intune } from './intune'
import { GroupPolicyConfigurations } from './groupPolicyConfigurations/groupPolicyConfigurations'
import { Groups } from './groups/groups'
import { Users } from './users/users'
import { DeviceManagementTemplates } from './deviceManagementTemplates/deviceManagementTemplates'
import { DeviceManagementIntents } from './deviceManagementIntents/deviceManagementIntents'
import { Autopilot } from './autopilot/autopilot'
import { DeviceConfigurationPolicies } from './deviceConfigurationPolicies/deviceConfigurationPolicies'

jest.mock('./utils/auth-provider.ts', () => {
    return {
        AuthProvider: jest.fn().mockImplementation(() => {
            return {
                refreshToken: 'mockedRefreshToken',
            }
        }),
    }
})
describe('Intune', () => {
    it('creates component instances', () => {
        const intune = new Intune({
            tenantId: 'tenantId',
            authentication: { clientId: 'clientId', clientSecret: 'clientSecret' },
        })

        expect(intune.devices).toBeInstanceOf(Devices)
        expect(intune.deviceConfigurations).toBeInstanceOf(DeviceConfigurations)
        expect(intune.deviceManagementScripts).toBeInstanceOf(DeviceManagementScripts)
        expect(intune.groupPolicyConfigurations).toBeInstanceOf(GroupPolicyConfigurations)
        expect(intune.groups).toBeInstanceOf(Groups)
        expect(intune.mobileApps).toBeInstanceOf(MobileApps)
        expect(intune.users).toBeInstanceOf(Users)
        expect(intune.customRequest).toBeDefined()
        expect(intune.deviceManagementTemplates).toBeInstanceOf(DeviceManagementTemplates)
        expect(intune.deviceManagementIntents).toBeInstanceOf(DeviceManagementIntents)
        expect(intune.autopilot).toBeInstanceOf(Autopilot)
        expect(intune.deviceConfigurationPolicies).toBeInstanceOf(DeviceConfigurationPolicies)
    })

    // Test for the refreshToken getter
    it('returns the mocked refresh token', () => {
        const intune = new Intune({
            tenantId: 'tenantId',
            authentication: {
                clientId: 'clientId',
                clientSecret: 'clientSecret',
                refreshToken: 'mockedRefreshToken',
            },
        })

        expect(intune.refreshToken).toBe('mockedRefreshToken')
    })
})
