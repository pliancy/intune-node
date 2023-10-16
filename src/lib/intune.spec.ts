import { Devices } from './devices/devices'
import { DeviceConfigurations } from './deviceConfigurations/deviceConfigurations'
import { DeviceManagementScripts } from './deviceManagementScripts/deviceManagementScripts'
import { MobileApps } from './mobileApps/mobileApps'
import { Intune } from './intune'
import { GroupPolicyConfigurations } from './groupPolicyConfigurations/groupPolicyConfigurations'
import { Groups } from './groups/groups'
import { Users } from './users/users'

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
    })
})
