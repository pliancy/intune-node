import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceConfigurations } from './deviceConfigurations'
import { mockClient } from '../../test/mocks/@microsoft/microsoft-graph-client'

describe('Device Configurations', () => {
    let deviceConfigurations: DeviceConfigurations
    const deviceConfiguration = {
        name: 'test',
        '@odata.type': '#microsoft.graph.deviceConfiguration',
    }

    const groupAssignment = {
        '@odata.type': '#microsoft.graph.deviceManagementScriptGroupAssignment',
        targetGroupId: '1',
    }

    it('should get a device configuration', async () => {
        const graphClient = mockClient() as never as Client
        deviceConfigurations = new DeviceConfigurations(graphClient)
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(deviceConfiguration)
        const result = await deviceConfigurations.get('')
        expect(result).toEqual(deviceConfiguration)
    })

    it('should list all device configurations', async () => {
        const graphClient = mockClient() as never as Client
        deviceConfigurations = new DeviceConfigurations(graphClient)
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [deviceConfiguration],
            '@odata.nextLink': 'next',
        })
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [deviceConfiguration],
        })

        const result = await deviceConfigurations.list()
        expect(result).toEqual([deviceConfiguration, deviceConfiguration])
    })

    it('should create a device configuration', async () => {
        const graphClient = mockClient() as never as Client
        deviceConfigurations = new DeviceConfigurations(graphClient)
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(deviceConfiguration)
        const result = await deviceConfigurations.create(deviceConfiguration)
        expect(result).toEqual(deviceConfiguration)
    })

    it('should update a device configuration', async () => {
        const graphClient = mockClient() as never as Client
        deviceConfigurations = new DeviceConfigurations(graphClient)
        jest.spyOn(graphClient.api(''), 'patch')
        const result = await deviceConfigurations.update('id', deviceConfiguration)
        expect(result).toBeUndefined()
    })

    it('should delete a device configuration', async () => {
        const graphClient = mockClient() as never as Client
        deviceConfigurations = new DeviceConfigurations(graphClient)
        jest.spyOn(graphClient.api(''), 'delete')
        const result = await deviceConfigurations.delete('id')
        expect(result).toBeUndefined()
    })

    it('should create a group assignment', async () => {
        const graphClient = mockClient() as never as Client
        deviceConfigurations = new DeviceConfigurations(graphClient)
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(groupAssignment)
        const spy = jest.spyOn(graphClient, 'api')
        const result = await deviceConfigurations.createGroupAssignment('id', 'groupId')
        expect(result).toEqual(groupAssignment)
        expect(spy).toHaveBeenCalledWith(
            '/deviceManagement/deviceConfigurations/id/groupAssignments',
        )
    })

    it('should list group assignments', async () => {
        const graphClient = mockClient() as never as Client
        deviceConfigurations = new DeviceConfigurations(graphClient)
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue({ value: [groupAssignment] })
        const result = await deviceConfigurations.listGroupAssignments('id')
        expect(result).toEqual([groupAssignment])
    })

    it('should get a group assignment', async () => {
        const graphClient = mockClient() as never as Client
        deviceConfigurations = new DeviceConfigurations(graphClient)
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(groupAssignment)
        const result = await deviceConfigurations.getGroupAssignment('id', 'groupAssignmentId')
        expect(result).toEqual(groupAssignment)
    })

    it('should delete a group assignment', async () => {
        const graphClient = mockClient() as never as Client
        deviceConfigurations = new DeviceConfigurations(graphClient)
        jest.spyOn(graphClient.api(''), 'delete')
        const result = await deviceConfigurations.deleteGroupAssignment('id', 'groupId')
        expect(result).toBeUndefined()
    })
})
