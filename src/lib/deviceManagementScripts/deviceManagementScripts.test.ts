import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceManagementScripts } from './deviceManagementScripts'
import { mockClient } from '../../../__fixtures__/@microsoft/microsoft-graph-client'

describe('Devices Managament Scripts', () => {
    let graphClient: Client
    let deviceManagementScripts: DeviceManagementScripts
    const deviceManagementScript = {
        name: 'test',
        '@odata.type': '#microsoft.graph.deviceManagementScript',
        id: '1',
    }

    const groupAssignment = {
        '@odata.type': '#microsoft.graph.deviceManagementScriptGroupAssignment',
        targetGroupId: '1',
    }

    beforeEach(() => {
        graphClient = mockClient() as never as Client
        deviceManagementScripts = new DeviceManagementScripts(graphClient)
    })

    it('should get a device management script', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(deviceManagementScript)
        const result = await deviceManagementScripts.get('')
        expect(result).toEqual(deviceManagementScript)
    })

    it('should list all device management scripts', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue({
            value: [deviceManagementScript],
        })
        const result = await deviceManagementScripts.list()
        expect(result).toEqual([deviceManagementScript])
    })

    it('should create a device management script', async () => {
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(deviceManagementScript)
        const result = await deviceManagementScripts.create(deviceManagementScript)
        expect(result).toEqual(deviceManagementScript)
    })

    it('should update a device management script', async () => {
        jest.spyOn(graphClient.api(''), 'patch')
        const result = await deviceManagementScripts.update('id', deviceManagementScript)
        expect(result).toBeUndefined()
    })

    it('should delete a device management script', async () => {
        jest.spyOn(graphClient.api(''), 'delete')
        const result = await deviceManagementScripts.delete('id')
        expect(result).toBeUndefined()
    })

    it('should create a group assignment', async () => {
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(groupAssignment)
        const result = await deviceManagementScripts.createGroupAssignment('id', 'groupId')
        expect(result).toEqual(groupAssignment)
    })

    it('should list group assignments', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue({ value: [groupAssignment] })
        const result = await deviceManagementScripts.listGroupAssignments('id')
        expect(result).toEqual([groupAssignment])
    })

    it('should get a group assignment', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(groupAssignment)
        const result = await deviceManagementScripts.getGroupAssignment('id', 'groupAssignmentId')
        expect(result).toEqual(groupAssignment)
    })

    it('should delete a group assignment', async () => {
        jest.spyOn(graphClient.api(''), 'delete')
        const result = await deviceManagementScripts.deleteGroupAssignment('id', 'groupId')
        expect(result).toBeUndefined()
    })
})
