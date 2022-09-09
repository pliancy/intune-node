import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceHealthScript, DeviceHealthScriptAssignment } from 'lib/types'
import { mockClient } from '../../../__fixtures__/@microsoft/microsoft-graph-client'
import { DeviceHealthScripts } from './deviceHealthScripts'

describe('Device Health Scripts', () => {
    let graphClient: Client
    let deviceHealthScripts: DeviceHealthScripts
    const deviceHealthScript = {
        name: 'test',
        '@odata.type': '#microsoft.graph.deviceHealthScript',
        id: '1',
    } as DeviceHealthScript

    const assignment = {
        '@odata.type': '#microsoft.graph.deviceHealthScriptAssignment',
        targetGroupId: '1',
    } as DeviceHealthScriptAssignment

    beforeEach(() => {
        graphClient = mockClient() as never as Client
        deviceHealthScripts = new DeviceHealthScripts(graphClient)
    })

    it('should get a device health script', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(deviceHealthScript)
        const result = await deviceHealthScripts.get('')
        expect(result).toEqual(deviceHealthScript)
    })

    it('should list all device health scripts', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue({
            value: [deviceHealthScript],
        })
        const result = await deviceHealthScripts.list()
        expect(result).toEqual([deviceHealthScript])
    })

    it('should create a device health script', async () => {
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(deviceHealthScript)
        const result = await deviceHealthScripts.create(deviceHealthScript)
        expect(result).toEqual(deviceHealthScript)
    })

    it('should update a device health script', async () => {
        jest.spyOn(graphClient.api(''), 'patch')
        const result = await deviceHealthScripts.update('id', deviceHealthScript)
        expect(result).toBeUndefined()
    })

    it('should delete a device health script', async () => {
        jest.spyOn(graphClient.api(''), 'delete')
        const result = await deviceHealthScripts.delete('id')
        expect(result).toBeUndefined()
    })

    it('should create an assignment', async () => {
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(assignment)
        const result = await deviceHealthScripts.createAssignment('id', assignment)
        expect(result).toEqual(assignment)
    })

    it('should list assignments', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue({ value: [assignment] })
        const result = await deviceHealthScripts.listAssignments('id')
        expect(result).toEqual([assignment])
    })

    it('should get an assignment', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(assignment)
        const result = await deviceHealthScripts.getAssignment('id', 'groupAssignmentId')
        expect(result).toEqual(assignment)
    })

    it('should delete an assignment', async () => {
        jest.spyOn(graphClient.api(''), 'delete')
        const result = await deviceHealthScripts.deleteAssignment('id', 'groupId')
        expect(result).toBeUndefined()
    })
})
