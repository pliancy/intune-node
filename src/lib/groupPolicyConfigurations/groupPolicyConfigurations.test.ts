import { Client } from '@microsoft/microsoft-graph-client'
import { GroupPolicyConfigurations } from './groupPolicyConfigurations'
import { mockClient } from '../../../__fixtures__/@microsoft/microsoft-graph-client'
import {
    GroupPolicyConfiguration,
    GroupPolicyConfigurationAssignment,
    GroupPolicyDefinitionValue,
} from '@microsoft/microsoft-graph-types-beta'

describe('Device Configurations', () => {
    let graphClient: Client
    let groupPolicyConfigurations: GroupPolicyConfigurations
    const groupPolicyConfiguration = {
        name: 'test',
        '@odata.type': '#microsoft.graph.groupPolicyConfiguration',
        id: '1',
    } as GroupPolicyConfiguration

    const definitionValue = {
        name: 'test',
        id: '1',
    } as GroupPolicyDefinitionValue

    const groupAssignment = {
        '@odata.type': '#microsoft.graph.deviceManagementScriptGroupAssignment',
        targetGroupId: '1',
    } as GroupPolicyConfigurationAssignment

    beforeEach(() => {
        graphClient = mockClient() as never as Client
        groupPolicyConfigurations = new GroupPolicyConfigurations(graphClient)
    })

    it('should get a group policy configuration', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(groupPolicyConfiguration)
        const result = await groupPolicyConfigurations.get('')
        expect(result).toEqual(groupPolicyConfiguration)
    })

    it('should list all group policy configurations', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [groupPolicyConfiguration],
            '@odata.nextLink': 'next',
        })
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [groupPolicyConfiguration],
        })

        const result = await groupPolicyConfigurations.list()
        expect(result).toEqual([groupPolicyConfiguration, groupPolicyConfiguration])
    })

    it('should create a group policy configuration', async () => {
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(groupPolicyConfiguration)
        const result = await groupPolicyConfigurations.create(groupPolicyConfiguration)
        expect(result).toEqual(groupPolicyConfiguration)
    })

    it('should update a group policy configuration', async () => {
        jest.spyOn(graphClient.api(''), 'patch')
        const result = await groupPolicyConfigurations.update('id', groupPolicyConfiguration)
        expect(result).toBeUndefined()
    })

    it('should delete a group policy configuration', async () => {
        jest.spyOn(graphClient.api(''), 'delete')
        const result = await groupPolicyConfigurations.delete('id')
        expect(result).toBeUndefined()
    })

    it('should get a group policy configuration and definition values', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce(groupPolicyConfiguration)
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({ value: [definitionValue] })
        const result = await groupPolicyConfigurations.getWithDefinitionValues('1')
        expect(result).toEqual({ ...groupPolicyConfiguration, definitionValues: [definitionValue] })
    })

    it('should create an assignment', async () => {
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(groupAssignment)
        const spy = jest.spyOn(graphClient, 'api')
        const result = await groupPolicyConfigurations.createAssignment('id', groupAssignment)
        expect(result).toEqual(groupAssignment)
        expect(spy).toHaveBeenCalledWith(
            '/deviceManagement/groupPolicyConfigurations/id/assignments',
        )
    })

    it('should list assignments', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue({ value: [groupAssignment] })
        const result = await groupPolicyConfigurations.listAssignments('id')
        expect(result).toEqual([groupAssignment])
    })

    it('should get an assignment', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(groupAssignment)
        const result = await groupPolicyConfigurations.getAssignment('id', 'groupAssignmentId')
        expect(result).toEqual(groupAssignment)
    })

    it('should delete an assignment', async () => {
        jest.spyOn(graphClient.api(''), 'delete')
        const result = await groupPolicyConfigurations.deleteAssignment('id', 'groupId')
        expect(result).toBeUndefined()
    })
})
