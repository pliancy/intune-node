import { Client } from '@microsoft/microsoft-graph-client'
import { mockClient } from '../../../__fixtures__/@microsoft/microsoft-graph-client'
import { Group } from '@microsoft/microsoft-graph-types-beta'
import { Groups } from './groups'

describe('Device Configurations', () => {
    let graphClient: Client
    let groups: Groups
    const group = {
        name: 'test',
        id: 'test',
    } as Group

    const groupAssignment = {
        '@odata.type': '#microsoft.graph.deviceManagementScriptGroupAssignment',
        targetGroupId: '1',
    }

    beforeEach(() => {
        graphClient = mockClient() as never as Client
        groups = new Groups(graphClient)
    })

    it('should get a group', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(group)
        const result = await groups.get('')
        expect(result).toEqual(group)
    })

    it('should list all groups', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [groups],
            '@odata.nextLink': 'next',
        })
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [groups],
        })

        const result = await groups.list()
        expect(result).toEqual([groups, groups])
    })

    it('should update a group', async () => {
        const postSpy = jest.spyOn(graphClient.api(''), 'patch').mockResolvedValue(group)
        const result = await groups.update('id', group)
        expect(result).toEqual(group)
        expect(postSpy).toHaveBeenCalledWith(group)
    })

    it('should delete a group', async () => {
        const spy = jest.spyOn(graphClient, 'api')
        jest.spyOn(graphClient.api(''), 'patch')
        const result = await groups.delete('id')
        expect(result).toBeUndefined()
        expect(spy).toHaveBeenCalledWith('/groups/id')
    })

    it('should create a group', async () => {
        const apiSpy = jest.spyOn(graphClient, 'api')
        const postSpy = jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(group)
        const result = await groups.create(group)
        expect(result).toEqual(group)
        expect(apiSpy).toHaveBeenCalledWith('/groups')
        expect(postSpy).toHaveBeenCalledWith(group)
    })
})
