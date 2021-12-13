import { Client } from '@microsoft/microsoft-graph-client'
import { Users } from './users'
import { mockClient } from '../../../__fixtures__/@microsoft/microsoft-graph-client'
import { User } from '@microsoft/microsoft-graph-types-beta'

describe('Device Configurations', () => {
    let graphClient: Client
    let users: Users
    const user = {
        name: 'test',
        id: 'test',
    } as User

    const groupAssignment = {
        '@odata.type': '#microsoft.graph.deviceManagementScriptGroupAssignment',
        targetGroupId: '1',
    }
    beforeEach(() => {
        graphClient = mockClient() as never as Client
        users = new Users(graphClient)
    })

    it('should get a user', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(user)
        const result = await users.get('')
        expect(result).toEqual(user)
    })

    it('should list all users', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [users],
            '@odata.nextLink': 'next',
        })
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [users],
        })

        const result = await users.list()
        expect(result).toEqual([users, users])
    })

    it('should update a user', async () => {
        const postSpy = jest.spyOn(graphClient.api(''), 'patch').mockResolvedValue(user)
        const result = await users.update('id', user)
        expect(result).toEqual(user)
        expect(postSpy).toHaveBeenCalledWith(user)
    })

    it('should delete a user', async () => {
        const spy = jest.spyOn(graphClient, 'api')
        jest.spyOn(graphClient.api(''), 'patch')
        const result = await users.delete('id')
        expect(result).toBeUndefined()
        expect(spy).toHaveBeenCalledWith('/users/id')
    })

    it('should create a user', async () => {
        const apiSpy = jest.spyOn(graphClient, 'api')
        const postSpy = jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(user)
        const result = await users.create(user)
        expect(result).toEqual(user)
        expect(apiSpy).toHaveBeenCalledWith('/users')
        expect(postSpy).toHaveBeenCalledWith(user)
    })

    it('should list all user app intent and states', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [{}],
            '@odata.nextLink': 'next',
        })
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [{}],
        })

        const result = await users.listAppIntentAndStates('id')
        expect(result).toEqual([{}, {}])
    })
})
