import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceManagementIntent } from 'lib/types'
import { mockClient } from '../../../__fixtures__/@microsoft/microsoft-graph-client'
import { DeviceManagementIntents } from './deviceManagementIntents'

describe('DeviceManagementIntents', () => {
    let graphClient: Client
    let deviceManagementIntents: DeviceManagementIntents
    const intent = {
        displayName: 'test',
        id: 'test',
    } as DeviceManagementIntent

    beforeEach(() => {
        graphClient = mockClient() as never as Client
        deviceManagementIntents = new DeviceManagementIntents(graphClient)
    })

    it('should get a intent', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(intent)
        const result = await deviceManagementIntents.get('')
        expect(result).toEqual(intent)
    })

    it('should list all intents', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [intent],
            '@odata.nextLink': 'next',
        })
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [intent],
        })

        const result = await deviceManagementIntents.list()
        expect(result).toEqual([intent, intent])
    })

    it('should update a intent', async () => {
        const postSpy = jest.spyOn(graphClient.api(''), 'patch').mockResolvedValue(intent)
        const result = await deviceManagementIntents.update('id', intent)
        expect(result).toEqual(intent)
        expect(postSpy).toHaveBeenCalledWith(intent)
    })

    it('should delete a intent', async () => {
        const spy = jest.spyOn(graphClient, 'api')
        jest.spyOn(graphClient.api(''), 'patch')
        const result = await deviceManagementIntents.delete('id')
        expect(result).toBeUndefined()
        expect(spy).toHaveBeenCalledWith('/deviceManagement/intents/id')
    })

    it('should create a intent', async () => {
        const apiSpy = jest.spyOn(graphClient, 'api')
        const postSpy = jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(intent)
        const result = await deviceManagementIntents.create(intent)
        expect(result).toEqual(intent)
        expect(apiSpy).toHaveBeenCalledWith('/deviceManagement/intents')
        expect(postSpy).toHaveBeenCalledWith(intent)
    })
})
