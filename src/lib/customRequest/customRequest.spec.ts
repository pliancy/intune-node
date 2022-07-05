import { Client } from '@microsoft/microsoft-graph-client'
import { mockClient } from '../../../__fixtures__/@microsoft/microsoft-graph-client'
import { CustomRequest } from './customRequest'

describe('Custom Request', () => {
    let graphClient: Client
    let customRequest: CustomRequest

    beforeEach(() => {
        graphClient = mockClient() as never as Client
        customRequest = new CustomRequest(graphClient)
    })

    it('should make a get request', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue({})
        const result = await customRequest.get('')
        expect(result).toEqual({})
    })

    it('should make a post request', async () => {
        const postSpy = jest.spyOn(graphClient.api(''), 'post').mockResolvedValue({})
        const apiSpy = jest.spyOn(graphClient, 'api')
        const result = await customRequest.post('', {})
        expect(result).toEqual({})
        expect(apiSpy).toHaveBeenCalledWith('')
        expect(postSpy).toHaveBeenCalledWith({})
    })

    it('should make a patch request', async () => {
        const patchSpy = jest.spyOn(graphClient.api(''), 'patch').mockResolvedValue({})
        const apiSpy = jest.spyOn(graphClient, 'api')
        const result = await customRequest.patch('', {})
        expect(result).toEqual({})
        expect(apiSpy).toHaveBeenCalledWith('')
        expect(patchSpy).toHaveBeenCalledWith({})
    })

    it('should make a delete request', async () => {
        jest.spyOn(graphClient.api(''), 'delete').mockResolvedValue({})
        const apiSpy = jest.spyOn(graphClient, 'api')
        const result = await customRequest.delete('')
        expect(result).toEqual({})
        expect(apiSpy).toHaveBeenCalledWith('')
    })

    it('should make a put request', async () => {
        const putSpy = jest.spyOn(graphClient.api(''), 'put').mockResolvedValue({})
        const apiSpy = jest.spyOn(graphClient, 'api')
        const result = await customRequest.put('', {})
        expect(result).toEqual({})
        expect(apiSpy).toHaveBeenCalledWith('')
        expect(putSpy).toHaveBeenCalledWith({})
    })
})
