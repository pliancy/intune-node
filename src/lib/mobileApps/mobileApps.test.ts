import { Client } from '@microsoft/microsoft-graph-client'
import { MobileApps } from './mobileApps'
import { mockClient } from '../../test/mocks/@microsoft/microsoft-graph-client'
import { MobileApp, MobileAppContentFile } from '../types'
require('isomorphic-fetch')

describe('Device Configurations', () => {
    let mobileApps: MobileApps
    const mobileApp = {
        name: 'test',
        '@odata.type': '#microsoft.graph.win32LobApp',
        id: '1',
    } as MobileApp

    const groupAssignment = {
        '@odata.type': '#microsoft.graph.mobileAppAssignment',
        targetGroupId: '1',
    }

    beforeEach(() => {
        jest.clearAllMocks()
    })

    it('should get a app', async () => {
        const graphClient = mockClient() as never as Client
        mobileApps = new MobileApps(graphClient)
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(mobileApp)
        const result = await mobileApps.get('')
        expect(result).toEqual(mobileApp)
    })

    it('should list apps', async () => {
        const graphClient = mockClient() as never as Client
        mobileApps = new MobileApps(graphClient)
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [mobileApp],
            '@odata.nextLink': 'next',
        })
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({ value: [mobileApp] })
        const result = await mobileApps.list()
        expect(result).toEqual([mobileApp, mobileApp])
    })

    it('should create a app', async () => {
        const graphClient = mockClient() as never as Client
        mobileApps = new MobileApps(graphClient)
        const spy = jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(mobileApp)
        const result = await mobileApps.create(mobileApp)
        expect(result).toEqual(mobileApp)
        expect(spy).toHaveBeenCalledWith(mobileApp)
    })

    it('should update a app', async () => {
        const graphClient = mockClient() as never as Client
        mobileApps = new MobileApps(graphClient)
        const spy = jest.spyOn(graphClient.api(''), 'patch').mockResolvedValue(mobileApp)
        const result = await mobileApps.update('1', mobileApp)
        expect(result).toEqual(mobileApp)
        expect(spy).toHaveBeenCalledWith(mobileApp)
    })

    it('should delete a app', async () => {
        const graphClient = mockClient() as never as Client
        mobileApps = new MobileApps(graphClient)
        jest.spyOn(graphClient.api(''), 'delete').mockResolvedValue(null)
        const spy = jest.spyOn(graphClient, 'api')
        const result = await mobileApps.delete('1')
        expect(result).toBeUndefined()
        expect(spy).toHaveBeenCalledWith('/deviceAppManagement/mobileApps/1')
    })

    it('should create a win32 lob content version', async () => {
        const graphClient = mockClient() as never as Client
        mobileApps = new MobileApps(graphClient)
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue({ id: 1 })
        const result = await mobileApps.createWin32LobContentVersion('1')
        expect(result).toEqual({ id: 1 })
    })

    it('should create a win32 lob file upload', async () => {
        const graphClient = mockClient() as never as Client
        mobileApps = new MobileApps(graphClient)
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue({ id: 1 })
        const result = await mobileApps.createWin32LobFileUpload('1', 1, {} as MobileAppContentFile)
        expect(result).toEqual({ id: 1 })
    })
})
