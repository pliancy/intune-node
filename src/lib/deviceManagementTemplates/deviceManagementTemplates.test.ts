import { mockClient } from '../../../__fixtures__/@microsoft/microsoft-graph-client'
import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceManagementTemplates } from './deviceManagementTemplates'
import { CreateTemplateInstance, DeviceManagementTemplate } from '../types'

describe('Devices Management Templates', () => {
    let graphClient: Client
    let deviceManagementTemplates: DeviceManagementTemplates

    const deviceManagementTemplate = {
        name: 'test',
        '@odata.type': '#microsoft.graph.deviceManagementTemplate',
        id: '1',
    }

    beforeEach(() => {
        graphClient = mockClient() as never as Client
        deviceManagementTemplates = new DeviceManagementTemplates(graphClient)
    })

    it('should list device management templates', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue({
            value: [deviceManagementTemplate],
        })
        const res = await deviceManagementTemplates.list()
        expect(res).toEqual([deviceManagementTemplate])
    })

    it('should get a device management template', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(deviceManagementTemplate)
        const res = await deviceManagementTemplates.get('')
        expect(res).toEqual(deviceManagementTemplate)
    })

    it('should create a device management template', async () => {
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(deviceManagementTemplate)
        const res = await deviceManagementTemplates.create({} as DeviceManagementTemplate)
        expect(res).toEqual(deviceManagementTemplate)
    })

    it('should update a device management template', async () => {
        jest.spyOn(graphClient.api(''), 'patch')
        const res = await deviceManagementTemplates.update('id', {} as DeviceManagementTemplate)
        expect(res).toBeUndefined()
    })

    it('should delete a device management template', async () => {
        jest.spyOn(graphClient.api(''), 'delete')
        const res = await deviceManagementTemplates.delete('id')
        expect(res).toBeUndefined()
    })

    it('should create create an instance ', async () => {
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(deviceManagementTemplate)
        const res = await deviceManagementTemplates.createInstance(
            '1',
            {} as CreateTemplateInstance,
        )
        expect(res).toEqual(deviceManagementTemplate)
    })
})
