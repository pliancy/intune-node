import { Client } from '@microsoft/microsoft-graph-client'
import { AutoPilot } from './autopilot'
import { mockClient } from '../../../__fixtures__/@microsoft/microsoft-graph-client'
import { AutoPilotUpload } from '../types'

describe('AutoPilot', () => {
    let graphClient: Client
    let autopilot: AutoPilot
    const device = { id: '123' }

    beforeEach(() => {
        graphClient = mockClient() as never as Client
        autopilot = new AutoPilot(graphClient)
    })

    it('should get an autopilot device', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(device)
        const result = await autopilot.getDevice('')
        expect(result).toEqual(device)
    })

    it('should get autopilot devices', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue({ value: [device] })
        const result = await autopilot.listDevices()
        expect(result).toEqual([device])
    })

    it('should import an autopilot device', async () => {
        const postSpy = jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(device)
        const result = await autopilot.importDevice(device as AutoPilotUpload)
        expect(result).toEqual(device)
        expect(postSpy).toHaveBeenCalledWith({
            '@odata.type': '#microsoft.graph.importedWindowsAutopilotDeviceIdentity',
            orderIdentifier: null,
            serialNumber: null,
            productKey: null,
            hardwareIdentifier: null,
            assignedUserPrincipalName: null,
        })
    })
})
