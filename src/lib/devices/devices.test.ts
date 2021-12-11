import { Client } from '@microsoft/microsoft-graph-client'
import { Devices } from './devices'
import { mockClient } from '../../test/mocks/@microsoft/microsoft-graph-client'

describe('Devices', () => {
    let devices: Devices
    const device = { id: '123' }

    describe('when finding devices', () => {
        it('should list all devices', async () => {
            const graphClient = mockClient() as never as Client
            devices = new Devices(graphClient)
            jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
                value: [device],
                '@odata.nextLink': 'next',
            })

            jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
                value: [device],
            })

            const result = await devices.list()
            expect(result).toEqual([device, device])
        })

        it('should get one device', async () => {
            const graphClient = mockClient() as never as Client
            devices = new Devices(graphClient)
            jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(device)
            const result = await devices.get('')
            expect(result).toEqual(device)
        })

        it('should update device', async () => {
            const graphClient = mockClient() as never as Client
            devices = new Devices(graphClient)
            jest.spyOn(graphClient.api(''), 'patch').mockResolvedValue(device)
            const result = await devices.update('id', device)
            expect(result).toEqual(device)
        })

        it('should delete device', async () => {
            const graphClient = mockClient() as never as Client
            devices = new Devices(graphClient)
            jest.spyOn(graphClient.api(''), 'delete')
            const result = await devices.delete('id')
            expect(result).toBeUndefined()
        })

        it('should get azure devices', async () => {
            const graphClient = mockClient() as never as Client
            devices = new Devices(graphClient)
            jest.spyOn(graphClient.api(''), 'get').mockResolvedValue({ value: [device] })
            const result = await devices.getAzureAdDevices()
            expect(result).toEqual([device])
        })

        it('should get one azure device', async () => {
            const graphClient = mockClient() as never as Client
            devices = new Devices(graphClient)
            jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(device)
            const result = await devices.getAzureAdDevice('')
            expect(result).toEqual(device)
        })

        it('should get autopilot devices', async () => {
            const graphClient = mockClient() as never as Client
            devices = new Devices(graphClient)
            jest.spyOn(graphClient.api(''), 'get').mockResolvedValue({ value: [device] })
            const result = await devices.listAutopilotDevices()
            expect(result).toEqual([device])
        })

        it('should get an autopilot device', async () => {
            const graphClient = mockClient() as never as Client
            devices = new Devices(graphClient)
            jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(device)
            const result = await devices.getAutopilotDevice('')
            expect(result).toEqual(device)
        })
    })

    describe('when updating devices', () => {
        it('should set device name', async () => {
            const graphClient = mockClient() as never as Client
            devices = new Devices(graphClient)
            const apiSpy = jest.spyOn(graphClient, 'api')
            const postSpy = jest.spyOn(graphClient.api(''), 'post')
            await devices.setDeviceName('id', 'name')

            expect(apiSpy).toHaveBeenCalledWith(`/deviceManagement/managedDevices/id/setDeviceName`)
            expect(postSpy).toHaveBeenCalledWith({ deviceName: 'name' })
        })

        it('should reboot device', async () => {
            const graphClient = mockClient() as never as Client
            devices = new Devices(graphClient)
            const apiSpy = jest.spyOn(graphClient, 'api')
            const postSpy = jest.spyOn(graphClient.api(''), 'post')
            await devices.rebootDevice('id')

            expect(apiSpy).toHaveBeenCalledWith(`/deviceManagement/managedDevices/id/rebootNow`)
            expect(postSpy).toHaveBeenCalledWith({})
        })

        it('should retire device', async () => {
            const graphClient = mockClient() as never as Client
            devices = new Devices(graphClient)
            const apiSpy = jest.spyOn(graphClient, 'api')
            const postSpy = jest.spyOn(graphClient.api(''), 'post')
            await devices.retireDevice('id')

            expect(apiSpy).toHaveBeenCalledWith(`/deviceManagement/managedDevices/id/retire`)
            expect(postSpy).toHaveBeenCalledWith({})
        })

        it('should shutdown device', async () => {
            const graphClient = mockClient() as never as Client
            devices = new Devices(graphClient)
            const apiSpy = jest.spyOn(graphClient, 'api')
            const postSpy = jest.spyOn(graphClient.api(''), 'post')
            await devices.shutdownDevice('id')

            expect(apiSpy).toHaveBeenCalledWith('/deviceManagement/managedDevices/id/shutDown')
            expect(postSpy).toHaveBeenCalledWith({})
        })

        it('should wipe device', async () => {
            const graphClient = mockClient() as never as Client
            devices = new Devices(graphClient)
            const apiSpy = jest.spyOn(graphClient, 'api')
            const postSpy = jest.spyOn(graphClient.api(''), 'post')
            await devices.wipeDevice('id', true, false, true, '23')

            expect(apiSpy).toHaveBeenCalledWith('/deviceManagement/managedDevices/id/wipe')
            expect(postSpy).toHaveBeenCalledWith({
                keepEnrollmentData: true,
                keepUserData: false,
                useProtectedWipe: true,
                macOsUnlockCode: '23',
            })
        })
    })
})