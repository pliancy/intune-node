import { Client } from '@microsoft/microsoft-graph-client'
import { Devices } from './devices'
import { mockClient } from '../../../__fixtures__/@microsoft/microsoft-graph-client'

describe('Devices', () => {
    let graphClient: Client
    let devices: Devices
    const device = { id: '123' }

    beforeEach(() => {
        graphClient = mockClient() as never as Client
        devices = new Devices(graphClient)
    })

    describe('when finding devices', () => {
        it('should list all devices', async () => {
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
            jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(device)
            const result = await devices.get('')
            expect(result).toEqual(device)
        })

        it('should update device', async () => {
            jest.spyOn(graphClient.api(''), 'patch').mockResolvedValue(device)
            const result = await devices.update('id', device)
            expect(result).toEqual(device)
        })

        it('should delete device', async () => {
            jest.spyOn(graphClient.api(''), 'delete')
            const result = await devices.delete('id')
            expect(result).toBeUndefined()
        })

        it('should get azure devices', async () => {
            jest.spyOn(graphClient.api(''), 'get').mockResolvedValue({ value: [device] })
            const result = await devices.listAzureAdDevices()
            expect(result).toEqual([device])
        })

        it('should get one azure device', async () => {
            jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(device)
            const result = await devices.getAzureAdDevice('')
            expect(result).toEqual(device)
        })

        it('should list  detected apps', async () => {
            jest.spyOn(graphClient.api(''), 'get').mockResolvedValue({ value: [device] })
            const result = await devices.listDetectedApps('id')
            expect(result).toEqual([device])
        })
    })

    describe('when updating devices', () => {
        it('should set device name', async () => {
            const apiSpy = jest.spyOn(graphClient, 'api')
            const postSpy = jest.spyOn(graphClient.api(''), 'post')
            await devices.setDeviceName('id', 'name')

            expect(apiSpy).toHaveBeenCalledWith(`/deviceManagement/managedDevices/id/setDeviceName`)
            expect(postSpy).toHaveBeenCalledWith({ deviceName: 'name' })
        })

        it('should reboot device', async () => {
            const apiSpy = jest.spyOn(graphClient, 'api')
            const postSpy = jest.spyOn(graphClient.api(''), 'post')
            await devices.rebootDevice('id')

            expect(apiSpy).toHaveBeenCalledWith(`/deviceManagement/managedDevices/id/rebootNow`)
            expect(postSpy).toHaveBeenCalledWith({})
        })

        it('should sync  device', async () => {
            const apiSpy = jest.spyOn(graphClient, 'api')
            const postSpy = jest.spyOn(graphClient.api(''), 'post')
            await devices.syncDevice('id')

            expect(apiSpy).toHaveBeenCalledWith(`/deviceManagement/managedDevices/id/syncDevice`)
            expect(postSpy).toHaveBeenCalledWith({})
        })

        it('should retire device', async () => {
            const apiSpy = jest.spyOn(graphClient, 'api')
            const postSpy = jest.spyOn(graphClient.api(''), 'post')
            await devices.retireDevice('id')

            expect(apiSpy).toHaveBeenCalledWith(`/deviceManagement/managedDevices/id/retire`)
            expect(postSpy).toHaveBeenCalledWith({})
        })

        it('should shutdown device', async () => {
            const apiSpy = jest.spyOn(graphClient, 'api')
            const postSpy = jest.spyOn(graphClient.api(''), 'post')
            await devices.shutdownDevice('id')

            expect(apiSpy).toHaveBeenCalledWith('/deviceManagement/managedDevices/id/shutDown')
            expect(postSpy).toHaveBeenCalledWith({})
        })

        it('should wipe device', async () => {
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

        it('should assign user to device', async () => {
            const apiSpy = jest.spyOn(graphClient, 'api')
            const postSpy = jest.spyOn(graphClient.api(''), 'post')
            await devices.assignUserToDevice('id', 'userId')

            expect(apiSpy).toHaveBeenCalledWith("/deviceManagement/managedDevices('id')/users/$ref")
            expect(postSpy).toHaveBeenCalledWith({
                '@odata.id': `https://graph.microsoft.com/beta/users/userId`,
            })
        })

        it('should unassign user from device', async () => {
            const apiSpy = jest.spyOn(graphClient, 'api')
            const deleteSpy = jest.spyOn(graphClient.api(''), 'delete')
            await devices.unassignUserFromDevice('id')

            expect(apiSpy).toHaveBeenCalledWith("/deviceManagement/managedDevices('id')/users/$ref")
            expect(deleteSpy).toHaveBeenCalled()
        })
    })
})
