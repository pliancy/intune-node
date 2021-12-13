import { Client } from '@microsoft/microsoft-graph-client'
import { Device, ManagedDevice } from '@microsoft/microsoft-graph-types-beta'

export class Devices {
    constructor(private readonly graphClient: Client) {}

    async list() {
        let res = await this.graphClient.api('/deviceManagement/managedDevices').get()
        const devices: ManagedDevice[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextDevices = res.value as ManagedDevice[]
            devices.push(...nextDevices)
        }
        return devices
    }

    async get(deviceId: string): Promise<ManagedDevice> {
        return this.graphClient.api(`/deviceManagement/managedDevices/${deviceId}`).get()
    }

    async update(deviceId: string, update: Partial<ManagedDevice>): Promise<ManagedDevice> {
        return this.graphClient.api(`/deviceManagement/managedDevices/${deviceId}`).patch(update)
    }

    async delete(deviceId: string): Promise<void> {
        await this.graphClient.api(`/deviceManagement/managedDevices/${deviceId}`).delete()
    }

    async listAzureAdDevices() {
        let res = await this.graphClient.api('/devices').get()
        const devices: Device[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextDevices = res.value as Device[]
            devices.push(...nextDevices)
        }
        return devices
    }

    async getAzureAdDevice(deviceId: string): Promise<Device> {
        return this.graphClient.api(`/devices/${deviceId}`).get()
    }

    async setDeviceName(deviceId: string, newDeviceName: string): Promise<void> {
        return this.graphClient
            .api(`/deviceManagement/managedDevices/${deviceId}/setDeviceName`)
            .post({ deviceName: newDeviceName })
    }

    async rebootDevice(deviceId: string): Promise<void> {
        return this.graphClient
            .api(`/deviceManagement/managedDevices/${deviceId}/rebootNow`)
            .post({})
    }

    async retireDevice(deviceId: string): Promise<void> {
        return this.graphClient.api(`/deviceManagement/managedDevices/${deviceId}/retire`).post({})
    }

    async shutdownDevice(deviceId: string): Promise<void> {
        return this.graphClient
            .api(`/deviceManagement/managedDevices/${deviceId}/shutDown`)
            .post({})
    }

    async wipeDevice(
        deviceId: string,
        keepEnrollmentData: boolean,
        keepUserData: boolean,
        useProtectedWipe: boolean,
        macOsUnlockCode?: string,
    ): Promise<any> {
        const body: any = {
            keepEnrollmentData,
            keepUserData,
            useProtectedWipe,
        }
        if (macOsUnlockCode) body.macOsUnlockCode = macOsUnlockCode
        return this.graphClient.api(`/deviceManagement/managedDevices/${deviceId}/wipe`).post(body)
    }
}
