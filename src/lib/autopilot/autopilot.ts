import { AutoPilotUpload } from '../types'
import { Client } from '@microsoft/microsoft-graph-client'
import { WindowsAutopilotDeviceIdentity } from '@microsoft/microsoft-graph-types-beta'

export class AutoPilot {
    constructor(private readonly graphClient: Client) {}

    // Autopilot
    async listDevices() {
        let res = await this.graphClient
            .api('/deviceManagement/windowsAutopilotDeviceIdentities')
            .get()
        const devices: WindowsAutopilotDeviceIdentity[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextDevices = res.value as WindowsAutopilotDeviceIdentity[]
            devices.push(...nextDevices)
        }
        return devices
    }

    async getDevice(deviceId: string): Promise<WindowsAutopilotDeviceIdentity> {
        return this.graphClient
            .api(`/deviceManagement/windowsAutopilotDeviceIdentities/${deviceId}`)
            .get()
    }

    async importDevice(autoPilotUpload: AutoPilotUpload) {
        const body = {
            '@odata.type': '#microsoft.graph.importedWindowsAutopilotDeviceIdentity',
            orderIdentifier: autoPilotUpload.groupTag ?? null,
            serialNumber: autoPilotUpload.serialNumber ?? null,
            productKey: autoPilotUpload.productKey ?? null,
            hardwareIdentifier: autoPilotUpload.hardwareIdentifier ?? null,
            assignedUserPrincipalName: autoPilotUpload.assignedUser ?? null,
        }
        const res = await this.graphClient
            .api('/deviceManagement/importedWindowsAutopilotDeviceIdentities/import')
            .post(body)

        const device = res as WindowsAutopilotDeviceIdentity
        return device
    }
}
