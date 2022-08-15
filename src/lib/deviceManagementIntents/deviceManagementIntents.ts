import { Client } from '@microsoft/microsoft-graph-client'
import { Group } from '@microsoft/microsoft-graph-types-beta'
import { DeviceManagementIntent } from '../types'

export class DeviceManagementIntents {
    constructor(private readonly graphClient: Client) {}

    async list() {
        let res = await this.graphClient.api('/deviceManagement/intents').get()
        const intents: Group[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextIntents = res.value as DeviceManagementIntent[]
            intents.push(...nextIntents)
        }
        return intents
    }

    async get(intentId: string): Promise<Group> {
        return this.graphClient.api(`/deviceManagement/intents/${intentId}`).get()
    }

    async update(
        intentId: string,
        intent: DeviceManagementIntent,
    ): Promise<DeviceManagementIntent> {
        return this.graphClient.api(`/deviceManagement/intents/${intentId}`).patch(intent)
    }

    async delete(intentId: string): Promise<void> {
        return this.graphClient.api(`/deviceManagement/intents/${intentId}`).delete()
    }

    async create(intent: DeviceManagementIntent): Promise<DeviceManagementIntent> {
        return this.graphClient.api('/deviceManagement/intents').post(intent)
    }
}
