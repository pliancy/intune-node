import { Client } from '@microsoft/microsoft-graph-client'
import { CreateTemplateInstance, DeviceManagementIntent, DeviceManagementTemplate } from '../types'

export class DeviceManagementTemplates {
    constructor(private readonly graphClient: Client) {}

    async list() {
        let res = await this.graphClient.api('/deviceManagement/templates').get()
        const scripts: DeviceManagementTemplate[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextScripts = res.value as DeviceManagementTemplate[]
            scripts.push(...nextScripts)
        }
        return scripts
    }

    async get(id: string): Promise<DeviceManagementTemplate> {
        return await this.graphClient.api(`/deviceManagement/templates/${id}`).get()
    }

    async create(script: DeviceManagementTemplate): Promise<DeviceManagementTemplate> {
        return await this.graphClient.api('/deviceManagement/templates').post(script)
    }

    async update(id: string, script: DeviceManagementTemplate): Promise<DeviceManagementTemplate> {
        return await this.graphClient.api(`/deviceManagement/templates/${id}`).patch(script)
    }

    async delete(id: string): Promise<void> {
        await this.graphClient.api(`/deviceManagement/templates/${id}`).delete()
    }

    async createInstance(
        templateId: string,
        instance: CreateTemplateInstance,
    ): Promise<DeviceManagementIntent> {
        return this.graphClient
            .api(`/deviceManagement/templates/${templateId}/createInstance`)
            .post(instance)
    }
}
