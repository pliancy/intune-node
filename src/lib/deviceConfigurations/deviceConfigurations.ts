import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceConfigurationGroupAssignment } from '@microsoft/microsoft-graph-types-beta'
import { DeviceConfiguration } from '../types'
export class DeviceConfigurations {
    constructor(private readonly graphClient: Client) {}

    async list() {
        let res = await this.graphClient.api('/deviceManagement/deviceConfigurations').get()
        const deviceConfigurations: DeviceConfiguration[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextdeviceConfigurations = res.value as DeviceConfiguration[]
            deviceConfigurations.push(...nextdeviceConfigurations)
        }
        return deviceConfigurations
    }

    async get(deviceConfigurationId: string): Promise<DeviceConfiguration> {
        return await this.graphClient
            .api(`/deviceManagement/deviceConfigurations/${deviceConfigurationId}`)
            .get()
    }

    async create(deviceConfigurations: DeviceConfiguration): Promise<DeviceConfiguration> {
        return this.graphClient
            .api('/deviceManagement/deviceConfigurations')
            .post(deviceConfigurations)
    }

    async update(
        deviceConfigurationId: string,
        deviceConfiguration: DeviceConfiguration,
    ): Promise<void> {
        await this.graphClient
            .api(`/deviceManagement/deviceConfigurations/${deviceConfigurationId}`)
            .patch(deviceConfiguration)
    }

    async delete(deviceConfigurationId: string): Promise<void> {
        await this.graphClient
            .api(`/deviceManagement/deviceConfigurations/${deviceConfigurationId}`)
            .delete()
    }

    async createGroupAssignment(
        id: string,
        groupId: string,
    ): Promise<DeviceConfigurationGroupAssignment> {
        return this.graphClient
            .api(`/deviceManagement/deviceConfigurations/${id}/groupAssignments`)
            .post({
                '@odata.type': '#microsoft.graph.deviceConfigurationGroupAssignment',
                targetGroupId: groupId,
            })
    }

    async listGroupAssignments(id: string): Promise<DeviceConfigurationGroupAssignment[]> {
        let res = await this.graphClient
            .api(`/deviceManagement/deviceConfigurations/${id}/groupAssignments`)
            .get()
        const groupAssignments: DeviceConfigurationGroupAssignment[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextGroupAssignments = res.value as DeviceConfigurationGroupAssignment[]
            groupAssignments.push(...nextGroupAssignments)
        }
        return groupAssignments
    }

    async deleteGroupAssignment(id: string, groupAssignmentId: string): Promise<void> {
        await this.graphClient
            .api(
                `/deviceManagement/deviceConfigurations/${id}/groupAssignments/${groupAssignmentId}`,
            )
            .delete()
    }

    async getGroupAssignment(
        id: string,
        groupAssignmentId: string,
    ): Promise<DeviceConfigurationGroupAssignment> {
        return this.graphClient
            .api(
                `/deviceManagement/deviceConfigurations/${id}/groupAssignments/${groupAssignmentId}`,
            )
            .get()
    }
}
