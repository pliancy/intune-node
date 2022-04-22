import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceManagementScriptGroupAssignment } from '@microsoft/microsoft-graph-types-beta'
import { DeviceManagementScript } from '../types'
export class DeviceManagementScripts {
    constructor(private readonly graphClient: Client) {}

    async list() {
        let res = await this.graphClient.api('/deviceManagement/deviceManagementScripts').get()
        const scripts: DeviceManagementScript[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextScripts = res.value as DeviceManagementScript[]
            scripts.push(...nextScripts)
        }
        return scripts
    }

    async get(id: string): Promise<DeviceManagementScript> {
        return await this.graphClient.api(`/deviceManagement/deviceManagementScripts/${id}`).get()
    }

    async create(script: DeviceManagementScript): Promise<DeviceManagementScript> {
        return await this.graphClient.api('/deviceManagement/deviceManagementScripts').post(script)
    }

    async update(id: string, script: DeviceManagementScript): Promise<DeviceManagementScript> {
        return await this.graphClient
            .api(`/deviceManagement/deviceManagementScripts/${id}`)
            .patch(script)
    }

    async delete(id: string): Promise<void> {
        await this.graphClient.api(`/deviceManagement/deviceManagementScripts/${id}`).delete()
    }

    async createGroupAssignment(
        id: string,
        groupId: string,
    ): Promise<DeviceManagementScriptGroupAssignment> {
        return this.graphClient
            .api(`/deviceManagement/deviceManagementScripts/${id}/assignments`)
            .post({
                '@odata.type': '#microsoft.graph.deviceManagementScriptGroupAssignment',
                targetGroupId: groupId,
            })
    }

    async listGroupAssignments(id: string): Promise<DeviceManagementScriptGroupAssignment[]> {
        let res = await this.graphClient
            .api(`/deviceManagement/deviceManagementScripts/${id}/groupAssignments`)
            .get()
        const groupAssignments: DeviceManagementScriptGroupAssignment[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextGroupAssignments = res.value as DeviceManagementScriptGroupAssignment[]
            groupAssignments.push(...nextGroupAssignments)
        }
        return groupAssignments
    }

    async deleteGroupAssignment(id: string, groupAssignmentId: string): Promise<void> {
        await this.graphClient
            .api(
                `/deviceManagement/deviceManagementScripts/${id}/groupAssignments/${groupAssignmentId}`,
            )
            .delete()
    }

    async getGroupAssignment(
        id: string,
        groupAssignmentId: string,
    ): Promise<DeviceManagementScriptGroupAssignment> {
        return this.graphClient
            .api(
                `/deviceManagement/deviceManagementScripts/${id}/groupAssignments/${groupAssignmentId}`,
            )
            .get()
    }
}
