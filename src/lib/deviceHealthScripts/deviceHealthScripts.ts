import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceHealthScript, DeviceHealthScriptAssignment } from '../types'
export class DeviceHealthScripts {
    constructor(private readonly graphClient: Client) {}

    async list() {
        let res = await this.graphClient.api('/deviceManagement/deviceHealthScripts').get()
        const scripts: DeviceHealthScript[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextScripts = res.value as DeviceHealthScript[]
            scripts.push(...nextScripts)
        }
        return scripts
    }

    async get(id: string): Promise<DeviceHealthScript> {
        return await this.graphClient.api(`/deviceManagement/deviceHealthScripts/${id}`).get()
    }

    async create(script: DeviceHealthScript): Promise<DeviceHealthScript> {
        return await this.graphClient.api('/deviceManagement/deviceHealthScripts').post(script)
    }

    async update(id: string, script: DeviceHealthScript): Promise<DeviceHealthScript> {
        return await this.graphClient
            .api(`/deviceManagement/deviceHealthScripts/${id}`)
            .patch(script)
    }

    async delete(id: string): Promise<void> {
        await this.graphClient.api(`/deviceManagement/deviceHealthScripts/${id}`).delete()
    }

    async createAssignment(
        id: string,
        assignment: DeviceHealthScriptAssignment,
    ): Promise<DeviceHealthScriptAssignment> {
        return this.graphClient
            .api(`/deviceManagement/deviceHealthScripts/${id}/assignments`)
            .post(assignment)
    }

    async listAssignments(id: string): Promise<DeviceHealthScriptAssignment[]> {
        let res = await this.graphClient
            .api(`/deviceManagement/deviceHealthScripts/${id}/assignments`)
            .get()
        const groupAssignments: DeviceHealthScriptAssignment[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextGroupAssignments = res.value as DeviceHealthScriptAssignment[]
            groupAssignments.push(...nextGroupAssignments)
        }
        return groupAssignments
    }

    async deleteAssignment(id: string, assignmentId: string): Promise<void> {
        await this.graphClient
            .api(`/deviceManagement/deviceHealthScripts/${id}/assignments/${assignmentId}`)
            .delete()
    }

    async getAssignment(id: string, assignmentId: string): Promise<DeviceHealthScriptAssignment> {
        return this.graphClient
            .api(`/deviceManagement/deviceHealthScripts/${id}/assignments/${assignmentId}`)
            .get()
    }
}
