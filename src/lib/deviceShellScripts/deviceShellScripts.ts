import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceShellScript, DeviceShellScriptAssignment } from '../types'

export class DeviceShellScripts {
    constructor(private readonly graphClient: Client) {}

    async list() {
        let res = await this.graphClient.api('/deviceManagement/deviceShellScripts').get()
        const scripts: DeviceShellScript[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextScripts = res.value as DeviceShellScript[]
            scripts.push(...nextScripts)
        }
        return scripts
    }

    /**
     * Get a device shell script by ID
     * @param id The ID of the script to retrieve
     */
    async get(id: string): Promise<DeviceShellScript>

    /**
     * Get a device shell script by ID with optional assignments
     * @param id The ID of the script to retrieve
     * @param includeAssignments Whether to include assignment information
     */
    async get(
        id: string,
        includeAssignments: boolean,
    ): Promise<DeviceShellScript & { assignments?: Array<{ id: string; target: any }> }>

    /**
     * Implementation of get method
     */
    async get(
        id: string,
        includeAssignments: boolean = false,
    ): Promise<DeviceShellScript & { assignments?: Array<{ id: string; target: any }> }> {
        let request = this.graphClient.api(`/deviceManagement/deviceShellScripts/${id}`)

        if (includeAssignments) {
            request = request.expand('assignments')
        }

        return await request.get()
    }

    async create(script: DeviceShellScript): Promise<DeviceShellScript> {
        return await this.graphClient.api('/deviceManagement/deviceShellScripts').post(script)
    }

    async update(id: string, script: DeviceShellScript): Promise<DeviceShellScript> {
        return await this.graphClient
            .api(`/deviceManagement/deviceShellScripts/${id}`)
            .patch(script)
    }

    async delete(id: string): Promise<void> {
        await this.graphClient.api(`/deviceManagement/deviceShellScripts/${id}`).delete()
    }

    /**
     * THIS WILL OVERWRITE ANY EXISTING ASSIGNMENTS!
     * Assigns a device shell script to groups or devices.
     * This is a full replacement operation - any existing assignments not included will be removed.
     *
     * @param id The ID of the device shell script
     * @param assignments The complete set of assignments to apply
     */
    async setAssignment(id: string, assignments: DeviceShellScriptAssignment[]): Promise<void> {
        await this.graphClient
            .api(`/deviceManagement/DeviceShellScripts/${id}/assign`)
            .post({ deviceManagementScriptAssignments: assignments.map((a) => ({ target: a })) })
    }

    /**
     * THIS WILL OVERWRITE ANY EXISTING ASSIGNMENTS!
     * Assigns a device shell script to all devices.
     *
     * @param id The ID of the device shell script
     */
    async setAllDevicesAssignment(id: string): Promise<void> {
        await this.setAssignment(id, [
            {
                '@odata.type': '#microsoft.graph.allDevicesAssignmentTarget',
            },
        ])
    }

    /**
     * THIS WILL OVERWRITE ANY EXISTING ASSIGNMENTS!
     * Assigns a device shell script to all users.
     *
     * @param id The ID of the device shell script
     */
    async setAllUsersAssignment(id: string): Promise<void> {
        await this.setAssignment(id, [
            {
                '@odata.type': '#microsoft.graph.allLicensedUsersAssignmentTarget',
            },
        ])
    }
}
