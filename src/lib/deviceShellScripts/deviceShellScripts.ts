import { Client } from '@microsoft/microsoft-graph-client'
import {
    DeviceShellScript,
    DeviceShellScriptAssignment,
    DeviceShellScriptGroupAssignment,
} from '../types'

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
     * Assigns a device shell script to groups or devices.
     * This is a full replacement operation - any existing assignments not included will be removed.
     *
     * @param id The ID of the device shell script
     * @param assignments The complete set of assignments to apply
     */
    async assign(
        id: string,
        assignments: {
            deviceManagementScriptGroupAssignments?: DeviceShellScriptGroupAssignment[]
            deviceManagementScriptAssignments?: DeviceShellScriptAssignment[]
        },
    ): Promise<void> {
        await this.graphClient
            .api(`/deviceManagement/deviceShellScripts/${id}/assign`)
            .post(assignments)
    }

    /**
     * Assigns a device shell script to a group.
     * This is a convenience method that gets current assignments and adds the new group.
     *
     * @param id The ID of the device shell script
     * @param groupId The ID of the group to assign the script to
     */
    async assignToGroup(id: string, groupId: string): Promise<void> {
        // Create a new group assignment
        const newGroupAssignment: DeviceShellScriptGroupAssignment = {
            '@odata.type': '#microsoft.graph.deviceShellScriptGroupAssignment',
            targetGroupId: groupId,
        }

        // Construct the assignment payload
        const assignmentPayload = {
            deviceManagementScriptGroupAssignments: [newGroupAssignment],
        }

        // Assign the script
        await this.assign(id, assignmentPayload)
    }

    /**
     * Assigns a device shell script to all devices.
     *
     * @param id The ID of the device shell script
     */
    async assignToAllDevices(id: string): Promise<void> {
        // Create assignment for all devices
        const assignment: DeviceShellScriptAssignment = {
            '@odata.type': '#microsoft.graph.deviceShellScriptAssignment',
            target: {
                '@odata.type': '#microsoft.graph.allDevicesAssignmentTarget',
            },
        } as DeviceShellScriptAssignment

        // Construct the assignment payload
        const assignmentPayload = {
            deviceManagementScriptAssignments: [assignment],
        }

        // Assign the script
        await this.assign(id, assignmentPayload)
    }

    /**
     * Removes an assignment from a device shell script.
     * This is a convenience method that gets current assignments and removes the specified one.
     *
     * @param id The ID of the device shell script
     * @param assignmentId The ID of the assignment to remove
     */
    async removeAssignment(id: string, assignmentId: string): Promise<void> {
        // Get current script with assignments
        const script = await this.get(id, true)

        if (!script.assignments) {
            return // No assignments to remove
        }

        // Filter out the assignment to remove
        const updatedAssignments = script.assignments.filter(
            (assignment) => assignment.id !== assignmentId,
        )

        // If there are no changes, return early
        if (updatedAssignments.length === script.assignments.length) {
            return
        }

        // Construct the assignment payload based on assignment types
        const groupAssignments: DeviceShellScriptGroupAssignment[] = []
        const otherAssignments: DeviceShellScriptAssignment[] = []

        // Process each assignment
        updatedAssignments.forEach((assignment) => {
            // Use type assertion to access @odata.type
            const targetType = assignment.target && (assignment.target as any)['@odata.type']

            if (targetType === '#microsoft.graph.groupAssignmentTarget') {
                // Handle group assignments
                groupAssignments.push({
                    '@odata.type': '#microsoft.graph.deviceShellScriptGroupAssignment',
                    id: assignment.id,
                    targetGroupId: (assignment.target as any).groupId,
                })
            } else {
                // Handle other assignment types
                otherAssignments.push({
                    '@odata.type': '#microsoft.graph.deviceShellScriptAssignment',
                    ...assignment,
                })
            }
        })

        // Create the final payload
        const assignmentPayload: {
            deviceManagementScriptGroupAssignments?: DeviceShellScriptGroupAssignment[]
            deviceManagementScriptAssignments?: DeviceShellScriptAssignment[]
        } = {}

        if (groupAssignments.length > 0) {
            assignmentPayload.deviceManagementScriptGroupAssignments = groupAssignments
        }

        if (otherAssignments.length > 0) {
            assignmentPayload.deviceManagementScriptAssignments = otherAssignments
        }

        // Update assignments
        await this.assign(id, assignmentPayload)
    }
}
