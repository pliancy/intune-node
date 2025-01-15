import { Client } from '@microsoft/microsoft-graph-client'
import { Group } from '@microsoft/microsoft-graph-types-beta'
import { DeviceManagementIntent } from '../types'

interface AssignmentTarget {
    '@odata.type': string
    groupId?: string
}

interface Assignment {
    id?: string
    target: AssignmentTarget
}

interface AssignmentOptions {
    includeGroups?: string[]
    excludeGroups?: string[]
    allDevices?: boolean
    allUsers?: boolean
}

export class DeviceManagementIntents {
    constructor(private readonly graphClient: Client) {}

    async list() {
        let res = await this.graphClient.api('/deviceManagement/intents').get()
        const intents: DeviceManagementIntent[] = res.value
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

    /**
     * Set assignments for an intent
     *
     * THIS WILL OVERWRITE ANY EXISTING ASSIGNMENTS!
     * @param id - The ID of the intent
     * @param options - Assignment options including groups to include/exclude and whether to assign to all devices/users
     * @returns Promise<void>
     */
    async setAssignments(id: string, options: AssignmentOptions): Promise<void> {
        const assignments: Assignment[] = []

        // Add all devices assignment if specified
        if (options.allDevices) {
            assignments.push({
                target: {
                    '@odata.type': '#microsoft.graph.allDevicesAssignmentTarget',
                },
            })

            // When all devices is selected, we can only have exclusion groups
            if (options.includeGroups?.length) {
                throw new Error('Cannot include specific groups when allDevices is true')
            }
        }

        // Add all licensed users assignment if specified
        if (options.allUsers) {
            assignments.push({
                target: {
                    '@odata.type': '#microsoft.graph.allLicensedUsersAssignmentTarget',
                },
            })
        }

        // Add included groups
        if (options.includeGroups?.length) {
            assignments.push(
                ...options.includeGroups.map(
                    (groupId): Assignment => ({
                        target: {
                            '@odata.type': '#microsoft.graph.groupAssignmentTarget',
                            groupId,
                        },
                    }),
                ),
            )
        }

        // Add excluded groups
        if (options.excludeGroups?.length) {
            assignments.push(
                ...options.excludeGroups.map(
                    (groupId): Assignment => ({
                        target: {
                            '@odata.type': '#microsoft.graph.exclusionGroupAssignmentTarget',
                            groupId,
                        },
                    }),
                ),
            )
        }

        await this.graphClient.api(`/deviceManagement/intents/${id}/assign`).post({ assignments })
    }
}
