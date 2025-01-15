import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceConfigurationGroupAssignment } from '@microsoft/microsoft-graph-types-beta'
import { DeviceConfiguration } from '../types'

interface AssignmentTarget {
    '@odata.type': string
    groupId?: string
}

interface Assignment {
    target: AssignmentTarget
}

interface AssignmentOptions {
    includeGroups?: string[]
    excludeGroups?: string[]
    allDevices?: boolean
    allUsers?: boolean
}
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

    /**
     * Set assignments for a device configuration
     *
     * THIS WILL OVERWRITE ANY EXISTING ASSIGNMENTS!
     * @param id - The ID of the device configuration
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

        await this.graphClient
            .api(`/deviceManagement/deviceConfigurations/${id}/assign`)
            .post({ assignments })
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
            const nextGroupAssignments: DeviceConfigurationGroupAssignment[] = res.value
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
