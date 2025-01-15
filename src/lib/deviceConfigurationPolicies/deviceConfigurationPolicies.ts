import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceManagementConfigurationPolicy } from '@microsoft/microsoft-graph-types-beta'

interface AssignmentTarget {
    '@odata.type': string
    deviceAndAppManagementAssignmentFilterType: 'none' | 'include' | 'exclude'
    groupId?: string
}

interface Assignment {
    id?: string
    source?: 'direct'
    target: AssignmentTarget
}

interface AssignmentOptions {
    includeGroups?: string[]
    excludeGroups?: string[]
    allDevices?: boolean
    allUsers?: boolean
}

export class DeviceConfigurationPolicies {
    constructor(private readonly graphClient: Client) {}

    /**
     *  List all device management configuration policies
     *
     * @returns
     */
    async list(): Promise<DeviceManagementConfigurationPolicy[]> {
        let res = await this.graphClient.api('/deviceManagement/configurationPolicies').get()
        const configurationPolicies: DeviceManagementConfigurationPolicy[] = res.value

        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextConfigurationPolicies = res.value as DeviceManagementConfigurationPolicy[]
            configurationPolicies.push(...nextConfigurationPolicies)
        }

        return configurationPolicies
    }

    /**
     *  Get a device management configuration policy
     * @param configurationPolicyId
     * @returns
     */
    async get(configurationPolicyId: string): Promise<DeviceManagementConfigurationPolicy> {
        return await this.graphClient
            .api(`/deviceManagement/configurationPolicies/${configurationPolicyId}`)
            .get()
    }

    /**
     *  Create a device management configuration policy
     * @param configurationPolicy
     * @returns
     */
    async create(
        configurationPolicy: DeviceManagementConfigurationPolicy,
    ): Promise<DeviceManagementConfigurationPolicy> {
        return this.graphClient
            .api('/deviceManagement/configurationPolicies')
            .post(configurationPolicy)
    }

    /**
     *  Update a device management configuration policy
     * @param configurationPolicyId
     * @param configurationPolicy
     */
    async update(
        configurationPolicyId: string,
        configurationPolicy: DeviceManagementConfigurationPolicy,
    ): Promise<void> {
        await this.graphClient
            .api(`/deviceManagement/configurationPolicies/${configurationPolicyId}`)
            .patch(configurationPolicy)
    }

    /**
     *  Delete a device management configuration policy
     * @param configurationPolicyId
     */
    async delete(configurationPolicyId: string): Promise<void> {
        await this.graphClient
            .api(`/deviceManagement/configurationPolicies/${configurationPolicyId}`)
            .delete()
    }

    /**
     * Set assignments for a configuration policy
     *
     * THIS WILL OVERWRITE ANY EXISTING ASSIGNMENTS!
     * @param id - The ID of the configuration policy
     * @param options - Assignment options including groups to include/exclude and whether to assign to all devices/users
     * @returns Promise<void>
     */
    async setAssignments(id: string, options: AssignmentOptions): Promise<void> {
        const assignments: Assignment[] = []

        // Add all devices assignment if specified
        if (options.allDevices) {
            assignments.push({
                id: '',
                source: 'direct',
                target: {
                    '@odata.type': '#microsoft.graph.allDevicesAssignmentTarget',
                    deviceAndAppManagementAssignmentFilterType: 'none',
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
                id: '',
                source: 'direct',
                target: {
                    '@odata.type': '#microsoft.graph.allLicensedUsersAssignmentTarget',
                    deviceAndAppManagementAssignmentFilterType: 'none',
                },
            })
        }

        // Add included groups
        if (options.includeGroups?.length) {
            assignments.push(
                ...options.includeGroups.map(
                    (groupId): Assignment => ({
                        id: '',
                        source: 'direct',
                        target: {
                            '@odata.type': '#microsoft.graph.groupAssignmentTarget',
                            deviceAndAppManagementAssignmentFilterType: 'none',
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
                        id: '',
                        source: 'direct',
                        target: {
                            '@odata.type': '#microsoft.graph.exclusionGroupAssignmentTarget',
                            deviceAndAppManagementAssignmentFilterType: 'none',
                            groupId,
                        },
                    }),
                ),
            )
        }

        await this.graphClient
            .api(`/deviceManagement/configurationPolicies('${id}')/assign`)
            .post({ assignments })
    }
}
