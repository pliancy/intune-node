import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceManagementConfigurationPolicy } from '@microsoft/microsoft-graph-types-beta'

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
     *  Create a group assignment
     * @param id
     * @param groupId
     * @returns
     */
    async assignToGroup(id: string, groupId: string): Promise<any> {
        return this.graphClient
            .api(`/deviceManagement/configurationPolicies/${id}/assignments`)
            .post({
                '@odata.type': '#microsoft.graph.deviceManagementConfigurationPolicyAssignment',
                target: {
                    '@odata.type': '#microsoft.graph.groupAssignmentTarget',
                    groupId: groupId,
                },
            })
    }

    /**
     *  List all assignments
     * @param id
     * @returns
     */
    async listAssignments(id: string): Promise<any[]> {
        let res = await this.graphClient
            .api(`/deviceManagement/configurationPolicies/${id}/assignments`)
            .get()

        const assignments: any[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextAssignments = res.value
            assignments.push(...nextAssignments)
        }

        return assignments
    }

    /**
     *  Delete an assignment
     * @param id
     * @param assignmentId
     */
    async deleteAssignment(id: string, assignmentId: string): Promise<void> {
        await this.graphClient
            .api(`/deviceManagement/configurationPolicies/${id}/assignments/${assignmentId}`)
            .delete()
    }

    /**
     *  Get an assignment
     * @param id
     * @param assignmentId
     * @returns
     */
    async getAssignment(id: string, assignmentId: string): Promise<any> {
        return this.graphClient
            .api(`/deviceManagement/configurationPolicies/${id}/assignments/${assignmentId}`)
            .get()
    }
}
