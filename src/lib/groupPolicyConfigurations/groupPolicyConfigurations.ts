import { Client } from '@microsoft/microsoft-graph-client'
import {
    GroupPolicyConfiguration,
    GroupPolicyConfigurationAssignment,
    GroupPolicyDefinitionValue,
} from '@microsoft/microsoft-graph-types-beta'
export class GroupPolicyConfigurations {
    constructor(private readonly graphClient: Client) {}

    async list(): Promise<GroupPolicyConfiguration[]> {
        let res = await this.graphClient.api('/deviceManagement/groupPolicyConfigurations').get()
        const data: any[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextData = res.value as any[]
            data.push(...nextData)
        }
        return data
    }

    async get(groupPolicyConfigurationId: string): Promise<GroupPolicyConfiguration> {
        return await this.graphClient
            .api(`/deviceManagement/groupPolicyConfigurations/${groupPolicyConfigurationId}`)
            .get()
    }

    async create(
        groupPolicyConfiguration: GroupPolicyConfiguration,
    ): Promise<GroupPolicyConfiguration> {
        return this.graphClient
            .api('/deviceManagement/groupPolicyConfigurations')
            .post(groupPolicyConfiguration)
    }

    async update(
        groupPolicyConfigurationId: string,
        groupPolicyConfiguration: GroupPolicyConfiguration,
    ): Promise<void> {
        await this.graphClient
            .api(`/deviceManagement/groupPolicyConfigurations/${groupPolicyConfigurationId}`)
            .patch(groupPolicyConfiguration)
    }

    async delete(groupPolicyConfigurationId: string): Promise<void> {
        await this.graphClient
            .api(`/deviceManagement/groupPolicyConfigurations/${groupPolicyConfigurationId}`)
            .delete()
    }

    async listPolicyDefinitionValues(
        groupPolicyConfigurationId: string,
    ): Promise<GroupPolicyDefinitionValue[]> {
        let res = await this.graphClient
            .api(
                `/deviceManagement/groupPolicyConfigurations/${groupPolicyConfigurationId}/definitionValues`,
            )
            .get()
        const data: any[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextData = res.value as any[]
            data.push(...nextData)
        }
        return data
    }

    async getPolicyDefinitionValue(
        groupPolicyConfigurationId: string,
        groupPolicyDefinitionValueId: string,
    ): Promise<GroupPolicyDefinitionValue> {
        return await this.graphClient
            .api(
                `/deviceManagement/groupPolicyConfigurations/${groupPolicyConfigurationId}/definitionValues/${groupPolicyDefinitionValueId}`,
            )
            .get()
    }

    async createPolicyDefinitionValue(
        groupPolicyConfigurationId: string,
        groupPolicyDefinitionValue: GroupPolicyDefinitionValue,
    ): Promise<GroupPolicyDefinitionValue> {
        return this.graphClient
            .api(
                `/deviceManagement/groupPolicyConfigurations/${groupPolicyConfigurationId}/definitionValues`,
            )
            .post(groupPolicyDefinitionValue)
    }

    async updatePolicyDefinitionValue(
        groupPolicyConfigurationId: string,
        groupPolicyDefinitionValueId: string,
        groupPolicyDefinitionValue: GroupPolicyDefinitionValue,
    ): Promise<GroupPolicyDefinitionValue> {
        return this.graphClient
            .api(
                `/deviceManagement/groupPolicyConfigurations/${groupPolicyConfigurationId}/definitionValues/${groupPolicyDefinitionValueId}`,
            )
            .patch(groupPolicyDefinitionValue)
    }

    async deletePolicyDefinitionValue(
        groupPolicyConfigurationId: string,
        groupPolicyDefinitionValueId: string,
    ): Promise<void> {
        await this.graphClient
            .api(
                `/deviceManagement/groupPolicyConfigurations/${groupPolicyConfigurationId}/definitionValues/${groupPolicyDefinitionValueId}`,
            )
            .delete()
    }

    async createAssignment(
        groupPolicyConfigurationId: string,
        groupPolicyConfigurationAssignment: GroupPolicyConfigurationAssignment,
    ): Promise<GroupPolicyConfigurationAssignment> {
        return this.graphClient
            .api(
                `/deviceManagement/groupPolicyConfigurations/${groupPolicyConfigurationId}/assignments`,
            )
            .post(groupPolicyConfigurationAssignment)
    }

    async listAssignments(
        groupPolicyConfigurationId: string,
    ): Promise<GroupPolicyConfigurationAssignment[]> {
        let res = await this.graphClient
            .api(
                `/deviceManagement/groupPolicyConfigurations/${groupPolicyConfigurationId}/assignments`,
            )
            .get()
        const data: any[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextData = res.value as any[]
            data.push(...nextData)
        }
        return data
    }

    async deleteAssignment(
        groupPolicyConfigurationId: string,
        groupPolicyConfigurationAssignmentId: string,
    ): Promise<void> {
        await this.graphClient
            .api(
                `/deviceManagement/groupPolicyConfigurations/${groupPolicyConfigurationId}/assignments/${groupPolicyConfigurationAssignmentId}`,
            )
            .delete()
    }

    async getAssignment(
        groupPolicyConfigurationId: string,
        groupPolicyConfigurationAssignmentId: string,
    ): Promise<GroupPolicyConfigurationAssignment> {
        return this.graphClient
            .api(
                `/deviceManagement/groupPolicyConfigurations/${groupPolicyConfigurationId}/assignments/${groupPolicyConfigurationAssignmentId}`,
            )
            .get()
    }
}
