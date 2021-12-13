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

    /**
     * Creates a new Group Policy Configuration with definition values
     *
     * @param groupPolicyConfiguration - Group Policy Configuration to create
     * @param groupPolicyDefinitionValues - Group Policy Definition Values to create
     * @param mobileAppContentFile - the content info for the .intunewin file
     * @returns the created Group Policy Configuration and the created Group Policy Definition Values
     */
    async createWithDefinitionValues(
        groupPolicyConfiguration: GroupPolicyConfiguration,
        groupPolicyDefinitionValues: GroupPolicyDefinitionValue[],
    ): Promise<GroupPolicyConfiguration> {
        const groupPolicyConfigurationRes = await this.create(groupPolicyConfiguration)
        const groupPolicyConfigurationId = groupPolicyConfigurationRes.id as string
        const definitionValues: GroupPolicyDefinitionValue[] = []
        for (const groupPolicyDefinitionValue of groupPolicyDefinitionValues) {
            const res = await this.createPolicyDefinitionValue(
                groupPolicyConfigurationId,
                groupPolicyDefinitionValue,
            )
            definitionValues.push(res)
        }
        return {
            ...groupPolicyConfigurationRes,
            definitionValues,
        }
    }

    async getWithDefinitionValues(
        groupPolicyConfigurationId: string,
    ): Promise<GroupPolicyConfiguration> {
        const groupPolicyConfigurationRes = await this.get(groupPolicyConfigurationId)
        const definitionValues = await this.listPolicyDefinitionValues(groupPolicyConfigurationId)
        return {
            ...groupPolicyConfigurationRes,
            definitionValues,
        }
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
