import { Client } from '@microsoft/microsoft-graph-client'
import { Group } from '@microsoft/microsoft-graph-types-beta'
import { GroupMember } from 'lib/types'
export class Groups {
    constructor(private readonly graphClient: Client) {}

    async list() {
        let res = await this.graphClient.api('/groups').get()
        const groups: Group[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextGroups = res.value as Group[]
            groups.push(...nextGroups)
        }
        return groups
    }

    async get(groupId: string): Promise<Group> {
        return this.graphClient.api(`/groups/${groupId}`).get()
    }

    async update(groupId: string, group: Group): Promise<Group> {
        return this.graphClient.api(`/groups/${groupId}`).patch(group)
    }

    async delete(groupId: string): Promise<void> {
        return this.graphClient.api(`/groups/${groupId}`).delete()
    }

    async create(group: Group): Promise<Group> {
        return this.graphClient.api('/groups').post(group)
    }

    async listMembers(groupId: string): Promise<GroupMember[]> {
        let res = await this.graphClient.api(`/groups/${groupId}/members`).get()
        const members: GroupMember[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextMembers = res.value as GroupMember[]
            members.push(...nextMembers)
        }
        return members
    }

    /**
     * Add members to a group
     * https://learn.microsoft.com/en-us/graph/api/group-post-members?view=graph-rest-beta&tabs=http
     * @param groupId The ID of the group
     * @param directoryIds The IDs of the directory objects to add
     * @returns
     */
    async addMembers(groupId: string, directoryIds: string[]): Promise<void> {
        return this.graphClient.api(`/groups/${groupId}`).patch({
            'members@odata.bind': directoryIds.map(
                (id) => `https://graph.microsoft.com/beta/directoryObjects/${id}`,
            ),
        })
    }

    /**
     * Remove member from a group
     * https://learn.microsoft.com/en-us/graph/api/group-delete-members?view=graph-rest-beta&tabs=http
     * @param groupId The ID of the group
     * @param directoryId The ID of the directory object to remove
     * @returns
     */
    async removeMember(groupId: string, directoryId: string): Promise<void> {
        return this.graphClient.api(`/groups/${groupId}/members/${directoryId}/$ref`).delete()
    }
}
