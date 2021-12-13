import { Client } from '@microsoft/microsoft-graph-client'
import { Group } from '@microsoft/microsoft-graph-types-beta'
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
}
