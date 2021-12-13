import { Client } from '@microsoft/microsoft-graph-client'
import { MobileAppIntentAndState, User } from '@microsoft/microsoft-graph-types-beta'
export class Users {
    constructor(private readonly graphClient: Client) {}

    async list() {
        let res = await this.graphClient.api('/users').get()
        const users: User[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextUsers = res.value as User[]
            users.push(...nextUsers)
        }
        return users
    }

    async get(userId: string): Promise<User> {
        return this.graphClient.api(`/users/${userId}`).get()
    }

    async update(userId: string, user: User): Promise<User> {
        return this.graphClient.api(`/users/${userId}`).patch(user)
    }

    async delete(userId: string): Promise<void> {
        return this.graphClient.api(`/users/${userId}`).delete()
    }

    async create(user: User): Promise<User> {
        return this.graphClient.api('/users').post(user)
    }

    async listAppIntentAndStates(userId: string): Promise<any> {
        let res = await this.graphClient.api(`/users/${userId}/mobileAppIntentAndStates`).get()
        const state: MobileAppIntentAndState[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextState = res.value as MobileAppIntentAndState[]
            state.push(...nextState)
        }
        return state
    }
}
