import { Client } from '@microsoft/microsoft-graph-client'

export class CustomRequest {
    constructor(private readonly graphClient: Client) {}

    async get(path: string) {
        return this.graphClient.api(path).get()
    }

    async post(path: string, body: any) {
        return this.graphClient.api(path).post(body)
    }

    async patch(path: string, body: any) {
        return this.graphClient.api(path).patch(body)
    }

    async delete(path: string) {
        return this.graphClient.api(path).delete()
    }

    async put(path: string, body: any) {
        return this.graphClient.api(path).put(body)
    }
}
