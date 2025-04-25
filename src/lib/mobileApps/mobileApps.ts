import { AnonymousCredential, BlobServiceClient, newPipeline } from '@azure/storage-blob'
import { Client } from '@microsoft/microsoft-graph-client'
import { FileEncryptionInfo } from '@microsoft/microsoft-graph-types-beta'
import { URL } from 'url'
import {
    MobileApp,
    MobileAppContentFile,
    ReadStream,
    Win32LobApp,
    MobileAppRelationship,
    MobileAppAssignment,
} from '../types'
import { sleep } from '../utils/sleep'

export class MobileApps {
    constructor(private readonly graphClient: Client) {}

    async list() {
        let res = await this.graphClient.api('/deviceAppManagement/mobileApps').get()
        const apps: Array<MobileApp> = res.value

        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextApps = res.value as Array<MobileApp>
            apps.push(...nextApps)
        }
        return apps
    }

    async get(id: string): Promise<MobileApp> {
        return await this.graphClient.api(`/deviceAppManagement/mobileApps/${id}`).get()
    }

    async create(mobileApp: MobileApp): Promise<MobileApp> {
        return this.graphClient.api('/deviceAppManagement/mobileApps').post(mobileApp)
    }

    async update(id: string, mobileApp: MobileApp): Promise<MobileApp> {
        return this.graphClient.api(`/deviceAppManagement/mobileApps/${id}`).patch(mobileApp)
    }

    async delete(id: string): Promise<void> {
        await this.graphClient.api(`/deviceAppManagement/mobileApps/${id}`).delete()
    }

    async createWin32LobContentVersion(appId: string): Promise<any> {
        return this.graphClient
            .api(
                `/deviceAppManagement/mobileApps/${appId}/microsoft.graph.win32LobApp/contentVersions`,
            )
            .post({})
    }

    async createWin32LobFileUpload(
        appId: string,
        contentVersionId: number,
        mobileAppContentFile: MobileAppContentFile,
    ): Promise<MobileAppContentFile> {
        return this.graphClient
            .api(
                `/deviceAppManagement/mobileApps/${appId}/microsoft.graph.win32LobApp/contentversions/${contentVersionId}/files`,
            )
            .post(mobileAppContentFile)
    }

    async getAzureStorageUri(
        appId: string,
        contentVersionId: number,
        fileId: string,
    ): Promise<string> {
        const res = await this.graphClient
            .api(
                `/deviceAppManagement/mobileApps/${appId}/microsoft.graph.win32LobApp/contentversions/${contentVersionId}/files/${fileId}`,
            )
            .get()
        return res.azureStorageUri
    }

    async uploadToAzureBlob(
        azureStorageUri: string,
        file: ReadStream,
        fileSize: number,
        statusCallback?: any,
        customBufferSize?: number,
    ) {
        let bufferSize: number = customBufferSize ?? 4 * 1024 * 1024
        if (fileSize < 4000 && !customBufferSize) {
            bufferSize = 1 * 1024
        }

        // Parse Storage URI
        const parseURL = new URL(azureStorageUri)
        const azureStorageUriArray = azureStorageUri.split('/')
        const sasUrl = `${parseURL.origin}${parseURL.search}`
        const blobContainer = `${azureStorageUriArray[3]}`
        const blobName = parseURL.pathname.replace(`/${blobContainer}/`, '')

        // Azure Upload
        const pipeline = newPipeline(new AnonymousCredential())
        const blobServiceClient = new BlobServiceClient(sasUrl, pipeline)
        const containerClient = blobServiceClient.getContainerClient(blobContainer)
        const blockBlobClient = containerClient.getBlockBlobClient(blobName)
        let uploadBlobResponse = {}
        if (statusCallback) {
            uploadBlobResponse = await blockBlobClient.uploadStream(file, bufferSize, 5, {
                onProgress: (ev: any) => statusCallback(fileSize, ev.loadedBytes),
            })
        } else {
            uploadBlobResponse = await blockBlobClient.uploadStream(file, bufferSize, 5)
        }
    }

    async commitWin32LobFileUpload(
        appId: string,
        contentVersionId: number,
        fileId: string,
        fileEncryptionInfo: FileEncryptionInfo,
    ) {
        await this.graphClient
            .api(
                `/deviceAppManagement/mobileApps/${appId}/microsoft.graph.win32LobApp/contentversions/${contentVersionId}/files/${fileId}/commit`,
            )
            .post({ fileEncryptionInfo })
    }

    async commitWin32LobApp(appId: string, contentVersionId: number) {
        await this.graphClient.api(`/deviceAppManagement/mobileApps/${appId}`).patch({
            '@odata.type': '#microsoft.graph.win32LobApp',
            committedContentVersion: `${contentVersionId}`,
        })
    }

    async getFileUploadStatus(
        appId: string,
        contentVersionId: number,
        fileId: string,
    ): Promise<string> {
        const res = await this.graphClient
            .api(
                `/deviceAppManagement/mobileApps/${appId}/microsoft.graph.win32LobApp/contentversions/${contentVersionId}/files/${fileId}`,
            )
            .get()
        return res.uploadState
    }

    /**
     * Creates a new win32 lob app and uploads the file .intunewin file for the app, this requires that you first unecrypt the .intunewin file
     *
     * @param mobileApp - The mobile app info to create the app
     * @param fileEncryptionInfo - the file encryption info for the .intunewin file
     * @param mobileAppContentFile - the content info for the .intunewin file
     * @param file - the unecrypted .intunewin file
     * @returns the created app info
     */
    async createWin32LobApp(
        mobileApp: Win32LobApp,
        fileEncryptionInfo: FileEncryptionInfo,
        mobileAppContentFile: MobileAppContentFile,
        file: ReadStream,
    ): Promise<Win32LobApp> {
        const createApp = (await this.create(mobileApp)) as Win32LobApp
        const appId = createApp.id as string

        const { id: contentVersionId } = await this.createWin32LobContentVersion(appId)

        const fileUpload = await this.createWin32LobFileUpload(
            appId,
            contentVersionId,
            mobileAppContentFile,
        )

        const fileId = fileUpload.id as string

        let azureStorageUri: string
        do {
            azureStorageUri = await this.getAzureStorageUri(appId, contentVersionId, fileId)
            //  Need to sleep for a second before trying again
            sleep(1000)
        } while (!azureStorageUri)

        await this.uploadToAzureBlob(azureStorageUri, file, mobileAppContentFile.size)

        await this.commitWin32LobFileUpload(appId, contentVersionId, fileId, fileEncryptionInfo)

        let status: string
        do {
            status = await this.getFileUploadStatus(appId, contentVersionId, fileId)
            //  Need to sleep for a second before trying again
            sleep(1000)
        } while (status !== 'commitFileSuccess')

        await this.commitWin32LobApp(appId, contentVersionId)

        return createApp
    }

    /**
     * Updates a win32 lob .intunewin file for the app
     *
     * @param appId - The App Id to update
     * @param fileEncryptionInfo - the file encryption info for the .intunewin file
     * @param mobileAppContentFile - the content info for the .intunewin file
     * @param file - the unecrypted .intunewin file
     */
    async updateWin32LobAppUpload(
        appId: string,
        fileEncryptionInfo: FileEncryptionInfo,
        mobileAppContentFile: MobileAppContentFile,
        file: ReadStream,
    ): Promise<void> {
        const { id: contentVersionId } = await this.createWin32LobContentVersion(appId)

        const fileUpload = await this.createWin32LobFileUpload(
            appId,
            contentVersionId,
            mobileAppContentFile,
        )

        const fileId = fileUpload.id as string
        let azureStorageUri: string
        do {
            azureStorageUri = await this.getAzureStorageUri(appId, contentVersionId, fileId)
            sleep(1000)
        } while (!azureStorageUri)

        await this.uploadToAzureBlob(azureStorageUri, file, mobileAppContentFile.size)

        await this.commitWin32LobFileUpload(appId, contentVersionId, fileId, fileEncryptionInfo)
        let status: string
        do {
            status = await this.getFileUploadStatus(appId, contentVersionId, fileId)
            sleep(1000)
        } while (status !== 'commitFileSuccess')

        await this.commitWin32LobApp(appId, contentVersionId)
    }

    async listAppRelationships(appId: string) {
        let res = await this.graphClient
            .api(`/deviceAppManagement/mobileApps/${appId}/relationships`)
            .get()
        const appRelationships: MobileAppRelationship[] = res.value

        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextAppsRelationships = res.value as MobileAppRelationship[]
            appRelationships.push(...nextAppsRelationships)
        }
        return appRelationships
    }

    async getAppRelationship(
        appId: string,
        relationshipId: string,
    ): Promise<MobileAppRelationship> {
        return this.graphClient
            .api(`/deviceAppManagement/mobileApps/${appId}/relationships/${relationshipId}`)
            .get()
    }

    //Forced to use the updateRelationship endpoint since the relationship post method is not working(https://docs.microsoft.com/en-us/graph/api/intune-apps-mobileappdependency-create?view=graph-rest-beta)
    async createAppRelationship(
        appId: string,
        relationship: MobileAppRelationship,
    ): Promise<MobileAppRelationship> {
        const relationships = (await this.listAppRelationships(appId)) ?? []
        relationships.push(relationship)
        await this.graphClient
            .api(`/deviceAppManagement/mobileApps/${appId}/updateRelationships`)
            .post({ relationships: relationships })
        const updatedRelationships = await this.listAppRelationships(appId)
        const res = updatedRelationships.find(
            (appRelationship) => appRelationship.targetId === relationship.targetId,
        ) as MobileAppRelationship
        return res
    }

    async deleteAppRelationship(appId: string, relationshipId: string): Promise<void> {
        return this.graphClient
            .api(`/deviceAppManagement/mobileApps/${appId}/relationships/${relationshipId}`)
            .delete()
    }

    async removeAllAppRelationships(appId: string): Promise<void> {
        await this.graphClient
            .api(`/deviceAppManagement/mobileApps/${appId}/updateRelationships`)
            .post({ relationships: [] })
    }

    async updateAppRelationship(
        appId: string,
        relationshipId: string,
        relationship: MobileAppRelationship,
    ): Promise<MobileAppRelationship> {
        return this.graphClient
            .api(`/deviceAppManagement/mobileApps/${appId}/relationships/${relationshipId}`)
            .patch(relationship)
    }
    async createAssignment(
        appId: string,
        assignment: MobileAppAssignment,
    ): Promise<MobileAppAssignment> {
        return this.graphClient
            .api(`/deviceAppManagement/mobileApps/${appId}/assignments`)
            .post(assignment)
    }

    async listAssignments(appId: string): Promise<MobileAppAssignment[]> {
        let res = await this.graphClient
            .api(`/deviceAppManagement/mobileApps/${appId}/assignments`)
            .get()
        const assignments: MobileAppAssignment[] = res.value
        while (res['@odata.nextLink']) {
            const nextLink = res['@odata.nextLink'].replace('https://graph.microsoft.com/beta', '')
            res = await this.graphClient.api(nextLink).get()
            const nextAssignments = res.value as MobileAppAssignment[]
            assignments.push(...nextAssignments)
        }
        return assignments
    }

    async deleteAssignment(appId: string, assignmentId: string): Promise<void> {
        await this.graphClient
            .api(`/deviceAppManagement/mobileApps/${appId}/assignments/${assignmentId}`)
            .delete()
    }

    async getAssignment(appId: string, assignmentId: string): Promise<MobileAppAssignment> {
        return this.graphClient
            .api(`/deviceAppManagement/mobileApps/${appId}/assignments/${assignmentId}`)
            .get()
    }
}
