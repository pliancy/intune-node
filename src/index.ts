import got from 'got'
import qs from 'qs'
import {
  BlobServiceClient,
  AnonymousCredential,
  newPipeline
} from '@azure/storage-blob'

/**
 * The Config for the Intune class
 * This interface allows utilization of intune API
 * @export
 * @interface IntuneConfig
 */
interface IntuneConfig {
  tenantId: string
  authentication: ClientAuth | BearerAuth
}

interface ClientAuth {
  clientId: string
  clientSecret: string
}

interface BearerAuth {
  bearerToken: string
}
interface IOAuthResponse {
  token_type: string
  expires_in: string
  ext_expires_in: string
  expires_on: string
  not_before: string
  resource: string
  access_token: string
}

class Intune {
  config: IntuneConfig
  domain: string
  accessToken: string
  reqHeaders: any

  constructor (_config: IntuneConfig) {
    this.config = _config
    this.domain = 'https://graph.microsoft.com/beta'
    this.accessToken = ''
    this.reqHeaders = {
      authorization: `Bearer ${this.accessToken}`,
      'content-type': 'application/json',
      accept: 'application/json'
    }
  }

  /* eslint-disable no-useless-catch */

  async getIntuneDevices (): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/managedDevices?$top=999`,
        {
          method: 'GET',
          headers: this.reqHeaders
        }
      )
      const resbody = JSON.parse(res.body)
      return resbody.value
    } catch (err) {
      throw err
    }
  }

  async getAzureAdDevices (): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/devices?$top=999`, {
        method: 'GET',
        headers: this.reqHeaders
      })
      const resbody = JSON.parse(res.body)
      return resbody.value
    } catch (err) {
      throw err
    }
  }

  async getApps (): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps?$top=999`,
        {
          method: 'GET',
          headers: this.reqHeaders
        }
      )
      const resbody = JSON.parse(res.body)
      return resbody.value
    } catch (err) {
      throw err
    }
  }

  async getDeviceConfigurations (): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/deviceConfigurations?$top=999`,
        {
          method: 'GET',
          headers: this.reqHeaders
        }
      )
      const resbody = JSON.parse(res.body)
      return resbody.value
    } catch (err) {
      throw err
    }
  }

  async getAppDependencies (appId: string): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps/${appId}/relationships`,
        {
          method: 'GET',
          headers: this.reqHeaders
        }
      )
      const resbody = JSON.parse(res.body)
      return resbody.value
    } catch (err) {
      throw err
    }
  }

  async createAppDependency (appId: string, dependencyAppId: string, autoInstall: boolean): Promise<any> {
    try {
      let dependencyType = 'autoInstall'
      if (!autoInstall) {
        dependencyType = 'detect'
      }
      await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps/${appId}/updateRelationships`,
        {
          method: 'POST',
          headers: this.reqHeaders,
          body: JSON.stringify(
            {
              relationships: [
                {
                  '@odata.type': '#microsoft.graph.mobileAppDependency',
                  targetId: dependencyAppId,
                  dependencyType: dependencyType
                }
              ]
            })
        }
      )
      return
    } catch (err) {
      throw err
    }
  }

  async createDeviceConfiguration (postBody: object): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/deviceConfigurations`,
        {
          method: 'POST',
          headers: this.reqHeaders,
          body: JSON.stringify(postBody)
        }
      )
      const resbody = JSON.parse(res.body)
      return resbody
    } catch (err) {
      throw err
    }
  }

  async createApp (postBody: object): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps`,
        {
          method: 'POST',
          headers: this.reqHeaders,
          body: JSON.stringify(postBody)
        }
      )

      return JSON.parse(res.body)
    } catch (err) {
      throw err
    }
  }

  async createDeviceConfigurationAdmx (
    displayName: string,
    description: string,
    definitionValues: object[]
  ): Promise<object[]> {
    try {
      const resArray: object[] = []
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/groupPolicyConfigurations`,
        {
          method: 'POST',
          headers: this.reqHeaders,
          body: JSON.stringify({
            description: description,
            displayName: displayName
          })
        }
      )
      const resBody = JSON.parse(res.body)
      resArray.push(resBody)
      const appId: string = resBody.id

      definitionValues.map(async e => {
        const res = await this._IntuneRequest(
          `${this.domain}/deviceManagement/groupPolicyConfigurations/${appId}/definitionValues`,
          {
            method: 'POST',
            headers: this.reqHeaders,
            body: JSON.stringify(e)
          }
        )
        const resBody = JSON.parse(res.body)
        resArray.push(resBody)
      })
      return resArray
    } catch (err) {
      throw err
    }
  }

  async createContentVersion (appId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps/${appId}/microsoft.graph.win32LobApp/contentVersions`,
        {
          body: JSON.stringify({}),
          method: 'POST',
          headers: this.reqHeaders
        }
      )
      return JSON.parse(res.body)
    } catch (err) {
      throw err
    }
  }

  async createFileUpload (
    appId: string,
    contentVersionId: number,
    postBody: object
  ): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps/${appId}/microsoft.graph.win32LobApp/contentversions/${contentVersionId}/files`,
        {
          method: 'POST',
          headers: this.reqHeaders,
          body: JSON.stringify(postBody)
        }
      )

      return JSON.parse(res.body)
    } catch (err) {
      throw err
    }
  }

  async getAzureStorageUri (
    appId: string,
    contentVersionId: number,
    fileId: string
  ): Promise<string> {
    let azureStorageUri: string = ''
    let loop = true
    const delay = async (ms: any): Promise<object> =>
      await new Promise(resolve => setTimeout(resolve, ms))
    try {
      while (loop) {
        const res = await this._IntuneRequest(
          `${this.domain}/deviceAppManagement/mobileApps/${appId}/microsoft.graph.win32LobApp/contentversions/${contentVersionId}/files/${fileId}`,
          {
            method: 'GET',
            headers: this.reqHeaders
          }
        )
        const resBody = await JSON.parse(res.body)
        azureStorageUri = resBody.azureStorageUri
        if (azureStorageUri !== null) {
          loop = false
        } else {
          await delay(1000)
        }
      }
    } catch (err) {
      throw err
    }
    return azureStorageUri
  }

  async uploadToAzureBlob (
    azureStorageUri: string,
    file: any,
    fileSize: number,
    statusCallback?: any
  ): Promise<Object> {
    // Parse azureStorageUri
    let bufferSize: number = 1 * 1024
    if (fileSize < 4000) {
      const buffer: number = 1 * 1024
      bufferSize = buffer
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

    const uploadBlobResponse = await blockBlobClient.uploadStream(
      file,
      bufferSize,
      20,
      {
        onProgress: (ev: any) =>
          statusCallback(fileSize, ev.loadedBytes) ?? null
      }
    )
    return uploadBlobResponse
  }

  async commitFileUpload (
    appId: string,
    contentVersionId: number,
    fileId: string,
    postBody: object
  ): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps/${appId}/microsoft.graph.win32LobApp/contentversions/${contentVersionId}/files/${fileId}/commit`,
        {
          method: 'POST',
          headers: this.reqHeaders,
          body: JSON.stringify(postBody)
        }
      )
      return res.body
    } catch (err) {
      throw err
    }
  }

  async commitApp (appId: string, contentVersionId: number): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps/${appId}`,
        {
          method: 'PATCH',
          headers: this.reqHeaders,
          body: JSON.stringify({
            '@odata.type': '#microsoft.graph.win32LobApp',
            committedContentVersion: `${contentVersionId}`
          })
        }
      )
      return res.body
    } catch (err) {
      throw err
    }
  }

  async getFileUploadStatus (
    appId: string,
    contentVersionId: number,
    fileId: string
  ): Promise<string> {
    let loop = true
    const delay = async (ms: any): Promise<object> =>
      await new Promise(resolve => setTimeout(resolve, ms))
    let uploadState = ''
    try {
      while (loop) {
        const res = await this._IntuneRequest(
          `${this.domain}/deviceAppManagement/mobileApps/${appId}/microsoft.graph.win32LobApp/contentversions/${contentVersionId}/files/${fileId}`,
          {
            method: 'GET',
            headers: this.reqHeaders
          }
        )
        const resBody = await JSON.parse(res.body)

        if (resBody.uploadState === 'commitFileSuccess') {
          loop = false
        } else {
          await delay(1000)
        }
        uploadState = resBody.uploadState
      }
      return uploadState
    } catch (err) {
      throw err
    }
  }

  async createWin32app (
    appCreationBody: object,
    encryptionBody: object,
    fileInfoBody: any,
    file: any
  ): Promise<object> {
    try {
      const appCreationRes: any = await this.createApp(appCreationBody)

      const appId = appCreationRes.id

      const fileInfoRes: any = await this.createContentVersion(appId)

      const contentVersionId = fileInfoRes.id
      const fileUploadRes: any = await this.createFileUpload(
        appId,
        contentVersionId,
        fileInfoBody
      )

      const fileId = fileUploadRes.id
      const azureStorageUri = await this.getAzureStorageUri(
        appId,
        contentVersionId,
        fileId
      )
      const fileSize = fileInfoBody.size

      await this.uploadToAzureBlob(azureStorageUri, file, fileSize)

      await this.commitFileUpload(
        appId,
        contentVersionId,
        fileId,
        encryptionBody
      )

      await this.getFileUploadStatus(appId, contentVersionId, fileId)

      await this.commitApp(appId, contentVersionId)
      return appCreationRes
    } catch (err) {
      throw err
    }
  }

  async updateWin32AppFile (
    appId: string,
    encryptionBody: object,
    fileInfoBody: any,
    file: any
  ): Promise<string> {
    try {
      const fileInfoRes: any = await this.createContentVersion(appId)

      const contentVersionId = fileInfoRes.id
      const fileUploadRes: any = await this.createFileUpload(
        appId,
        contentVersionId,
        fileInfoBody
      )

      const fileId = fileUploadRes.id
      const azureStorageUri = await this.getAzureStorageUri(
        appId,
        contentVersionId,
        fileId
      )
      const fileSize = fileInfoBody.size

      await this.uploadToAzureBlob(azureStorageUri, file, fileSize)

      await this.commitFileUpload(
        appId,
        contentVersionId,
        fileId,
        encryptionBody
      )

      await this.getFileUploadStatus(appId, contentVersionId, fileId)
      return `${appId} Updated`
    } catch (err) {
      throw err
    }
  }

  isClientAuth = (e: ClientAuth | BearerAuth): e is ClientAuth => {
    return (e as ClientAuth).clientId !== undefined
  }

  isBearerAuth = (e: ClientAuth | BearerAuth): e is BearerAuth => {
    return (e as BearerAuth).bearerToken !== undefined
  }

  private async _authenticate (): Promise<string | undefined> {
    if (this.isClientAuth(this.config.authentication)) {
      const res = await got(
        `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/token`,
        {
          method: 'post',
          headers: {
            'content-type': 'application/x-www-form-urlencoded'
          },
          body: qs.stringify({
            grant_type: 'client_credentials',
            scope: 'https://graph.microsoft.com/.default',
            client_id: this.config.authentication.clientId,
            client_secret: this.config.authentication.clientSecret
          })
        }
      )
      const body: IOAuthResponse = JSON.parse(res.body)
      this.accessToken = body.access_token
      return body.access_token
    }
  }

  private async _IntuneRequest (url: string, options: any): Promise<any> {
    if (this.isClientAuth(this.config.authentication)) {
      try {
        if (this.accessToken === '') {
          const token = await this._authenticate()
          if (typeof token === 'string') {
            options.headers.Authorization = `Bearer ${token}`
          }
        } else {
          options.headers.Authorization = `Bearer ${this.accessToken}`
        }
        const res = await got(url, options)
        return res
      } catch (err) {
        if (err.statusCode === 401) {
          const token = await this._authenticate()
          if (typeof token === 'string') {
            options.headers.Authorization = `Bearer ${token}`
          }
          const res = await got(url, options)
          return res
        }
        throw err
      }
    } else if (this.isBearerAuth(this.config.authentication)) {
      try {
        if (this.accessToken === '') {
          this.accessToken = this.config.authentication.bearerToken
          options.headers.Authorization = `Bearer ${this.accessToken}`
        }
        const res = await got(url, options)
        return res
      } catch (err) {
        if (err.statusCode === 401) {
          options.headers.Authorization = `Bearer ${this.config.authentication.bearerToken}`
          const res = await got(url, options)
          return res
        }
        throw (err)
      }
    }
  }
}

export = Intune
