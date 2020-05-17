import Progress from 'cli-progress'
import chalk from 'chalk'
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
  clientId: string
  clientSecret: string
  tenantId: string
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

  constructor(_config: IntuneConfig) {
    this.config = _config
    this.domain = `https://graph.microsoft.com/beta`
    this.accessToken = ''
    this.reqHeaders = {
      authorization: `Bearer ${this.accessToken}`,
      'content-type': 'application/json',
      accept: 'application/json'
    }
  }

  async getIntuneDevices (): Promise<object[]> {
    let res = await this._IntuneRequest(
      `${this.domain}/deviceManagement/managedDevices?$top=999`,
      {
        method: 'GET',
        headers: this.reqHeaders
      }
    )
    const resbody = JSON.parse(res.body)
    return resbody.value
  }

  async getAzureAdDevices (): Promise<object[]> {
    let res = await this._IntuneRequest(`${this.domain}/devices?$top=999`, {
      method: 'GET',
      headers: this.reqHeaders
    })
    const resbody = JSON.parse(res.body)
    return resbody.value
  }

  async getApps (): Promise<object[]> {
    let res = await this._IntuneRequest(
      `${this.domain}/deviceAppManagement/mobileApps?$top=999`,
      {
        method: 'GET',
        headers: this.reqHeaders
      }
    )
    const resbody = JSON.parse(res.body)
    return resbody.value
  }

  async getDeviceConfigurations (): Promise<object[]> {
    let res = await this._IntuneRequest(
      `${this.domain}/deviceManagement/deviceConfigurations?$top=999`,
      {
        method: 'GET',
        headers: this.reqHeaders
      }
    )
    const resbody = JSON.parse(res.body)
    return resbody.value
  }

  async createDeviceConfiguration (postBody: object): Promise<object> {
    let res = await this._IntuneRequest(
      `${this.domain}/deviceManagement/deviceConfigurations`,
      {
        method: 'POST',
        headers: this.reqHeaders,
        body: JSON.stringify(postBody)
      }
    )
    const resbody = JSON.parse(res.body)
    return resbody
  }

  async createApp (postBody: object): Promise<object> {
    try {
      let res = await this._IntuneRequest(
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
      let resArray: object[] = []
      let res = await this._IntuneRequest(
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
      const appId = resBody.id

      definitionValues.map(async definitionValue => {
        let res = await this._IntuneRequest(
          `${this.domain}/deviceManagement/groupPolicyConfigurations/${appId}/definitionValues`,
          {
            method: 'POST',
            headers: this.reqHeaders,
            body: JSON.stringify(definitionValue)
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
      let res = await this._IntuneRequest(
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
      let res = await this._IntuneRequest(
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
    const delay = (ms: any) => new Promise(resolve => setTimeout(resolve, ms))
    try {
      while (loop) {
        let res = await this._IntuneRequest(
          `${this.domain}/deviceAppManagement/mobileApps/${appId}/microsoft.graph.win32LobApp/contentversions/${contentVersionId}/files/${fileId}`,
          {
            method: 'GET',
            headers: this.reqHeaders
          }
        )
        const resBody = await JSON.parse(res.body)

        if (resBody.azureStorageUri) {
          azureStorageUri = resBody.azureStorageUri
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
    fileSize: number
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

    const bar = new Progress.SingleBar(
      {
        format: `Uploading ${chalk.bold.hex('#71c6e5')('Package')} ${chalk.hex(
          '#7F4BAE'
        )('{bar}')} {percentage}%`
      },
      Progress.Presets.shades_grey
    )
    bar.start(fileSize, 0)

    const uploadBlobResponse = await blockBlobClient.uploadStream(
      file,
      bufferSize,
      20,
      { onProgress: (ev: any) => bar.update(ev.loadedBytes) }
    )
    bar.stop()
    console.log(chalk.bold.hex('#008000')('File Uploaded'))
    return uploadBlobResponse
  }

  async commitFileUpload (
    appId: string,
    contentVersionId: number,
    fileId: string,
    postBody: object
  ): Promise<object> {
    try {
      let res = await this._IntuneRequest(
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
      let res = await this._IntuneRequest(
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
  ) {
    let loop = true
    const delay = (ms: any) => new Promise(resolve => setTimeout(resolve, ms))
    try {
      while (loop) {
        let res = await this._IntuneRequest(
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
      }
    } catch (err) {
      throw err
    }
    return
  }

  async createWin32app (
    appCreationBody: object,
    encryptionBody: object,
    fileInfoBody: any,
    file: any
  ) {
    const arr: any = []

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
      console.log(chalk.bold.hex('#71c6e5')('Waiting for Storage URI'))
      const azureStorageUri = await this.getAzureStorageUri(
        appId,
        contentVersionId,
        fileId
      )
      const fileSize = fileInfoBody.size

      const blobUploadRes = await this.uploadToAzureBlob(
        azureStorageUri,
        file,
        fileSize
      )

      const commitFileRes = await this.commitFileUpload(
        appId,
        contentVersionId,
        fileId,
        encryptionBody
      )
      console.log(
        chalk.bold.hex('#71c6e5')('Waiting for Successful File Upload Status')
      )

      const fileUploadStatusRes = await this.getFileUploadStatus(
        appId,
        contentVersionId,
        fileId
      )

      const commitAppRes = await this.commitApp(appId, contentVersionId)
      return chalk.bold.hex('#008000')('App Creation Successful')
    } catch (err) {
      throw err
    }
  }

  async updateWin32AppFile (
    appId: string,
    encryptionBody: object,
    fileInfoBody: any,
    file: any
  ) {
    try {
      const fileInfoRes: any = await this.createContentVersion(appId)

      const contentVersionId = fileInfoRes.id
      const fileUploadRes: any = await this.createFileUpload(
        appId,
        contentVersionId,
        fileInfoBody
      )

      const fileId = fileUploadRes.id
      console.log(chalk.bold.hex('#71c6e5')('Waiting for Storage URI'))
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
      console.log(
        chalk.bold.hex('#71c6e5')('Waiting for Successful File Upload Status')
      )

      await this.getFileUploadStatus(appId, contentVersionId, fileId)
    } catch (err) {
      throw err
    }
  }

  private async _authenticate (): Promise<string> {
    let res = await got(
      `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/token`,
      {
        method: 'post',
        headers: {
          'content-type': 'application/x-www-form-urlencoded'
        },
        body: qs.stringify({
          grant_type: 'client_credentials',
          scope: 'https://graph.microsoft.com/.default',
          client_id: this.config.clientId,
          client_secret: this.config.clientSecret
        })
      }
    )
    let body: IOAuthResponse = JSON.parse(res.body)
    this.accessToken = body.access_token
    return body.access_token
  }

  private async _IntuneRequest (url: string, options: any): Promise<any> {
    try {
      if (!this.accessToken) {
        let token = await this._authenticate()
        options.headers.Authorization = `Bearer ${token}`
      } else {
        options.headers.Authorization = `Bearer ${this.accessToken}`
      }
      let res = await got(url, options)

      return res
    } catch (err) {
      if (err.statusCode === 401) {
        let token = await this._authenticate()
        options.headers.Authorization = `Bearer ${token}`
        let res = await got(url, options)
        return res
      }
      throw err
    }
  }
}

export = Intune
