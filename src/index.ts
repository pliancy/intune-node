import got from 'got'
import qs from 'qs'
import {
  BlobServiceClient,
  AnonymousCredential,
  newPipeline
} from '@azure/storage-blob'
import { IntuneConfig, ClientAuth, BearerAuth, IOAuthResponse, IntuneScript, AutoPilotUpload, IntuneDeviceResponse, AzureDeviceResponse } from './types'

/**
 * The Config for the Intune class
 * This interface allows utilization of intune API
 * @export
 * @interface IntuneConfig
 */

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

  async request (method: string, endpoint: string, body?: object): Promise<object[]> {
    try {
      const req: any = {
        method: method,
        headers: this.reqHeaders
      }

      if (typeof body !== 'undefined') {
        req.body = body
      }

      const res = await this._IntuneRequest(`${this.domain}/${endpoint}`, req)
      const resbody = JSON.parse(res.body)
      return resbody
    } catch (err) {
      throw err
    }
  }

  async getIntuneDevices (): Promise<IntuneDeviceResponse[]> {
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

  async getAutopilotDevices (): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/windowsAutopilotDeviceIdentities?$top=999`,
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

  async syncDevice (deviceId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/managedDevices/${deviceId}/syncDevice`,
        {
          method: 'POST',
          headers: this.reqHeaders
        }
      )
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async autoPilotSync (): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/windowsAutopilotSettings/sync`,
        {
          method: 'POST',
          headers: this.reqHeaders
        }
      )
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async autopilotUpload ({ serialNumber, groupTag, productKey, hardwareIdentifier, assignedUser }: AutoPilotUpload): Promise<object> {
    try {
      const postBody = {
        '@odata.type': '#microsoft.graph.importedWindowsAutopilotDeviceIdentity',
        orderIdentifier: groupTag ?? null,
        serialNumber: serialNumber ?? null,
        productKey: productKey ?? null,
        hardwareIdentifier: hardwareIdentifier ?? null,
        assignedUserPrincipalName: assignedUser ?? null
      }
      console.log(postBody)
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/importedWindowsAutopilotDeviceIdentities`,
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

  async setDeviceName (deviceId: string, newDeviceName: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/managedDevices/${deviceId}/setDeviceName`,
        {
          method: 'POST',
          headers: this.reqHeaders,
          body: JSON.stringify({
            deviceName: newDeviceName
          })
        }
      )
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async rebootDevice (deviceId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/managedDevices/${deviceId}/rebootNow`,
        {
          method: 'POST',
          headers: this.reqHeaders
        }
      )
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async shutDownDevice (deviceId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/managedDevices/${deviceId}/shutDown`,
        {
          method: 'POST',
          headers: this.reqHeaders
        }
      )
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async getScripts (): Promise<[object]> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/deviceManagementScripts`,
        {
          method: 'Get',
          headers: this.reqHeaders
        }
      )
      const resbody = JSON.parse(res.body)
      return resbody.value
    } catch (err) {
      throw err
    }
  }

  async getScript (scriptId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/deviceManagementScripts/${scriptId}`,
        {
          method: 'Get',
          headers: this.reqHeaders
        }
      )
      const resbody = JSON.parse(res.body)
      return resbody
    } catch (err) {
      throw err
    }
  }

  async createScript ({ displayName, description, scriptContent, fileName, runAsAccount, runAs32Bit = false, enforceSignatureCheck = false }: IntuneScript): Promise<object> {
    try {
      const postBody = {
        '@odata.type': '#microsoft.graph.deviceManagementScript',
        displayName: displayName,
        description: description,
        scriptContent: scriptContent,
        runAsAccount: runAsAccount,
        enforceSignatureCheck: enforceSignatureCheck,
        fileName: fileName,
        runAs32Bit: runAs32Bit
      }

      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/deviceManagementScripts`,
        {
          method: 'Post',
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

  async updateScript (scriptId: string, { displayName, description, scriptContent, fileName, runAsAccount, runAs32Bit = false, enforceSignatureCheck = false }: IntuneScript): Promise<object> {
    try {
      const postBody = {
        '@odata.type': '#microsoft.graph.deviceManagementScript',
        displayName: displayName,
        description: description,
        scriptContent: scriptContent,
        runAsAccount: runAsAccount,
        enforceSignatureCheck: enforceSignatureCheck,
        fileName: fileName,
        runAs32Bit: runAs32Bit
      }

      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/deviceManagementScripts${scriptId}`,
        {
          method: 'Patch',
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

  async deleteScript (scriptId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/deviceManagementScripts/${scriptId}`,
        {
          method: 'Delete',
          headers: this.reqHeaders
        }
      )
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async createScriptAssignment (scriptId: string, groupIds: string[]): Promise<object> {
    try {
      const groupAssignments = groupIds.map((e: string) => {
        return {
          '@odata.type': '#microsoft.graph.deviceManagementScriptGroupAssignment',
          targetGroupId: e
        }
      })

      const postBody = {
        deviceManagementScriptGroupAssignments: groupAssignments
      }

      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/deviceManagementScripts/${scriptId}/assign`,
        {
          method: 'Post',
          headers: this.reqHeaders,
          body: JSON.stringify(postBody)
        }
      )
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async getDevice (deviceId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/managedDevices/${deviceId}`,
        {
          method: 'Get',
          headers: this.reqHeaders
        }
      )
      const resbody = JSON.parse(res.body)
      return resbody
    } catch (err) {
      throw err
    }
  }

  async updateDevice (deviceId: string, patchBody: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/managedDevices/${deviceId}`,
        {
          method: 'PATCH',
          headers: this.reqHeaders,
          body: JSON.stringify(patchBody)
        }
      )

      const resbody = JSON.parse(res.body)
      return resbody
    } catch (err) {
      throw err
    }
  }

  async wipeDevice (deviceId: string, keepEnrollmentData: boolean, keepUserData: boolean, useProtectedWipe: boolean, macOsUnlockCode?: string): Promise<object> {
    try {
      const postbody: any = {
        keepEnrollmentData: keepEnrollmentData,
        keepUserData: keepUserData,
        useProtectedWipe: useProtectedWipe
      }

      if (macOsUnlockCode !== undefined) {
        postbody.macOsUnlockCode = macOsUnlockCode
      }

      const res = await this._IntuneRequest(
        `${this.domain}/deviceManagement/managedDevices/${deviceId}/wipe`,
        {
          method: 'POST',
          headers: this.reqHeaders,
          body: JSON.stringify(postbody)
        }
      )
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async getAzureAdDevices (): Promise<AzureDeviceResponse[]> {
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

  async createUser (postBody: object): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/users`, {
        method: 'POST',
        headers: this.reqHeaders,
        body: JSON.stringify(postBody)
      })
      const resbody = JSON.parse(res.body)
      return resbody
    } catch (err) {
      throw err
    }
  }

  async updateUser (userID: string, postBody: object): Promise<object> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/users/${userID}`, {
        method: 'PATCH',
        headers: this.reqHeaders,
        body: JSON.stringify(postBody)
      })
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async getUsers (): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/users?$top=999`, {
        method: 'GET',
        headers: this.reqHeaders
      })
      const resbody = JSON.parse(res.body)
      return resbody.value
    } catch (err) {
      throw err
    }
  }

  async getUser (userId: string): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/users/${userId}`, {
        method: 'GET',
        headers: this.reqHeaders
      })
      const resbody = JSON.parse(res.body)
      return resbody
    } catch (err) {
      throw err
    }
  }

  async getUserAppsIntentandState (userId: string): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/users/${userId}/mobileAppIntentAndStates`, {
        method: 'GET',
        headers: this.reqHeaders
      })
      const resbody = JSON.parse(res.body)
      return resbody.value
    } catch (err) {
      throw err
    }
  }

  async getUserAppIntentandState (userId: string, appIntentAndStateId: string): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/users/${userId}/mobileAppIntentAndStates/${appIntentAndStateId}`, {
        method: 'GET',
        headers: this.reqHeaders
      })
      const resbody = JSON.parse(res.body)
      return resbody
    } catch (err) {
      throw err
    }
  }

  async deleteUser (userId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/users/${userId}`, {
        method: 'DELETE',
        headers: this.reqHeaders
      })
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async deleteDevice (deviceId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/deviceManagement/managedDevices/${deviceId}`, {
        method: 'DELETE',
        headers: this.reqHeaders
      })
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async getGroups (): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/groups?$top=999`, {
        method: 'GET',
        headers: this.reqHeaders
      })
      const resbody = JSON.parse(res.body)
      return resbody.value
    } catch (err) {
      throw err
    }
  }

  async getGroup (groupId: string): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/groups/${groupId}`, {
        method: 'GET',
        headers: this.reqHeaders
      })
      const resbody = JSON.parse(res.body)
      return resbody
    } catch (err) {
      throw err
    }
  }

  async getGroupMembers (groupId: string): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/groups/${groupId}/members?$top=999`, {
        method: 'GET',
        headers: this.reqHeaders
      })
      const resbody = JSON.parse(res.body)
      return resbody.value
    } catch (err) {
      throw err
    }
  }

  async getGroupOwners (groupId: string): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/groups/${groupId}/owners?$top=999`, {
        method: 'GET',
        headers: this.reqHeaders
      })
      const resbody = JSON.parse(res.body)
      return resbody.value
    } catch (err) {
      throw err
    }
  }

  async addGroupMember (groupId: string, memberId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/groups/${groupId}/members/$ref`, {
        method: 'POST',
        headers: this.reqHeaders,
        body: JSON.stringify({
          '@odata.id': `https://graph.microsoft.com/v1.0/users/${memberId}`
        })
      })
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async addGroupOwner (groupId: string, memberId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/groups/${groupId}/owners/$ref`, {
        method: 'POST',
        headers: this.reqHeaders,
        body: JSON.stringify({
          '@odata.id': `https://graph.microsoft.com/v1.0/users/${memberId}`
        })
      })
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async removeGroupMember (groupId: string, userId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/groups/${groupId}/members/${userId}/$ref`, {
        method: 'DELETE',
        headers: this.reqHeaders
      })
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async removeGroupOwner (groupId: string, userId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/groups/${groupId}/owners/${userId}/$ref`, {
        method: 'DELETE',
        headers: this.reqHeaders
      })
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async createGroup (postBody: object): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/groups`, {
        method: 'POST',
        headers: this.reqHeaders,
        body: JSON.stringify(postBody)
      })
      const resbody = JSON.parse(res.body)
      return resbody
    } catch (err) {
      throw err
    }
  }

  async deleteGroup (groupId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(`${this.domain}/groups/${groupId}`, {
        method: 'DELETE',
        headers: this.reqHeaders
      })
      return { statusCode: res.statusCode }
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

  async getApp (appId: string): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps/${appId}/`,
        {
          method: 'GET',
          headers: this.reqHeaders
        }
      )
      const resbody = JSON.parse(res.body)
      return resbody
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
        `${this.domain}/deviceAppManagement/mobileApps/${appId}/relationships?$top=999`,
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

  async setAppDependencies (appId: string, postBody: object): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps/${appId}/updateRelationships`,
        {
          method: 'POST',
          headers: this.reqHeaders,
          body: JSON.stringify(postBody)
        }
      )
      return { statusCode: res.statusCode }
    } catch (err) {
      throw err
    }
  }

  async getAppAssignments (appId: string): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps/${appId}/assignments?$top=999`,
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

  async getAppAssignment (appId: string, mobileAppAssignmentId: string): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps/${appId}/assignments/${mobileAppAssignmentId}`,
        {
          method: 'GET',
          headers: this.reqHeaders
        }
      )
      const resbody = JSON.parse(res.body)
      return resbody
    } catch (err) {
      throw err
    }
  }

  async createAppAssignment (appId: string, postBody: object): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps/${appId}/assignments/`,
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

  async updateAppAssignment (appId: string, mobileAppAssignmentId: string, postBody: object): Promise<object[]> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps/${appId}/assignments/${mobileAppAssignmentId}`,
        {
          method: 'PATCH',
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

  async deleteAppAssignment (appId: string, mobileAppAssignmentId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps/${appId}/assignments/${mobileAppAssignmentId}`,
        {
          method: 'DELETE',
          headers: this.reqHeaders
        }
      )
      return { statusCode: res.statusCode }
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

  async deleteApp (appId: string): Promise<object> {
    try {
      const res = await this._IntuneRequest(
        `${this.domain}/deviceAppManagement/mobileApps/${appId}`,
        {
          method: 'DELETE',
          headers: this.reqHeaders
        }
      )

      return { statusCode: res.statusCode }
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
    statusCallback?: any,
    customBufferSize?: number
  ): Promise<Object> {
    let bufferSize: number = customBufferSize ?? 4 * 1024 * 1024
    if (fileSize < 4000 && typeof customBufferSize === 'undefined') {
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
    let uploadBlobResponse: object = {}
    if (typeof statusCallback !== 'undefined') {
      uploadBlobResponse = await blockBlobClient.uploadStream(
        file,
        bufferSize,
        5,
        {
          onProgress: (ev: any) =>
            statusCallback(fileSize, ev.loadedBytes)
        }
      )
    } else {
      uploadBlobResponse = await blockBlobClient.uploadStream(
        file,
        bufferSize,
        5
      )
    }
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
