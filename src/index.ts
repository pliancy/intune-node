import got from "got"
import qs from "qs"

/**
 * The Config for the Intune class
 * This interface allows utilization of intune API
 * @export
 * @interface IntuneConfig
 */
export interface IntuneConfig {
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

export class Intune {
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

  async getIntuneDevices(): Promise<object> {
    let res = await this._IntuneRequest(
      `${this.domain}/deviceManagement/managedDevices?$top=999`,
      {
        method: "GET",
        headers: this.reqHeaders
      }
    )
    return JSON.parse(res.body)
  };

  async getAzureAdDevices(): Promise<object> {
    let res = await this._IntuneRequest(
      `${this.domain}/devices?$top=999`,
      {
        method: "GET",
        headers: this.reqHeaders
      }
    )
    return JSON.parse(res.body)
  };





  private async _authenticate(): Promise<string> {
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

  private async _IntuneRequest(url: string, options: any): Promise<any> {
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