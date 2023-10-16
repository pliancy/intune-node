import axios from 'axios'
import { Config } from '../types'
import qs from 'qs'

export class AuthProvider {
    private _refreshToken = this.config?.authentication?.refreshToken

    constructor(private readonly config: Config) {}

    get refreshToken(): string | undefined {
        return this._refreshToken
    }

    async getAccessToken(): Promise<string> {
        const url = `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/token`

        const scope = 'https://graph.microsoft.com/.default'

        const options = {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
            },
        }

        let response: {
            data: {
                access_token: string
                refresh_token?: string
            }
        }

        if (this._refreshToken) {
            response = await axios.post(
                url,
                qs.stringify({
                    client_id: this.config.authentication.clientId,
                    client_secret: this.config.authentication.clientSecret,
                    grant_type: 'refresh_token',
                    refresh_token: this._refreshToken,
                    scope,
                }),
                options,
            )
            // Update refresh token if it exists
            this._refreshToken = response?.data?.refresh_token
                ? response?.data?.refresh_token
                : this._refreshToken
        } else {
            response = await axios.post(
                url,
                qs.stringify({
                    client_id: this.config.authentication.clientId,
                    client_secret: this.config.authentication.clientSecret,
                    grant_type: 'client_credentials',
                    scope,
                }),
                options,
            )
        }
        return response?.data?.access_token
    }
}
