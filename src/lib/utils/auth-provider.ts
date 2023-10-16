import axios from 'axios'
import { Config } from '../types'
import qs from 'qs'

export class AuthProvider {
    constructor(private readonly config: Config) {}

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
            }
        }

        if (this.config.authentication?.refreshToken) {
            response = await axios.post(
                url,
                qs.stringify({
                    client_id: this.config.authentication.clientId,
                    client_secret: this.config.authentication.clientSecret,
                    grant_type: 'refresh_token',
                    refresh_token: this.config.authentication.refreshToken,
                    scope,
                }),
                options,
            )
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
