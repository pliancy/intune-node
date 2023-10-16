import axios from 'axios'
import qs from 'qs'
import { AuthProvider } from './auth-provider'

jest.mock('axios')
const mockedAxios = axios as jest.Mocked<typeof axios>

describe('AuthProvider', () => {
    afterEach(() => {
        jest.clearAllMocks()
    })

    it('should get access token using refresh token if provided', async () => {
        const mockConfig = {
            tenantId: 'mockTenantId',
            authentication: {
                clientId: 'mockClientId',
                clientSecret: 'mockClientSecret',
                refreshToken: 'mockRefreshToken',
            },
        }

        mockedAxios.post.mockResolvedValue({
            data: {
                access_token: 'mockAccessToken',
                refresh_token: 'mockUpdatedRefreshToken',
            },
        })

        const authProvider = new AuthProvider(mockConfig)

        // Check that refresh token is set with the value from config
        expect(authProvider.refreshToken).toBe('mockRefreshToken')

        // Get access token
        const result = await authProvider.getAccessToken()

        expect(result).toBe('mockAccessToken')
        expect(mockedAxios.post).toHaveBeenCalledWith(
            `https://login.microsoftonline.com/${mockConfig.tenantId}/oauth2/v2.0/token`,
            qs.stringify({
                client_id: mockConfig.authentication.clientId,
                client_secret: mockConfig.authentication.clientSecret,
                grant_type: 'refresh_token',
                refresh_token: mockConfig.authentication.refreshToken,
                scope: 'https://graph.microsoft.com/.default',
            }),
            {
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            },
        )

        // Check that refresh token is updated with the value from response
        expect(authProvider.refreshToken).toBe('mockUpdatedRefreshToken')
    })

    it('should get access token using client credentials if refresh token is not provided', async () => {
        const mockConfig = {
            tenantId: 'mockTenantId',
            authentication: {
                clientId: 'mockClientId',
                clientSecret: 'mockClientSecret',
            },
        }

        mockedAxios.post.mockResolvedValue({
            data: { access_token: 'mockAccessToken' },
        })

        const authProvider = new AuthProvider(mockConfig)
        const result = await authProvider.getAccessToken()

        expect(result).toBe('mockAccessToken')
        expect(mockedAxios.post).toHaveBeenCalledWith(
            `https://login.microsoftonline.com/${mockConfig.tenantId}/oauth2/v2.0/token`,
            qs.stringify({
                client_id: mockConfig.authentication.clientId,
                client_secret: mockConfig.authentication.clientSecret,
                grant_type: 'client_credentials',
                scope: 'https://graph.microsoft.com/.default',
            }),
            {
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            },
        )
        expect(authProvider.refreshToken).toBeUndefined()
    })
})
