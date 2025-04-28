import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceShellScripts } from './deviceShellScripts'
import { mockClient } from '../../../__fixtures__/@microsoft/microsoft-graph-client'
import { DeviceShellScript } from '../types'

describe('DeviceShellScripts', () => {
    let graphClient: Client
    let deviceShellScripts: DeviceShellScripts
    const script: DeviceShellScript = {
        id: '123',
        '@odata.type': '#microsoft.graph.deviceShellScript',
        displayName: 'Test Script',
        description: 'Test Description',
        scriptContent: 'c2NyaXB0Q29udGVudA==',
        runAsAccount: 'user' as any,
        fileName: 'test.sh',
    }

    beforeEach(() => {
        graphClient = mockClient() as never as Client
        deviceShellScripts = new DeviceShellScripts(graphClient)
    })

    describe('when finding scripts', () => {
        it('should list all scripts', async () => {
            jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
                value: [script],
                '@odata.nextLink': 'next',
            })

            jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
                value: [script],
            })

            const result = await deviceShellScripts.list()
            expect(result).toEqual([script, script])
        })

        it('should get one script', async () => {
            jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(script)
            const result = await deviceShellScripts.get('123')
            expect(result).toEqual(script)
        })

        it('should get one script with assignments', async () => {
            const scriptWithAssignments = {
                ...script,
                assignments: [
                    {
                        id: 'assign1',
                        target: {
                            '@odata.type': '#microsoft.graph.allDevicesAssignmentTarget',
                        },
                    },
                ],
            }

            const expandSpy = jest.fn().mockReturnValue({
                get: jest.fn().mockResolvedValue(scriptWithAssignments),
            })

            jest.spyOn(graphClient, 'api').mockReturnValue({
                expand: expandSpy,
                get: jest.fn(),
            } as any)

            const result = await deviceShellScripts.get('123', true)
            expect(result).toEqual(scriptWithAssignments)
            expect(expandSpy).toHaveBeenCalledWith('assignments')
        })
    })

    describe('when managing scripts', () => {
        it('should create a script', async () => {
            jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(script)
            const result = await deviceShellScripts.create(script)

            expect(result).toEqual(script)
        })

        it('should update a script', async () => {
            const updatedScript = { ...script, displayName: 'Updated Script' }
            jest.spyOn(graphClient.api(''), 'patch').mockResolvedValue(updatedScript)

            const result = await deviceShellScripts.update('123', updatedScript)
            expect(result).toEqual(updatedScript)
        })

        it('should delete a script', async () => {
            const deleteSpy = jest.spyOn(graphClient.api(''), 'delete')
            await deviceShellScripts.delete('123')

            expect(deleteSpy).toHaveBeenCalled()
        })
    })

   
})
