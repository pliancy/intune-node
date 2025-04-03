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

    describe('when managing assignments', () => {
        it('should assign a script to specified targets', async () => {
            const apiSpy = jest.spyOn(graphClient, 'api')
            const postSpy = jest.spyOn(graphClient.api(''), 'post')

            const assignmentPayload = {
                deviceManagementScriptAssignments: [
                    {
                        '@odata.type': '#microsoft.graph.deviceShellScriptAssignment',
                        target: {
                            '@odata.type': '#microsoft.graph.allDevicesAssignmentTarget',
                        },
                    },
                ],
            }

            await deviceShellScripts.assign('123', assignmentPayload as never)

            expect(apiSpy).toHaveBeenCalledWith('/deviceManagement/deviceShellScripts/123/assign')
            expect(postSpy).toHaveBeenCalledWith(assignmentPayload)
        })

        it('should assign a script to a specific group', async () => {
            const assignSpy = jest.spyOn(deviceShellScripts, 'assign').mockResolvedValue()

            await deviceShellScripts.assignToGroup('123', 'group456')

            expect(assignSpy).toHaveBeenCalledWith('123', {
                deviceManagementScriptGroupAssignments: [
                    {
                        '@odata.type': '#microsoft.graph.deviceShellScriptGroupAssignment',
                        targetGroupId: 'group456',
                    },
                ],
            })
        })

        it('should assign a script to all devices', async () => {
            const assignSpy = jest.spyOn(deviceShellScripts, 'assign').mockResolvedValue()

            await deviceShellScripts.assignToAllDevices('123')

            expect(assignSpy).toHaveBeenCalledWith('123', {
                deviceManagementScriptAssignments: [
                    expect.objectContaining({
                        '@odata.type': '#microsoft.graph.deviceShellScriptAssignment',
                        target: {
                            '@odata.type': '#microsoft.graph.allDevicesAssignmentTarget',
                        },
                    }),
                ],
            })
        })

        it('should remove an assignment', async () => {
            // Mock script with assignments
            const scriptWithAssignments = {
                ...script,
                assignments: [
                    {
                        id: 'assign1',
                        target: {
                            '@odata.type': '#microsoft.graph.allDevicesAssignmentTarget',
                        },
                    },
                    {
                        id: 'assign2',
                        target: {
                            '@odata.type': '#microsoft.graph.groupAssignmentTarget',
                            groupId: 'group1',
                        },
                    },
                ],
            }

            jest.spyOn(deviceShellScripts, 'get').mockResolvedValue(scriptWithAssignments as never)
            const assignSpy = jest.spyOn(deviceShellScripts, 'assign').mockResolvedValue()

            await deviceShellScripts.removeAssignment('123', 'assign1')

            expect(assignSpy).toHaveBeenCalledWith('123', {
                deviceManagementScriptGroupAssignments: [
                    expect.objectContaining({
                        '@odata.type': '#microsoft.graph.deviceShellScriptGroupAssignment',
                        id: 'assign2',
                        targetGroupId: 'group1',
                    }),
                ],
            })
        })

        it('should do nothing when removing a non-existent assignment', async () => {
            // Mock script with assignments
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

            jest.spyOn(deviceShellScripts, 'get').mockResolvedValue(scriptWithAssignments as never)
            const assignSpy = jest.spyOn(deviceShellScripts, 'assign')

            await deviceShellScripts.removeAssignment('123', 'non-existent')

            expect(assignSpy).not.toHaveBeenCalled()
        })

        it('should do nothing when removing assignment from script with no assignments', async () => {
            jest.spyOn(deviceShellScripts, 'get').mockResolvedValue(script as never)
            const assignSpy = jest.spyOn(deviceShellScripts, 'assign')

            await deviceShellScripts.removeAssignment('123', 'assign1')

            expect(assignSpy).not.toHaveBeenCalled()
        })
    })
})
