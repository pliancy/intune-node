import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceManagementIntent } from '../types'
import { mockClient } from '../../../__fixtures__/@microsoft/microsoft-graph-client'
import { DeviceManagementIntents } from './deviceManagementIntents'

describe('DeviceManagementIntents', () => {
    let graphClient: Client
    let deviceManagementIntents: DeviceManagementIntents
    const intent = {
        displayName: 'test',
        id: 'test',
    } as DeviceManagementIntent

    beforeEach(() => {
        graphClient = mockClient() as never as Client
        deviceManagementIntents = new DeviceManagementIntents(graphClient)
    })

    it('should get a intent', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(intent)
        const result = await deviceManagementIntents.get('')
        expect(result).toEqual(intent)
    })

    it('should list all intents', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [intent],
            '@odata.nextLink': 'next',
        })
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [intent],
        })

        const result = await deviceManagementIntents.list()
        expect(result).toEqual([intent, intent])
    })

    it('should update a intent', async () => {
        const postSpy = jest.spyOn(graphClient.api(''), 'patch').mockResolvedValue(intent)
        const result = await deviceManagementIntents.update('id', intent)
        expect(result).toEqual(intent)
        expect(postSpy).toHaveBeenCalledWith(intent)
    })

    it('should delete a intent', async () => {
        const spy = jest.spyOn(graphClient, 'api')
        jest.spyOn(graphClient.api(''), 'patch')
        const result = await deviceManagementIntents.delete('id')
        expect(result).toBeUndefined()
        expect(spy).toHaveBeenCalledWith('/deviceManagement/intents/id')
    })

    it('should create a intent', async () => {
        const apiSpy = jest.spyOn(graphClient, 'api')
        const postSpy = jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(intent)
        const result = await deviceManagementIntents.create(intent)
        expect(result).toEqual(intent)
        expect(apiSpy).toHaveBeenCalledWith('/deviceManagement/intents')
        expect(postSpy).toHaveBeenCalledWith(intent)
    })

    describe('setAssignments', () => {
        it('should assign to all devices with exclusions', async () => {
            const spy = jest.spyOn(graphClient.api(''), 'post')

            await deviceManagementIntents.setAssignments('id', {
                allDevices: true,
                excludeGroups: ['group1'],
            })

            expect(spy).toHaveBeenCalledWith({
                assignments: [
                    {
                        target: {
                            '@odata.type': '#microsoft.graph.allDevicesAssignmentTarget',
                        },
                    },
                    {
                        target: {
                            '@odata.type': '#microsoft.graph.exclusionGroupAssignmentTarget',
                            groupId: 'group1',
                        },
                    },
                ],
            })
        })

        it('should assign to all licensed users', async () => {
            const spy = jest.spyOn(graphClient.api(''), 'post')

            await deviceManagementIntents.setAssignments('id', {
                allUsers: true,
            })

            expect(spy).toHaveBeenCalledWith({
                assignments: [
                    {
                        target: {
                            '@odata.type': '#microsoft.graph.allLicensedUsersAssignmentTarget',
                        },
                    },
                ],
            })
        })

        it('should support combination of all devices and all users with exclusions', async () => {
            const spy = jest.spyOn(graphClient.api(''), 'post')

            await deviceManagementIntents.setAssignments('id', {
                allDevices: true,
                allUsers: true,
                excludeGroups: ['group1'],
            })

            expect(spy).toHaveBeenCalledWith({
                assignments: [
                    {
                        target: {
                            '@odata.type': '#microsoft.graph.allDevicesAssignmentTarget',
                        },
                    },
                    {
                        target: {
                            '@odata.type': '#microsoft.graph.allLicensedUsersAssignmentTarget',
                        },
                    },
                    {
                        target: {
                            '@odata.type': '#microsoft.graph.exclusionGroupAssignmentTarget',
                            groupId: 'group1',
                        },
                    },
                ],
            })
        })

        it('should assign to specific included groups', async () => {
            const spy = jest.spyOn(graphClient.api(''), 'post')

            await deviceManagementIntents.setAssignments('id', {
                includeGroups: ['group1', 'group2'],
            })

            expect(spy).toHaveBeenCalledWith({
                assignments: [
                    {
                        target: {
                            '@odata.type': '#microsoft.graph.groupAssignmentTarget',
                            groupId: 'group1',
                        },
                    },
                    {
                        target: {
                            '@odata.type': '#microsoft.graph.groupAssignmentTarget',
                            groupId: 'group2',
                        },
                    },
                ],
            })
        })

        it('should throw error when including groups with allDevices', async () => {
            await expect(
                deviceManagementIntents.setAssignments('id', {
                    allDevices: true,
                    includeGroups: ['group1'],
                }),
            ).rejects.toThrow('Cannot include specific groups when allDevices is true')
        })

        it('should support mix of include and exclude groups', async () => {
            const spy = jest.spyOn(graphClient.api(''), 'post')

            await deviceManagementIntents.setAssignments('id', {
                includeGroups: ['group1'],
                excludeGroups: ['group2'],
            })

            expect(spy).toHaveBeenCalledWith({
                assignments: [
                    {
                        target: {
                            '@odata.type': '#microsoft.graph.groupAssignmentTarget',
                            groupId: 'group1',
                        },
                    },
                    {
                        target: {
                            '@odata.type': '#microsoft.graph.exclusionGroupAssignmentTarget',
                            groupId: 'group2',
                        },
                    },
                ],
            })
        })
    })
})
