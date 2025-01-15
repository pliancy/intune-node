import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceConfigurations } from './deviceConfigurations'
import { mockClient } from '../../../__fixtures__/@microsoft/microsoft-graph-client'

describe('Device Configurations', () => {
    let graphClient: Client
    let deviceConfigurations: DeviceConfigurations
    const deviceConfiguration = {
        name: 'test',
        '@odata.type': '#microsoft.graph.deviceConfiguration',
    }

    const groupAssignment = {
        '@odata.type': '#microsoft.graph.deviceManagementScriptGroupAssignment',
        targetGroupId: '1',
    }

    beforeEach(() => {
        graphClient = mockClient() as never as Client
        deviceConfigurations = new DeviceConfigurations(graphClient)
    })

    it('should get a device configuration', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(deviceConfiguration)
        const result = await deviceConfigurations.get('')
        expect(result).toEqual(deviceConfiguration)
    })

    it('should list all device configurations', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [deviceConfiguration],
            '@odata.nextLink': 'next',
        })
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [deviceConfiguration],
        })

        const result = await deviceConfigurations.list()
        expect(result).toEqual([deviceConfiguration, deviceConfiguration])
    })

    it('should create a device configuration', async () => {
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(deviceConfiguration)
        const result = await deviceConfigurations.create(deviceConfiguration)
        expect(result).toEqual(deviceConfiguration)
    })

    it('should update a device configuration', async () => {
        jest.spyOn(graphClient.api(''), 'patch')
        const result = await deviceConfigurations.update('id', deviceConfiguration)
        expect(result).toBeUndefined()
    })

    it('should delete a device configuration', async () => {
        jest.spyOn(graphClient.api(''), 'delete')
        const result = await deviceConfigurations.delete('id')
        expect(result).toBeUndefined()
    })

    it('should create a group assignment', async () => {
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(groupAssignment)
        const spy = jest.spyOn(graphClient, 'api')
        const result = await deviceConfigurations.createGroupAssignment('id', 'groupId')
        expect(result).toEqual(groupAssignment)
        expect(spy).toHaveBeenCalledWith(
            '/deviceManagement/deviceConfigurations/id/groupAssignments',
        )
    })

    it('should list group assignments', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue({ value: [groupAssignment] })
        const result = await deviceConfigurations.listGroupAssignments('id')
        expect(result).toEqual([groupAssignment])
    })

    it('should get a group assignment', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(groupAssignment)
        const result = await deviceConfigurations.getGroupAssignment('id', 'groupAssignmentId')
        expect(result).toEqual(groupAssignment)
    })

    it('should delete a group assignment', async () => {
        jest.spyOn(graphClient.api(''), 'delete')
        const result = await deviceConfigurations.deleteGroupAssignment('id', 'groupId')
        expect(result).toBeUndefined()
    })

    describe('setAssignments', () => {
        it('should assign to all devices with exclusions', async () => {
            const spy = jest.spyOn(graphClient.api(''), 'post')

            await deviceConfigurations.setAssignments('id', {
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

            await deviceConfigurations.setAssignments('id', {
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

            await deviceConfigurations.setAssignments('id', {
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

            await deviceConfigurations.setAssignments('id', {
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
                deviceConfigurations.setAssignments('id', {
                    allDevices: true,
                    includeGroups: ['group1'],
                }),
            ).rejects.toThrow('Cannot include specific groups when allDevices is true')
        })

        it('should support mix of include and exclude groups', async () => {
            const spy = jest.spyOn(graphClient.api(''), 'post')

            await deviceConfigurations.setAssignments('id', {
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
