import { Client } from '@microsoft/microsoft-graph-client'
import { DeviceConfigurationPolicies } from './deviceConfigurationPolicies'
import { mockClient } from '../../../__fixtures__/@microsoft/microsoft-graph-client'

describe('Device Configuration Policies', () => {
    let graphClient: Client
    let configurationPolicies: DeviceConfigurationPolicies

    const configurationPolicy = {
        name: 'test policy',
        '@odata.type': '#microsoft.graph.deviceManagementConfigurationPolicy',
    }

    const policyAssignment = {
        '@odata.type': '#microsoft.graph.deviceManagementConfigurationPolicyAssignment',
        target: {
            '@odata.type': '#microsoft.graph.groupAssignmentTarget',
            groupId: '1',
        },
    }

    beforeEach(() => {
        graphClient = mockClient() as never as Client
        configurationPolicies = new DeviceConfigurationPolicies(graphClient)
    })

    it('should list all configuration policies', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [configurationPolicy],
            '@odata.nextLink': 'next',
        })
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [configurationPolicy],
        })

        const result = await configurationPolicies.list()
        expect(result).toEqual([configurationPolicy, configurationPolicy])
    })

    it('should get a configuration policy', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(configurationPolicy)
        const result = await configurationPolicies.get('id')
        expect(result).toEqual(configurationPolicy)
    })

    it('should create a configuration policy', async () => {
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(configurationPolicy)
        const result = await configurationPolicies.create(configurationPolicy)
        expect(result).toEqual(configurationPolicy)
    })

    it('should update a configuration policy', async () => {
        jest.spyOn(graphClient.api(''), 'patch')
        const result = await configurationPolicies.update('id', configurationPolicy)
        expect(result).toBeUndefined()
    })

    it('should delete a configuration policy', async () => {
        jest.spyOn(graphClient.api(''), 'delete')
        const result = await configurationPolicies.delete('id')
        expect(result).toBeUndefined()
    })

    describe('setAssignments', () => {
        it('should assign to all devices with exclusions', async () => {
            const spy = jest.spyOn(graphClient.api(''), 'post')

            await configurationPolicies.setAssignments('id', {
                allDevices: true,
                excludeGroups: ['group1', 'group2'],
            })

            expect(spy).toHaveBeenCalledWith({
                assignments: [
                    {
                        id: '',
                        source: 'direct',
                        target: {
                            '@odata.type': '#microsoft.graph.allDevicesAssignmentTarget',
                            deviceAndAppManagementAssignmentFilterType: 'none',
                        },
                    },
                    {
                        id: '',
                        source: 'direct',
                        target: {
                            '@odata.type': '#microsoft.graph.exclusionGroupAssignmentTarget',
                            deviceAndAppManagementAssignmentFilterType: 'none',
                            groupId: 'group1',
                        },
                    },
                    {
                        id: '',
                        source: 'direct',
                        target: {
                            '@odata.type': '#microsoft.graph.exclusionGroupAssignmentTarget',
                            deviceAndAppManagementAssignmentFilterType: 'none',
                            groupId: 'group2',
                        },
                    },
                ],
            })
        })

        it('should assign to all licensed users', async () => {
            const spy = jest.spyOn(graphClient.api(''), 'post')

            await configurationPolicies.setAssignments('id', {
                allUsers: true,
            })

            expect(spy).toHaveBeenCalledWith({
                assignments: [
                    {
                        id: '',
                        source: 'direct',
                        target: {
                            '@odata.type': '#microsoft.graph.allLicensedUsersAssignmentTarget',
                            deviceAndAppManagementAssignmentFilterType: 'none',
                        },
                    },
                ],
            })
        })

        it('should assign to specific included groups', async () => {
            const spy = jest.spyOn(graphClient.api(''), 'post')

            await configurationPolicies.setAssignments('id', {
                includeGroups: ['group1', 'group2'],
            })

            expect(spy).toHaveBeenCalledWith({
                assignments: [
                    {
                        id: '',
                        source: 'direct',
                        target: {
                            '@odata.type': '#microsoft.graph.groupAssignmentTarget',
                            deviceAndAppManagementAssignmentFilterType: 'none',
                            groupId: 'group1',
                        },
                    },
                    {
                        id: '',
                        source: 'direct',
                        target: {
                            '@odata.type': '#microsoft.graph.groupAssignmentTarget',
                            deviceAndAppManagementAssignmentFilterType: 'none',
                            groupId: 'group2',
                        },
                    },
                ],
            })
        })

        it('should throw error when including groups with allDevices', async () => {
            await expect(
                configurationPolicies.setAssignments('id', {
                    allDevices: true,
                    includeGroups: ['group1'],
                }),
            ).rejects.toThrow('Cannot include specific groups when allDevices is true')
        })

        it('should support mix of include and exclude groups', async () => {
            const spy = jest.spyOn(graphClient.api(''), 'post')

            await configurationPolicies.setAssignments('id', {
                includeGroups: ['group1'],
                excludeGroups: ['group2'],
            })

            expect(spy).toHaveBeenCalledWith({
                assignments: [
                    {
                        id: '',
                        source: 'direct',
                        target: {
                            '@odata.type': '#microsoft.graph.groupAssignmentTarget',
                            deviceAndAppManagementAssignmentFilterType: 'none',
                            groupId: 'group1',
                        },
                    },
                    {
                        id: '',
                        source: 'direct',
                        target: {
                            '@odata.type': '#microsoft.graph.exclusionGroupAssignmentTarget',
                            deviceAndAppManagementAssignmentFilterType: 'none',
                            groupId: 'group2',
                        },
                    },
                ],
            })
        })
    })

    describe('pagination', () => {
        it('should handle pagination for list method', async () => {
            const firstPage = {
                value: [{ ...configurationPolicy, id: '1' }],
                '@odata.nextLink': 'https://graph.microsoft.com/beta/next-page',
            }
            const secondPage = {
                value: [{ ...configurationPolicy, id: '2' }],
            }

            jest.spyOn(graphClient.api(''), 'get')
                .mockResolvedValueOnce(firstPage)
                .mockResolvedValueOnce(secondPage)

            const result = await configurationPolicies.list()

            expect(result).toHaveLength(2)
            expect(result[0].id).toBe('1')
            expect(result[1].id).toBe('2')
        })
    })
})
