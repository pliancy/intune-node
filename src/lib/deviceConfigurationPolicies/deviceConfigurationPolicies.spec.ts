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

    it('should assign policy to group', async () => {
        jest.spyOn(graphClient.api(''), 'post').mockResolvedValue(policyAssignment)
        const spy = jest.spyOn(graphClient, 'api')
        const result = await configurationPolicies.assignToGroup('id', 'groupId')
        expect(result).toEqual(policyAssignment)
        expect(spy).toHaveBeenCalledWith('/deviceManagement/configurationPolicies/id/assignments')
    })

    it('should list policy assignments', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [policyAssignment],
            '@odata.nextLink': 'next',
        })
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValueOnce({
            value: [policyAssignment],
        })

        const result = await configurationPolicies.listAssignments('id')
        expect(result).toEqual([policyAssignment, policyAssignment])
    })

    it('should get a policy assignment', async () => {
        jest.spyOn(graphClient.api(''), 'get').mockResolvedValue(policyAssignment)
        const result = await configurationPolicies.getAssignment('id', 'assignmentId')
        expect(result).toEqual(policyAssignment)
    })

    it('should delete a policy assignment', async () => {
        jest.spyOn(graphClient.api(''), 'delete')
        const result = await configurationPolicies.deleteAssignment('id', 'assignmentId')
        expect(result).toBeUndefined()
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

        it('should handle pagination for listAssignments method', async () => {
            const firstPage = {
                value: [{ ...policyAssignment, id: '1' }],
                '@odata.nextLink': 'https://graph.microsoft.com/beta/next-page',
            }
            const secondPage = {
                value: [{ ...policyAssignment, id: '2' }],
            }

            jest.spyOn(graphClient.api(''), 'get')
                .mockResolvedValueOnce(firstPage)
                .mockResolvedValueOnce(secondPage)

            const result = await configurationPolicies.listAssignments('id')

            expect(result).toHaveLength(2)
            expect(result[0].id).toBe('1')
            expect(result[1].id).toBe('2')
        })
    })
})
