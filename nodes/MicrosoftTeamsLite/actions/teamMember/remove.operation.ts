import {
	type INodeProperties,
	type IExecuteFunctions,
	NodeOperationError,
	updateDisplayOptions,
} from 'n8n-workflow';

import { teamRLC, teamMemberRLC } from '../../descriptions/rlc.description';
import { microsoftApiRequest } from '../../transport';

const properties: INodeProperties[] = [teamRLC, teamMemberRLC];

const displayOptions = {
	show: {
		resource: ['teamMember'],
		operation: ['remove'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(this: IExecuteFunctions, i: number) {
	try {
		const teamId = this.getNodeParameter('teamId', i, '', { extractValue: true }) as string;
		const memberId = this.getNodeParameter('memberId', i, '', { extractValue: true }) as string;

		await microsoftApiRequest.call(
			this,
			'DELETE',
			`/v1.0/teams/${teamId}/members/${memberId}`,
		);

		return { success: true };
	// eslint-disable-next-line @typescript-eslint/no-unused-vars
	} catch (error) {
		throw new NodeOperationError(
			this.getNode(),
			'Failed to remove member from team',
			{
				description: 'Check that the member ID is correct and you have permission to remove them',
			},
		);
	}
}
