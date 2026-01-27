import {
	type INodeProperties,
	type IExecuteFunctions,
	type IDataObject,
	updateDisplayOptions,
} from 'n8n-workflow';

import { teamRLC, userRLC } from '../../descriptions/rlc.description';
import { microsoftApiRequest } from '../../transport';

const properties: INodeProperties[] = [
	teamRLC,
	userRLC,
];

const displayOptions = {
	show: {
		resource: ['teamMember'],
		operation: ['add'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(this: IExecuteFunctions, i: number) {
	const teamId = this.getNodeParameter('teamId', i, '', { extractValue: true }) as string;
	const userId = this.getNodeParameter('userId', i, '', { extractValue: true }) as string;

	const body: IDataObject = {
		'@odata.type': '#microsoft.graph.aadUserConversationMember',
		roles: [],
		'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${userId}')`,
	};

	return await microsoftApiRequest.call(this, 'POST', `/v1.0/teams/${teamId}/members`, body);
}
