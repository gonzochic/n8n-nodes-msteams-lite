import {
	type INodeProperties,
	type IExecuteFunctions,
	type IDataObject,
	updateDisplayOptions,
} from 'n8n-workflow';

import { teamRLC, channelRLC, userRLC } from '../../descriptions/rlc.description';
import { microsoftApiRequest } from '../../transport';

const properties: INodeProperties[] = [
	teamRLC,
	channelRLC,
	userRLC,
	{
		displayName: 'Role',
		name: 'role',
		type: 'options',
		options: [
			{
				name: 'Member',
				value: 'member',
				description: 'A regular channel member',
			},
			{
				name: 'Owner',
				value: 'owner',
				description: 'A channel owner with full permissions',
			},
		],
		default: 'member',
		description: 'The role to assign to the user in the channel',
	},
];

const displayOptions = {
	show: {
		resource: ['channelMember'],
		operation: ['add'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(this: IExecuteFunctions, i: number) {
	const teamId = this.getNodeParameter('teamId', i, '', { extractValue: true }) as string;
	const channelId = this.getNodeParameter('channelId', i, '', { extractValue: true }) as string;
	const userId = this.getNodeParameter('userId', i, '', { extractValue: true }) as string;
	const role = this.getNodeParameter('role', i) as string;

	const body: IDataObject = {
		'@odata.type': '#microsoft.graph.aadUserConversationMember',
		roles: role === 'owner' ? ['owner'] : [],
		'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${userId}')`,
	};

	return await microsoftApiRequest.call(
		this,
		'POST',
		`/v1.0/teams/${teamId}/channels/${channelId}/members`,
		body,
	);
}
