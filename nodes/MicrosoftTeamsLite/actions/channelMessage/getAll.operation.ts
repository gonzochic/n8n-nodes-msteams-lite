import { type INodeProperties, type IExecuteFunctions, updateDisplayOptions } from 'n8n-workflow';

import { teamRLC, channelRLC } from '../../descriptions/rlc.description';
import { microsoftApiRequestAllItems } from '../../transport';

const properties: INodeProperties[] = [
	teamRLC,
	channelRLC,
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		default: false,
		description: 'Whether to return all results or only up to a given limit',
	},
	{
		displayName: 'Limit',
		name: 'limit',
		type: 'number',
		default: 50,
		description: 'Max number of results to return',
		typeOptions: {
			minValue: 1,
		},
		displayOptions: {
			show: {
				returnAll: [false],
			},
		},
	},
];

const displayOptions = {
	show: {
		resource: ['channelMessage'],
		operation: ['getAll'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(this: IExecuteFunctions, i: number) {
	const teamId = this.getNodeParameter('teamId', i, '', { extractValue: true }) as string;
	const channelId = this.getNodeParameter('channelId', i, '', { extractValue: true }) as string;
	const returnAll = this.getNodeParameter('returnAll', i);

	if (returnAll) {
		return await microsoftApiRequestAllItems.call(
			this,
			'value',
			'GET',
			`/v1.0/teams/${teamId}/channels/${channelId}/messages`,
		);
	} else {
		const limit = this.getNodeParameter('limit', i);
		const responseData = await microsoftApiRequestAllItems.call(
			this,
			'value',
			'GET',
			`/v1.0/teams/${teamId}/channels/${channelId}/messages`,
			{},
		);
		return responseData.splice(0, limit);
	}
}
