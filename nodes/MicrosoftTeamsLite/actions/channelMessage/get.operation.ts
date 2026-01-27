import {
	type INodeProperties,
	type IExecuteFunctions,
	NodeOperationError,
	updateDisplayOptions,
} from 'n8n-workflow';

import { teamRLC, channelRLC } from '../../descriptions/rlc.description';
import { microsoftApiRequest } from '../../transport';

const properties: INodeProperties[] = [
	teamRLC,
	channelRLC,
	{
		displayName: 'Message ID',
		name: 'messageId',
		required: true,
		type: 'string',
		default: '',
		placeholder: 'e.g. 1673355049064',
		description: 'The ID of the message to retrieve',
	},
];

const displayOptions = {
	show: {
		resource: ['channelMessage'],
		operation: ['get'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(this: IExecuteFunctions, i: number) {
	try {
		const teamId = this.getNodeParameter('teamId', i, '', { extractValue: true }) as string;
		const channelId = this.getNodeParameter('channelId', i, '', { extractValue: true }) as string;
		const messageId = this.getNodeParameter('messageId', i) as string;

		return await microsoftApiRequest.call(
			this,
			'GET',
			`/v1.0/teams/${teamId}/channels/${channelId}/messages/${messageId}`,
		);
	// eslint-disable-next-line @typescript-eslint/no-unused-vars
	} catch (error) {
		throw new NodeOperationError(
			this.getNode(),
			"The message you are trying to get doesn't exist",
			{
				description: "Check that the 'Message ID' parameter is correctly set",
			},
		);
	}
}
