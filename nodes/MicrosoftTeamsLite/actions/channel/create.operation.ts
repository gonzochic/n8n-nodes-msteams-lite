import {
	type INodeProperties,
	type IExecuteFunctions,
	type IDataObject,
	updateDisplayOptions,
} from 'n8n-workflow';

import { teamRLC } from '../../descriptions/rlc.description';
import { microsoftApiRequest } from '../../transport';

const properties: INodeProperties[] = [
	teamRLC,
	{
		displayName: 'Display Name',
		name: 'displayName',
		type: 'string',
		required: true,
		default: '',
		placeholder: 'e.g. Project Alpha',
		description: 'The name of the channel',
	},
	{
		displayName: 'Description',
		name: 'description',
		type: 'string',
		default: '',
		placeholder: 'e.g. Channel for Project Alpha discussions',
		description: 'A description of the channel',
	},
	{
		displayName: 'Membership Type',
		name: 'membershipType',
		type: 'options',
		options: [
			{
				name: 'Standard',
				value: 'standard',
				description: 'A standard channel that all team members can access',
			},
			{
				name: 'Private',
				value: 'private',
				description: 'A private channel with restricted membership',
			},
			{
				name: 'Shared',
				value: 'shared',
				description: 'A shared channel that can be accessed by users outside the team',
			},
		],
		default: 'standard',
		description: 'The type of channel to create',
	},
];

const displayOptions = {
	show: {
		resource: ['channel'],
		operation: ['create'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(this: IExecuteFunctions, i: number) {
	const teamId = this.getNodeParameter('teamId', i, '', { extractValue: true }) as string;
	const displayName = this.getNodeParameter('displayName', i) as string;
	const channelDescription = this.getNodeParameter('description', i) as string;
	const membershipType = this.getNodeParameter('membershipType', i) as string;

	const body: IDataObject = {
		displayName,
		membershipType,
	};

	if (channelDescription) {
		body.description = channelDescription;
	}

	return await microsoftApiRequest.call(this, 'POST', `/v1.0/teams/${teamId}/channels`, body);
}
