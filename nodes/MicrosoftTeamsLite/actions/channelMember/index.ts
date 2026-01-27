import type { INodeProperties } from 'n8n-workflow';

import * as add from './add.operation';
import * as remove from './remove.operation';

export { add, remove };

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['channelMember'],
			},
		},
		options: [
			{
				name: 'Add',
				value: 'add',
				description: 'Add a member to a channel',
				action: 'Add channel member',
			},
			{
				name: 'Remove',
				value: 'remove',
				description: 'Remove a member from a channel',
				action: 'Remove channel member',
			},
		],
		default: 'add',
	},

	...add.description,
	...remove.description,
];
