import type { INodeProperties } from 'n8n-workflow';

import * as create from './create.operation';
import * as get from './get.operation';
import * as getAll from './getAll.operation';

export { create, get, getAll };

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['channelMessage'],
			},
		},
		options: [
			{
				name: 'Create',
				value: 'create',
				description: 'Send a message to a channel',
				action: 'Send channel message',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get a message from a channel',
				action: 'Get channel message',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many messages from a channel',
				action: 'Get many channel messages',
			},
		],
		default: 'create',
	},

	...create.description,
	...get.description,
	...getAll.description,
];
