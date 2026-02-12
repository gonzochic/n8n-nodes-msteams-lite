import type { INodeProperties } from 'n8n-workflow';

import * as createGroup from './createGroup.operation';
import * as createOneOnOne from './createOneOnOne.operation';

export { createGroup, createOneOnOne };

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['chat'],
			},
		},
		options: [
			{
				name: 'Create Group Chat',
				value: 'createGroup',
				description: 'Create a group chat with a name and multiple members',
				action: 'Create group chat',
			},
			{
				name: 'Create One-on-One Chat',
				value: 'createOneOnOne',
				description: 'Create a one-on-one chat with a single user',
				action: 'Create one on one chat',
			},
		],
		default: 'createGroup',
	},

	...createGroup.description,
	...createOneOnOne.description,
];
