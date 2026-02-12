/* eslint-disable n8n-nodes-base/node-filename-against-convention */
import { NodeConnectionTypes, type INodeTypeDescription } from 'n8n-workflow';

import * as chat from './chat';
import * as chatMessage from './chatMessage';
import * as channel from './channel';
import * as channelMessage from './channelMessage';
import * as teamMember from './teamMember';


export const versionDescription: INodeTypeDescription = {
	displayName: 'Microsoft Teams (Lite)',
	name: 'microsoftTeamsLite',
	icon: 'file:teams.svg',
	group: ['input'],
	version: 2,
	subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
	description: 'Consume Microsoft Teams API in a admin friendly way (Lite)',
	defaults: { name: 'Microsoft Teams Lite' },
	inputs: [NodeConnectionTypes.Main],
	outputs: [NodeConnectionTypes.Main],
	credentials: [
		{
			name: 'microsoftTeamsLiteOAuth2Api',
			required: true,
		},
	],
	usableAsTool: true,
	properties: [
		{
			displayName: 'Resource',
			name: 'resource',
			type: 'options',
			noDataExpression: true,
			options: [
				{
					name: 'Channel',
					value: 'channel',
				},
{
					name: 'Channel Message',
					value: 'channelMessage',
				},
				{
					name: 'Chat',
					value: 'chat',
				},
				{
					name: 'Chat Message',
					value: 'chatMessage',
				},
				{
					name: 'Team Member',
					value: 'teamMember',
				},
			],
			default: 'chatMessage',
		},

		...chat.description,
		...chatMessage.description,
		...channel.description,
		...channelMessage.description,
		...teamMember.description,

	],
};
