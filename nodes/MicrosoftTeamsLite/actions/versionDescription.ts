/* eslint-disable n8n-nodes-base/node-filename-against-convention */
import { NodeConnectionTypes, type INodeTypeDescription } from 'n8n-workflow';

import * as chatMessage from './chatMessage';

import { sendAndWaitWebhooksDescription } from 'n8n-nodes-base/dist/utils/sendAndWait/descriptions';
import { SEND_AND_WAIT_WAITING_TOOLTIP } from 'n8n-nodes-base/dist/utils/sendAndWait/utils';

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
			name: 'microsoftOAuth2Api',
			required: true,
		},
	],
	waitingNodeTooltip: SEND_AND_WAIT_WAITING_TOOLTIP,
	webhooks: sendAndWaitWebhooksDescription,
	usableAsTool: true,
	properties: [
		{
			displayName: 'Resource',
			name: 'resource',
			type: 'options',
			noDataExpression: true,
			options: [
				{
					name: 'Chat Message',
					value: 'chatMessage',
				},
			],
			default: 'chatMessage',
		},

		...chatMessage.description,
	],
};
