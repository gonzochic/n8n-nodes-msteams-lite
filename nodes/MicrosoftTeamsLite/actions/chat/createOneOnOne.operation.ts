import { type INodeProperties, type IExecuteFunctions, updateDisplayOptions } from 'n8n-workflow';

import { microsoftApiRequest } from '../../transport';
import { recipientRLC, senderRLC } from '../../descriptions/rlc.description';

const properties: INodeProperties[] = [
	senderRLC,
	recipientRLC,
];

const displayOptions = {
	show: {
		resource: ['chat'],
		operation: ['createOneOnOne'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(this: IExecuteFunctions, i: number) {
	const sender = this.getNodeParameter('sender', i, '', { extractValue: true }) as string;
	const recipient = this.getNodeParameter('recipient', i, '', { extractValue: true }) as string;

	return await microsoftApiRequest.call(this, 'POST', '/v1.0/chats', {
		chatType: 'oneOnOne',
		members: [
			{
				'@odata.type': '#microsoft.graph.aadUserConversationMember',
				roles: ['owner'],
				"user@odata.bind": "https://graph.microsoft.com/v1.0/users('" + sender + "')",
			},
			{
				'@odata.type': '#microsoft.graph.aadUserConversationMember',
				roles: ['owner'],
				"user@odata.bind": "https://graph.microsoft.com/v1.0/users('" + recipient + "')",
			},
		],
	});
}
