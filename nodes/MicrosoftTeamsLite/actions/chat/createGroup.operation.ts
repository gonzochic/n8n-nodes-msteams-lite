import { type INodeProperties, type IExecuteFunctions, updateDisplayOptions } from 'n8n-workflow';

import { microsoftApiRequest } from '../../transport';
import { senderRLC } from '../../descriptions/rlc.description';

const properties: INodeProperties[] = [
	senderRLC,
	{
		displayName: 'Chat Name',
		name: 'chatName',
		type: 'string',
		required: true,
		default: '',
		placeholder: 'e.g. Project Discussion',
		description: 'The name (topic) of the group chat',
	},
	{
		displayName: 'Members',
		name: 'members',
		type: 'string',
		required: true,
		default: '',
		placeholder: 'e.g. user1@company.com, user2@company.com',
		description: 'Comma-separated email addresses of the members to add to the group chat',
	},
];

const displayOptions = {
	show: {
		resource: ['chat'],
		operation: ['createGroup'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(this: IExecuteFunctions, i: number) {
	const sender = this.getNodeParameter('sender', i, '', { extractValue: true }) as string;
	const chatName = this.getNodeParameter('chatName', i) as string;
	const membersRaw = this.getNodeParameter('members', i) as string;

	const memberEmails = membersRaw
		.split(',')
		.map((email) => email.trim())
		.filter((email) => email.length > 0);

	const members = [
		{
			'@odata.type': '#microsoft.graph.aadUserConversationMember',
			roles: ['owner'],
			"user@odata.bind": "https://graph.microsoft.com/v1.0/users('" + sender + "')",
		},
		...memberEmails.map((email) => ({
			'@odata.type': '#microsoft.graph.aadUserConversationMember',
			roles: ['guest'],
			"user@odata.bind": "https://graph.microsoft.com/v1.0/users('" + email + "')",
		})),
	];

	return await microsoftApiRequest.call(this, 'POST', '/v1.0/chats', {
		chatType: 'group',
		topic: chatName,
		members,
	});
}
