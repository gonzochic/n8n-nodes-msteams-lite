import { type INodeProperties, type IExecuteFunctions, type IDataObject, updateDisplayOptions } from 'n8n-workflow';

import { prepareMessage } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';
import { chatRLC, recipientRLC, senderRLC } from '../../descriptions/rlc.description';

const properties: INodeProperties[] = [
	{
		displayName: 'Send To',
		name: 'chatType',
		type: 'options',
		options: [
			{
				name: 'User (One-on-One)',
				value: 'oneOnOne',
				description: 'Send a direct message to a user by email address',
			},
			{
				name: 'Existing Chat',
				value: 'existingChat',
				description: 'Send a message to an existing chat (group or one-on-one)',
			},
		],
		default: 'oneOnOne',
		description: 'Whether to send to a user directly or to an existing chat',
	},
	{
		...senderRLC,
		displayOptions: {
			show: {
				chatType: ['oneOnOne'],
			},
		},
	},
	{
		...recipientRLC,
		displayOptions: {
			show: {
				chatType: ['oneOnOne'],
			},
		},
	},
	{
		...chatRLC,
		displayOptions: {
			show: {
				chatType: ['existingChat'],
			},
		},
	},
	{
		displayName: 'Content Type',
		name: 'contentType',
		required: true,
		type: 'options',
		options: [
			{
				name: 'Text',
				value: 'text',
			},
			{
				name: 'HTML',
				value: 'html',
			},
		],
		default: 'text',
		description: 'Whether the message is plain text or HTML',
	},
	{
		displayName: 'Message',
		name: 'message',
		required: true,
		type: 'string',
		default: '',
		description: 'The content of the message to be sent',
		typeOptions: {
			rows: 2,
		},
	},
];

const displayOptions = {
	show: {
		resource: ['chatMessage'],
		operation: ['create'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(this: IExecuteFunctions, i: number) {
	const chatType = this.getNodeParameter('chatType', i, 'oneOnOne') as string;
	const contentType = this.getNodeParameter('contentType', i) as string;
	const message = this.getNodeParameter('message', i) as string;

	let chatId: string;

	if (chatType === 'existingChat') {
		chatId = this.getNodeParameter('chatId', i, '', { extractValue: true }) as string;
	} else {
		const sender = this.getNodeParameter('sender', i, '', { extractValue: true }) as string;
		const recipient = this.getNodeParameter('recipient', i, '', { extractValue: true }) as string;

		// get or create chat between sender and recipient
		const getOrCreateChatResponse = await microsoftApiRequest.call(
			this,
			'POST',
			'/v1.0/chats',
			{
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
			},
		);

		chatId = getOrCreateChatResponse.id;
	}

	const body: IDataObject = prepareMessage.call(
		this,
		message,
		contentType,
	);

	return await microsoftApiRequest.call(this, 'POST', `/v1.0/chats/${chatId}/messages`, body);
}
