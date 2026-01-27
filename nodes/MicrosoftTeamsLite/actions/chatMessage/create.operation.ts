import { type INodeProperties, type IExecuteFunctions, type IDataObject, updateDisplayOptions } from 'n8n-workflow';

import { prepareMessage } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';
import { senderRLC } from '../../descriptions/rlc.description';

const properties: INodeProperties[] = [
	senderRLC,
	{
		displayName: "Recipient Chat",
		name: "recipient",
		type: "string",
		required: true,
		default: "",
		placeholder: "e.g. user@company.com",
		description: "The (entra) email address of the message recipient",
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
	const sender = this.getNodeParameter('sender', i, '', { extractValue: true }) as string;
	const recipient = this.getNodeParameter('recipient', i, '', { extractValue: true }) as string;
	const contentType = this.getNodeParameter('contentType', i) as string;
	const message = this.getNodeParameter('message', i) as string;


	// get or create chat between sender and recipient
	const getOrCreateChatResponse = await microsoftApiRequest.call(
		this,
		'POST',
		`/v1.0/chats`,
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

	const chatId = getOrCreateChatResponse.id;

	const body: IDataObject = prepareMessage.call(
		this,
		message,
		contentType,
	);

	return await microsoftApiRequest.call(this, 'POST', `/v1.0/chats/${chatId}/messages`, body);
}
