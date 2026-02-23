/* eslint-disable no-useless-escape */
/* eslint-disable @typescript-eslint/no-unused-vars */
import type {
	IExecuteFunctions,
	INodeType,
	INodeTypeDescription,
	IHookFunctions,
	IWebhookFunctions,
	IWebhookResponseData,
	IDataObject,
	ILoadOptionsFunctions,
	JsonObject,
	INodeExecutionData,
} from 'n8n-workflow';
import { NodeApiError, NodeConnectionTypes } from 'n8n-workflow';

import type { WebhookNotification, SubscriptionResponse, ChatMessage } from './helpers/types';
import { SUBSCRIPTION_VALIDITY_THRESHOLD_MS } from './helpers/types';
import {
	createSubscription,
	renewSubscription,
	getResourcePath,
	extractChatIdFromResource,
	extractChannelInfoFromResource,
} from './helpers/utils-trigger';
import { listSearch } from './methods';
import { microsoftApiRequest, microsoftApiRequestAllItems } from './transport';

export class MicrosoftTeamsLiteTrigger implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Microsoft Teams (Lite) Trigger',
		name: 'microsoftTeamsLiteTrigger',
		icon: 'file:teams.svg',
		group: ['trigger'],
		version: 1,
		description:
			'Triggers workflows in n8n based on events from Microsoft Teams, such as new messages or team updates, using specified configurations. It uses minimal permissions',
		subtitle: 'Microsoft Teams (Lite) Trigger',
		defaults: {
			name: 'Microsoft Teams (Lite) Trigger',
		},
		credentials: [
			{
				name: 'microsoftTeamsLiteOAuth2Api',
				required: true,
			},
		],
		inputs: [],
		outputs: [NodeConnectionTypes.Main],
		usableAsTool: true,
		webhooks: [
			{
				name: 'default',
				httpMethod: 'POST',
				responseMode: 'onReceived',
				path: 'webhook',
			},
		],
		properties: [
			{
				displayName: 'Trigger On',
				name: 'event',
				type: 'options',
				default: 'newChatMessage',
				options: [
					{
						name: 'New Channel',
						value: 'newChannel',
						description: 'A new channel is created in a team',
					},
					{
						name: 'New Channel Message',
						value: 'newChannelMessage',
						description: 'A message is posted to a channel',
					},
					{
						name: 'New Chat',
						value: 'newChat',
						description: 'A new chat is created',
					},
					{
						name: 'New Chat Message',
						value: 'newChatMessage',
						description: 'A message is posted to a chat',
					},
					{
						name: 'New Team Member',
						value: 'newTeamMember',
						description: 'A new member joins a team',
					},
				],
				description: 'Select the event to trigger the workflow',
			},
			// Chat-related properties
			{
				displayName: 'Watch All Chats',
				name: 'watchAllChats',
				type: 'boolean',
				default: false,
				description: 'Whether to watch for the event in all the available chats',
				displayOptions: {
					show: {
						event: ['newChatMessage'],
					},
				},
			},
			{
				displayName: 'Chat',
				name: 'chatId',
				type: 'resourceLocator',
				default: {
					mode: 'list',
					value: '',
				},
				required: true,
				description: 'Select a chat from the list, enter an ID or a URL',
				modes: [
					{
						displayName: 'From List',
						name: 'list',
						type: 'list',
						placeholder: 'Select a chat...',
						typeOptions: {
							searchListMethod: 'getChats',
							searchable: true,
						},
					},
					{
						displayName: 'By ID',
						name: 'id',
						type: 'string',
						placeholder: '19:7e2f1174-e8ee-4859-b8b1-a8d1cc63d276@unq.gbl.spaces',
					},
					{
						displayName: 'By URL',
						name: 'url',
						type: 'string',
						placeholder: 'https://teams.microsoft.com/_#/conversations/CHAT_ID',
						extractValue: {
							type: 'regex',
							regex: /conversations\/([^\/?]+)/i,
						},
					},
				],
				displayOptions: {
					show: {
						event: ['newChatMessage'],
						watchAllChats: [false],
					},
				},
			},
			// Team-related properties
			{
				displayName: 'Team',
				name: 'teamId',
				type: 'resourceLocator',
				default: {
					mode: 'list',
					value: '',
				},
				required: true,
				description:
					'Select the team from the list, by URL, or by ID (the ID is the "groupId" parameter in the URL you get from "Get a link to the team")',
				modes: [
					{
						displayName: 'From List',
						name: 'list',
						type: 'list',
						placeholder: 'e.g. My Team',
						typeOptions: {
							searchListMethod: 'getTeams',
							searchable: true,
						},
					},
					{
						displayName: 'By ID',
						name: 'id',
						type: 'string',
						placeholder: 'e.g. 61165b04-e4cc-4026-b43f-926b4e2a7182',
						validation: [
							{
								type: 'regex',
								properties: {
									regex: '^([0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12})[ \t]*',
									errorMessage: 'Not a valid Microsoft Teams Team ID',
								},
							},
						],
						extractValue: {
							type: 'regex',
							regex: '^([0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12})',
						},
					},
				],
				displayOptions: {
					show: {
						event: ['newChannel', 'newChannelMessage', 'newTeamMember'],
					},
				},
			},
			// Channel-related properties
			{
				displayName: 'Watch All Channels',
				name: 'watchAllChannels',
				type: 'boolean',
				default: false,
				description: 'Whether to watch for the event in all channels of the selected team',
				displayOptions: {
					show: {
						event: ['newChannelMessage'],
					},
				},
			},
			{
				displayName: 'Channel',
				name: 'channelId',
				type: 'resourceLocator',
				default: {
					mode: 'list',
					value: '',
				},
				required: true,
				description:
					'Select the channel from the list, by URL, or by ID (the ID is the "threadId" in the URL)',
				typeOptions: {
					loadOptionsDependsOn: ['teamId.value'],
				},
				modes: [
					{
						displayName: 'From List',
						name: 'list',
						type: 'list',
						placeholder: 'Select a Channel...',
						typeOptions: {
							searchListMethod: 'getChannels',
							searchable: true,
						},
					},
					{
						displayName: 'By ID',
						name: 'id',
						type: 'string',
						placeholder: '19:-xlxyqXNSCxpI1SDzgQ_L9ZvzSR26pgphq1BJ9y7QJE1@thread.tacv2',
					},
				],
				displayOptions: {
					show: {
						event: ['newChannelMessage'],
						watchAllChannels: [false],
					},
				},
			},
			// Fetch full message option
			{
				displayName: 'Fetch Full Message',
				name: 'fetchFullMessage',
				type: 'boolean',
				default: false,
				description:
					'Whether to fetch the full message content from the API. When disabled, only the notification metadata (IDs) is returned.',
				displayOptions: {
					show: {
						event: ['newChatMessage', 'newChannelMessage'],
					},
				},
			},
			{
				displayName: 'Ignore Own Messages',
				name: 'ignoreOwnMessages',
				type: 'boolean',
				default: false,
				description:
					'Whether to ignore messages sent by the authenticated user. Useful to prevent loops when a workflow both sends and listens for messages.',
				displayOptions: {
					show: {
						event: ['newChatMessage', 'newChannelMessage'],
					},
				},
			},
		],
	};

	methods = {
		listSearch,
	};

	webhookMethods = {
		default: {
			async checkExists(this: IHookFunctions): Promise<boolean> {
				const event = this.getNodeParameter('event', 0) as string;
				const webhookUrl = this.getNodeWebhookUrl('default');
				const webhookData = this.getWorkflowStaticData('node');

				try {
					const subscriptions = (await microsoftApiRequestAllItems.call(
						this as unknown as ILoadOptionsFunctions,
						'value',
						'GET',
						'/v1.0/subscriptions',
					)) as SubscriptionResponse[];

					const matchingSubscriptions = subscriptions.filter(
						(subscription) => subscription.notificationUrl === webhookUrl,
					);

					const now = new Date();
					const validSubscriptions = matchingSubscriptions.filter((subscription) => {
						const expiration = new Date(subscription.expirationDateTime);
						return expiration.getTime() - now.getTime() > SUBSCRIPTION_VALIDITY_THRESHOLD_MS;
					});

					const resourcePaths = await getResourcePath.call(this, event);
					const requiredResources = Array.isArray(resourcePaths) ? resourcePaths : [resourcePaths];

					const subscribedResources = validSubscriptions.map((sub) => sub.resource);
					const allResourcesSubscribed = requiredResources.every((resource) =>
						subscribedResources.includes(resource),
					);

					if (allResourcesSubscribed) {
						webhookData.subscriptionIds = validSubscriptions.map((sub) => sub.id);
						return true;
					}

					return false;
				} catch (error) {
					return false;
				}
			},

			async create(this: IHookFunctions): Promise<boolean> {
				const event = this.getNodeParameter('event', 0) as string;
				const webhookUrl = this.getNodeWebhookUrl('default');
				const webhookData = this.getWorkflowStaticData('node');

				if (!webhookUrl?.startsWith('https://')) {
					throw new NodeApiError(this.getNode(), {
						message: 'Invalid Notification URL',
						description: `The webhook URL "${webhookUrl}" is invalid. Microsoft Graph requires an HTTPS URL.`,
					});
				}

				const resourcePaths = await getResourcePath.call(this, event);
				const subscriptionIds: string[] = [];

				if (Array.isArray(resourcePaths)) {
					await Promise.all(
						resourcePaths.map(async (resource) => {
							const subscription = await createSubscription.call(this, webhookUrl, resource);
							subscriptionIds.push(subscription.id);
							return subscription;
						}),
					);

					webhookData.subscriptionIds = subscriptionIds;
				} else {
					const subscription = await createSubscription.call(this, webhookUrl, resourcePaths);
					webhookData.subscriptionIds = [subscription.id];
				}

				return true;
			},

			async delete(this: IHookFunctions): Promise<boolean> {
				const webhookData = this.getWorkflowStaticData('node');
				const storedIds = webhookData.subscriptionIds as string[] | undefined;

				if (!Array.isArray(storedIds)) {
					return false;
				}

				try {
					await Promise.all(
						storedIds.map(async (subscriptionId) => {
							try {
								await microsoftApiRequest.call(
									this as unknown as IExecuteFunctions,
									'DELETE',
									`/v1.0/subscriptions/${subscriptionId}`,
								);
							} catch (error) {
								if ((error as JsonObject).httpStatusCode !== 404) {
									throw error;
								}
							}
						}),
					);

					delete webhookData.subscriptionIds;
					return true;
				} catch (error) {
					return false;
				}
			},
		},
	};

	async webhook(this: IWebhookFunctions): Promise<IWebhookResponseData> {
		const req = this.getRequestObject();
		const res = this.getResponseObject();

		// Handle Microsoft Graph validation request
		if (req.query.validationToken) {
			res.status(200).send(req.query.validationToken);
			return { noWebhookResponse: true };
		}

		const event = this.getNodeParameter('event', 'newChatMessage') as string;
		const fetchFullMessage = this.getNodeParameter('fetchFullMessage', false) as boolean;
		const ignoreOwnMessages = this.getNodeParameter('ignoreOwnMessages', false) as boolean;

		const allNotifications = req.body.value as WebhookNotification[];

		// Separate lifecycle notifications from regular notifications
		const lifecycleNotifications = allNotifications.filter((n) => n.lifecycleEvent);
		const eventNotifications = allNotifications.filter((n) => !n.lifecycleEvent);

		// Handle lifecycle notifications
		if (lifecycleNotifications.length > 0) {
			const webhookUrl = this.getNodeWebhookUrl('default') as string;

			await Promise.all(
				lifecycleNotifications.map(async (notification) => {
					try {
						if (notification.lifecycleEvent === 'reauthorizationRequired') {
							await renewSubscription.call(this, notification.subscriptionId);
						} else if (notification.lifecycleEvent === 'subscriptionRemoved') {
							// Recreate subscription using the resource from the notification
							await createSubscription.call(
								this as unknown as IHookFunctions,
								webhookUrl,
								notification.resource,
							);
						}
						// 'missed' lifecycle events don't require action — Graph will resend
					} catch {
						// Renewal/recreation failed — subscription will be recovered on next
						// workflow activation via checkExists()
					}
				}),
			);

			// If there are no regular notifications, don't trigger the workflow
			if (eventNotifications.length === 0) {
				return { noWebhookResponse: true };
			}
		}

		// Opportunistic renewal: if any subscription expires within 24 hours, renew it
		const RENEWAL_THRESHOLD_MS = 24 * 60 * 60 * 1000;
		for (const notification of eventNotifications) {
			if (notification.subscriptionExpirationDateTime) {
				const expiresAt = new Date(notification.subscriptionExpirationDateTime).getTime();
				if (expiresAt - Date.now() < RENEWAL_THRESHOLD_MS) {
					// Fire-and-forget renewal — don't block notification processing
					renewSubscription.call(this, notification.subscriptionId).catch(() => {});
				}
			}
		}

		const isMessageEvent = ['newChatMessage', 'newChannelMessage'].includes(event);
		const needsFetch = isMessageEvent && (fetchFullMessage || ignoreOwnMessages);

		// If no fetch needed, return notification data only
		if (!needsFetch) {
			const response: IWebhookResponseData = {
				workflowData: eventNotifications.map((notification) => [
					{
						json: (notification.resourceData as IDataObject) ?? notification,
					} as INodeExecutionData,
				]),
			};
			return response;
		}

		// Fetch authenticated user ID once if we need to filter own messages
		let selfId: string | undefined;
		if (ignoreOwnMessages) {
			const me = (await microsoftApiRequest.call(this, 'GET', '/v1.0/me')) as IDataObject;
			selfId = me.id as string;
		}

		// Fetch full message content for each notification
		const workflowData = (
			await Promise.all(
				eventNotifications.map(async (notification) => {
					try {
						let fullMessage: ChatMessage;

						if (event === 'newChatMessage') {
							// Extract chatId from notification resource path
							const chatId = extractChatIdFromResource(notification.resource);
							const messageId = notification.resourceData?.id;

							if (!chatId || !messageId) {
								return [
									{
										json: {
											...(notification.resourceData as IDataObject),
											_subscriptionId: notification.subscriptionId,
											_tenantId: notification.tenantId,
											_fetchError: 'Could not extract chatId or messageId from notification',
											_originalNotification: notification,
										} as IDataObject,
									} as INodeExecutionData,
								];
							}

							fullMessage = (await microsoftApiRequest.call(
								this,
								'GET',
								`/v1.0/chats/${chatId}/messages/${messageId}`,
							)) as ChatMessage;
						} else {
							// newChannelMessage - extract teamId and channelId
							const channelInfo = extractChannelInfoFromResource(notification.resource);
							const messageId = notification.resourceData?.id;

							if (!channelInfo || !messageId) {
								return [
									{
										json: {
											...(notification.resourceData as IDataObject),
											_subscriptionId: notification.subscriptionId,
											_tenantId: notification.tenantId,
											_fetchError:
												'Could not extract teamId, channelId or messageId from notification',
											_originalNotification: notification,
										} as IDataObject,
									} as INodeExecutionData,
								];
							}

							fullMessage = (await microsoftApiRequest.call(
								this,
								'GET',
								`/v1.0/teams/${channelInfo.teamId}/channels/${channelInfo.channelId}/messages/${messageId}`,
							)) as ChatMessage;
						}

						// Skip messages sent by the authenticated user
						if (selfId && fullMessage.from?.user?.id === selfId) {
							return null;
						}

						if (fetchFullMessage) {
							return [
								{
									json: {
										...fullMessage,
										_subscriptionId: notification.subscriptionId,
										_tenantId: notification.tenantId,
									} as IDataObject,
								} as INodeExecutionData,
							];
						}

						// ignoreOwnMessages is enabled but fetchFullMessage is not — return metadata only
						return [
							{
								json: (notification.resourceData as IDataObject) ?? notification,
							} as INodeExecutionData,
						];
					} catch (error) {
						// On API error, return notification data with error info
						return [
							{
								json: {
									...(notification.resourceData as IDataObject),
									_subscriptionId: notification.subscriptionId,
									_tenantId: notification.tenantId,
									_fetchError: (error as Error).message || 'Failed to fetch full message',
									_originalNotification: notification,
								} as IDataObject,
							} as INodeExecutionData,
						];
					}
				}),
			)
		).filter((item) => item !== null);

		// If all notifications were filtered out, don't trigger the workflow
		if (workflowData.length === 0) {
			return { noWebhookResponse: true };
		}

		return { workflowData };
	}
}
