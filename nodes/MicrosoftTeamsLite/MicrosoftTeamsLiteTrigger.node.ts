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

import type { WebhookNotification, SubscriptionResponse } from './helpers/types';
import { createSubscription, getResourcePath } from './helpers/utils-trigger';
import { listSearch } from './methods';
import { microsoftApiRequest, microsoftApiRequestAllItems } from './transport';

export class MicrosoftTeamsLiteTrigger implements INodeType {
    description: INodeTypeDescription = {
        displayName: 'Microsoft Teams Trigger (Lite)',
        name: 'microsoftTeamsLiteTrigger',
        icon: 'file:teams.svg',
        group: ['trigger'],
        version: 1,
        description:
            'Triggers workflows in n8n based on events from Microsoft Teams, such as new messages or team updates, using specified configurations. It uses minimal permissions',
        subtitle: 'Microsoft Teams Trigger (Lite)',
        defaults: {
            name: 'Microsoft Teams Trigger (Lite)',
        },
        credentials: [
            {
                name: 'microsoftOAuth2Api',
                required: true,
            },
        ],
        inputs: [],
        outputs: [NodeConnectionTypes.Main],
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
                        name: 'New Chat',
                        value: 'newChat',
                        description: 'A new chat is created',
                    },
                    {
                        name: 'New Chat Message',
                        value: 'newChatMessage',
                        description: 'A message is posted to a chat',
                    },
                ],
                description: 'Select the event to trigger the workflow',
            },
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
                    const thresholdMs = 5 * 60 * 1000;
                    const validSubscriptions = matchingSubscriptions.filter((subscription) => {
                        const expiration = new Date(subscription.expirationDateTime);
                        return expiration.getTime() - now.getTime() > thresholdMs;
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

        const eventNotifications = req.body.value as WebhookNotification[];
        const response: IWebhookResponseData = {
            workflowData: eventNotifications.map((event) => [
                {
                    json: (event.resourceData as IDataObject) ?? event,
                } as INodeExecutionData,
            ]),
        };

        return response;
    }
}
