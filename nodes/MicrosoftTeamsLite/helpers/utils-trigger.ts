import type { IHookFunctions, IDataObject } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';

import type { ChannelResponse, SubscriptionResponse } from './types';
import { SUBSCRIPTION_EXPIRATION_MINUTES } from './types';
import { microsoftApiRequest } from '../transport';

export function extractChatIdFromResource(resourcePath: string): string | null {
	// Match format: chats('chatId')/messages('messageId') used by Graph API notifications
	const match = resourcePath.match(/chats\('([^']+)'\)/);
	if (match && match[1]) {
		return decodeURIComponent(match[1]);
	}
	return null;
}

export function extractChannelInfoFromResource(
	resourcePath: string,
): { teamId: string; channelId: string } | null {
	// Match format: teams('teamId')/channels('channelId')/messages('messageId')
	const teamMatch = resourcePath.match(/teams\('([^']+)'\)/);
	const channelMatch = resourcePath.match(/channels\('([^']+)'\)/);

	if (teamMatch && teamMatch[1] && channelMatch && channelMatch[1]) {
		return {
			teamId: decodeURIComponent(teamMatch[1]),
			channelId: decodeURIComponent(channelMatch[1]),
		};
	}
	return null;
}

export async function fetchAllChannels(
	this: IHookFunctions,
	teamId: string,
): Promise<ChannelResponse[]> {
	const { value: channels } = (await microsoftApiRequest.call(
		this,
		'GET',
		`/v1.0/teams/${teamId}/channels`,
	)) as { value: ChannelResponse[] };
	return channels;
}

export async function createSubscription(
	this: IHookFunctions,
	webhookUrl: string,
	resourcePath: string,
): Promise<SubscriptionResponse> {
	const expirationTime = new Date(Date.now() + SUBSCRIPTION_EXPIRATION_MINUTES * 60 * 1000).toISOString();
	const body: IDataObject = {
		changeType: 'created',
		notificationUrl: webhookUrl,
		resource: resourcePath,
		expirationDateTime: expirationTime,
		latestSupportedTlsVersion: 'v1_2',
		lifecycleNotificationUrl: webhookUrl,
	};

	const response = (await microsoftApiRequest.call(
		this,
		'POST',
		'/v1.0/subscriptions',
		body,
	)) as SubscriptionResponse;

	return response;
}

export async function getResourcePath(
	this: IHookFunctions,
	event: string,
): Promise<string | string[]> {
	switch (event) {
		case 'newChat': {
			return '/me/chats';
		}

		case 'newChatMessage': {
			const watchAllChats = this.getNodeParameter('watchAllChats', false, {
				extractValue: true,
			}) as boolean;

			if (watchAllChats) {
				return '/me/chats/getAllMessages';
			} else {
				const chatId = this.getNodeParameter('chatId', undefined, { extractValue: true }) as string;
				return `/chats/${decodeURIComponent(chatId)}/messages`;
			}
		}

		case 'newChannel': {
			const teamId = this.getNodeParameter('teamId', undefined, { extractValue: true }) as string;
			return `/teams/${teamId}/channels`;
		}

		case 'newChannelMessage': {
			const teamId = this.getNodeParameter('teamId', undefined, { extractValue: true }) as string;
			const watchAllChannels = this.getNodeParameter('watchAllChannels', false, {
				extractValue: true,
			}) as boolean;

			if (watchAllChannels) {
				const channels = await fetchAllChannels.call(this, teamId);
				return channels.map((channel) => `/teams/${teamId}/channels/${channel.id}/messages`);
			} else {
				const channelId = this.getNodeParameter('channelId', undefined, {
					extractValue: true,
				}) as string;
				return `/teams/${teamId}/channels/${decodeURIComponent(channelId)}/messages`;
			}
		}

		case 'newTeamMember': {
			const teamId = this.getNodeParameter('teamId', undefined, { extractValue: true }) as string;
			return `/teams/${teamId}/members`;
		}

		default: {
			throw new NodeOperationError(this.getNode(), {
				message: `Invalid event: ${event}`,
				description: `The selected event "${event}" is not recognized.`,
			});
		}
	}
}
