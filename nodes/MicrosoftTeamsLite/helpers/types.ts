// Subscription expires after ~3 days (4318 minutes is the max allowed by Microsoft Graph)
export const SUBSCRIPTION_EXPIRATION_MINUTES = 60;

// Minimum time before expiration to consider subscription valid (5 minutes)
export const SUBSCRIPTION_VALIDITY_THRESHOLD_MS = 5 * 60 * 1000;

export interface TeamResponse {
	id: string;
	displayName: string;
}

export interface ChannelResponse {
	id: string;
	displayName: string;
}

export interface WebhookNotification {
	subscriptionId: string;
	resource: string;
	resourceData: ResourceData;
	tenantId: string;
	subscriptionExpirationDateTime: string;
	lifecycleEvent?: 'reauthorizationRequired' | 'subscriptionRemoved' | 'missed';
}

export interface ResourceData {
	id: string;
	[key: string]: unknown;
}

export interface SubscriptionResponse {
	id: string;
	expirationDateTime: string;
	notificationUrl: string;
	resource: string;
}

export interface ChatMessage {
	id: string;
	messageType: string;
	createdDateTime: string;
	from?: {
		user?: { id: string; displayName: string };
	};
	body: {
		contentType: string;
		content: string;
	};
	[key: string]: unknown;
}
