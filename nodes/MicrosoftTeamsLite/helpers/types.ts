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
