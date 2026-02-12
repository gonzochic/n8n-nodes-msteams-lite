import {
	NodeOperationError,
	type IDataObject,
	type ILoadOptionsFunctions,
	type INodeListSearchItems,
	type INodeListSearchResult,
	sleep,
} from 'n8n-workflow';

import { filterSortSearchListItems } from '../helpers/utils';
import { microsoftApiRequest } from '../transport';

export async function getChats(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const returnData: INodeListSearchItems[] = [];
	const qs: IDataObject = {
		$expand: 'members',
	};

	let value: IDataObject[] = [];
	let attempts = 5;
	do {
		try {
			value = ((await microsoftApiRequest.call(this, 'GET', '/v1.0/chats', {}, qs)) as IDataObject)
				.value as IDataObject[];
			break;
		} catch (error) {
			if (attempts > 0) {
				await sleep(1000);
				attempts--;
			} else {
				throw new NodeOperationError(this.getNode(), error);
			}
		}
	} while (attempts > 0);

	for (const chat of value) {
		if (!chat.topic) {
			chat.topic = (chat.members as IDataObject[])
				.filter((member: IDataObject) => member.displayName)
				.map((member: IDataObject) => member.displayName)
				.join(', ');
		}
		const chatName = `${chat.topic || '(no title) - ' + chat.id} (${chat.chatType})`;
		const chatId = chat.id;
		const url = chat.webUrl as string;
		returnData.push({
			name: chatName,
			value: chatId as string,
			url,
		});
	}

	const results = filterSortSearchListItems(returnData, filter);
	return { results };
}

export async function getTeams(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const returnData: INodeListSearchItems[] = [];
	const { value } = await microsoftApiRequest.call(this, 'GET', '/v1.0/me/joinedTeams');

	for (const team of value) {
		const teamName = team.displayName;
		const teamId = team.id;
		// let channelId: string = '';

		// try {
		// 	const channels = await microsoftApiRequestAllItems.call(
		// 		this,
		// 		'value',
		// 		'GET',
		// 		`/v1.0/teams/${teamId}/channels`,
		// 		{},
		// 	);

		// 	if (channels.length > 0) {
		// 		channelId = channels.find((channel: IDataObject) => channel.displayName === 'General').id;
		// 		if (!channelId) {
		// 			channelId = channels[0].id;
		// 		}
		// 	}
		// } catch (error) {}

		returnData.push({
			name: teamName,
			value: teamId,
			// url: channelId
			// 	? `https://teams.microsoft.com/l/team/${channelId}/conversations?groupId=${teamId}&tenantId=${team.tenantId}`
			// 	: undefined,
		});
	}
	const results = filterSortSearchListItems(returnData, filter);

	return { results };
}

export async function getChannels(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const returnData: INodeListSearchItems[] = [];
	const teamId = this.getCurrentNodeParameter('teamId', { extractValue: true }) as string;
	const operation = this.getNodeParameter('operation', 0) as string;
	const resource = this.getNodeParameter('resource', 0) as string;

	const excludeGeneralChannel = ['deleteChannel'];

	if (resource === 'channel') excludeGeneralChannel.push('update');

	const { value } = await microsoftApiRequest.call(this, 'GET', `/v1.0/teams/${teamId}/channels`);

	for (const channel of value) {
		if (channel.displayName === 'General' && excludeGeneralChannel.includes(operation)) {
			continue;
		}
		const channelName = channel.displayName;
		const channelId = channel.id;
		const url = channel.webUrl;
		returnData.push({
			name: channelName,
			value: channelId,
			url,
		});
	}

	const results = filterSortSearchListItems(returnData, filter);
	return { results };
}

export async function getCurrentUser(
	this: ILoadOptionsFunctions,
	// eslint-disable-next-line @typescript-eslint/no-unused-vars
	filter?: string,
): Promise<INodeListSearchResult> {
	const user = (await microsoftApiRequest.call(this, 'GET', '/v1.0/me')) as IDataObject;

	const result = {
		name: `${user.displayName} (${user.userPrincipalName})`,
		value: user.userPrincipalName as string,
	};

	return { results: [result] };
}

export async function getUsers(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const returnData: INodeListSearchItems[] = [];
	const qs: IDataObject = {
		$top: 100,
	};

	if (filter) {
		qs.$filter = `startswith(displayName,'${filter}') or startswith(userPrincipalName,'${filter}')`;
	}

	const { value } = (await microsoftApiRequest.call(
		this,
		'GET',
		'/v1.0/users',
		{},
		qs,
	)) as { value: IDataObject[] };

	for (const user of value) {
		returnData.push({
			name: `${user.displayName} (${user.userPrincipalName})`,
			value: user.id as string,
		});
	}

	const results = filterSortSearchListItems(returnData, filter);
	return { results };
}

export async function getTeamMembers(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const returnData: INodeListSearchItems[] = [];
	const teamId = this.getCurrentNodeParameter('teamId', { extractValue: true }) as string;

	const { value } = (await microsoftApiRequest.call(
		this,
		'GET',
		`/v1.0/teams/${teamId}/members`,
	)) as { value: IDataObject[] };

	for (const member of value) {
		returnData.push({
			name: `${member.displayName} (${member.email || 'No email'})`,
			value: member.id as string,
		});
	}

	const results = filterSortSearchListItems(returnData, filter);
	return { results };
}

