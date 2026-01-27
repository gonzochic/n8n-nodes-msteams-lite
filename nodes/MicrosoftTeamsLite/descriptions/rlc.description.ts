import type { INodeProperties } from 'n8n-workflow';

export const teamRLC: INodeProperties = {
	displayName: 'Team',
	name: 'teamId',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
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
			displayName: 'From URL',
			name: 'url',
			type: 'string',
			placeholder: 'e.g. https://teams.microsoft.com/l/team/19%3AP8l9gXd6oqlgq…',
			extractValue: {
				type: 'regex',
				regex: 'groupId=([a-f0-9-]+)\\&',
			},
			validation: [
				{
					type: 'regex',
					properties: {
						regex: 'https:\\/\\/teams.microsoft.com\\/.*groupId=[a-f0-9-]+\\&.*',
						errorMessage: 'Not a valid Microsoft Teams URL',
					},
				},
			],
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
};

export const channelRLC: INodeProperties = {
	displayName: 'Channel',
	name: 'channelId',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
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
			// validation missing because no documentation found how these unique ids look like.
		},
	],
};

export const chatRLC: INodeProperties = {
	displayName: 'Chat',
	name: 'chatId',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
	required: true,
	description:
		'Select the chat from the list, by URL, or by ID (find the chat ID after "conversations/" in the URL)',
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			placeholder: 'Select a Chat...',
			typeOptions: {
				searchListMethod: 'getChats',
				searchable: true,
			},
		},
		{
			displayName: 'By ID',
			name: 'id',
			type: 'string',
			placeholder:
				'19:7e2f1174-e8ee-4859-b8b1-a8d1cc63d276_0c5cfdbb-596f-4d39-b557-5d9516c94107@unq.gbl.spaces',
			// validation missing because no documentation found how these unique chat ids look like.
			url: '=https://teams.microsoft.com/l/chat/{{encodeURIComponent($value)}}/0',
		},
	],
};

export const senderRLC: INodeProperties = {
	displayName: 'Sender',
	name: 'sender',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
	required: true,
	description: 'The account to send the message from',
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			placeholder: 'Select your account...',
			typeOptions: {
				searchListMethod: 'getCurrentUser',
				searchable: false,
			},
		},
		{
			displayName: 'By Email',
			name: 'id',
			type: 'string',
			placeholder: 'e.g. user@company.com',
		},
	],
};

export const userRLC: INodeProperties = {
	displayName: 'User',
	name: 'userId',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
	required: true,
	description: 'Select the user from the list or enter their ID',
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			placeholder: 'Select a user...',
			typeOptions: {
				searchListMethod: 'getUsers',
				searchable: true,
			},
		},
		{
			displayName: 'By ID',
			name: 'id',
			type: 'string',
			placeholder: 'e.g. 12345678-1234-1234-1234-123456789012',
			validation: [
				{
					type: 'regex',
					properties: {
						regex: '^([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})[ \t]*',
						errorMessage: 'Not a valid user ID',
					},
				},
			],
			extractValue: {
				type: 'regex',
				regex: '^([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})',
			},
		},
	],
};

export const teamMemberRLC: INodeProperties = {
	displayName: 'Member',
	name: 'memberId',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
	required: true,
	description: 'Select the team member from the list or enter their membership ID',
	typeOptions: {
		loadOptionsDependsOn: ['teamId.value'],
	},
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			placeholder: 'Select a member...',
			typeOptions: {
				searchListMethod: 'getTeamMembers',
				searchable: true,
			},
		},
		{
			displayName: 'By ID',
			name: 'id',
			type: 'string',
			placeholder: 'e.g. MCMjMiMjZGNhNjE...',
		},
	],
};

export const channelMemberRLC: INodeProperties = {
	displayName: 'Member',
	name: 'memberId',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
	required: true,
	description: 'Select the channel member from the list or enter their membership ID',
	typeOptions: {
		loadOptionsDependsOn: ['teamId.value', 'channelId.value'],
	},
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			placeholder: 'Select a member...',
			typeOptions: {
				searchListMethod: 'getChannelMembers',
				searchable: true,
			},
		},
		{
			displayName: 'By ID',
			name: 'id',
			type: 'string',
			placeholder: 'e.g. MCMjMiMjZGNhNjE...',
		},
	],
};
