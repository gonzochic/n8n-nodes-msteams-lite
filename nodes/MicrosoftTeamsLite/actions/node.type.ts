import type { AllEntities } from 'n8n-workflow';

type NodeMap = {
	chatMessage: 'create' | 'get' | 'getAll';
	channel: 'create';
	channelMessage: 'create' | 'get' | 'getAll';
	teamMember: 'add' | 'remove';
	channelMember: 'add' | 'remove';
};

export type MicrosoftTeamsType = AllEntities<NodeMap>;
