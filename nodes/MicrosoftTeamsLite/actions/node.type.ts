import type { AllEntities } from 'n8n-workflow';

type NodeMap = {
	chat: 'createGroup' | 'createOneOnOne';
	chatMessage: 'create' | 'get' | 'getAll';
	channel: 'create';
	channelMessage: 'create' | 'get' | 'getAll';
	teamMember: 'add' | 'remove';

};

export type MicrosoftTeamsType = AllEntities<NodeMap>;
