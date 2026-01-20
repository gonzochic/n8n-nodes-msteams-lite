import type { AllEntities } from 'n8n-workflow';

type NodeMap = {
	chatMessage: 'create' | 'get' | 'getAll';
};

export type MicrosoftTeamsType = AllEntities<NodeMap>;
