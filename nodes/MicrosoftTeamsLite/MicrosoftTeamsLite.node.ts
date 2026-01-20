import type {
	IExecuteFunctions,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { versionDescription } from './actions/versionDescription';
import { router } from './actions/router';
import { listSearch } from './methods';

// eslint-disable-next-line @n8n/community-nodes/icon-validation
export class MicrosoftTeamsLite implements INodeType {
	description: INodeTypeDescription = versionDescription;
	methods = { listSearch };

	async execute(this: IExecuteFunctions) {
		return await router.call(this);
	}
}
