import type {
	IExecuteFunctions,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { versionDescription } from './actions/versionDescription';
import { router } from './actions/router';
import { listSearch } from './methods';
import { sendAndWaitWebhook } from 'n8n-nodes-base/dist/utils/sendAndWait/utils';

export class MicrosoftTeamsLite implements INodeType {
	description: INodeTypeDescription = versionDescription;
	methods = { listSearch };

	webhook = sendAndWaitWebhook;

	async execute(this: IExecuteFunctions) {
		return await router.call(this);
	}
}
