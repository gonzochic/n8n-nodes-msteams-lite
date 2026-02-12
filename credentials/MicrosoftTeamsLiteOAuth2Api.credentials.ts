import type { ICredentialType, INodeProperties, Icon } from 'n8n-workflow';

export class MicrosoftTeamsLiteOAuth2Api implements ICredentialType {
    name = 'microsoftTeamsLiteOAuth2Api';

    extends = ['oAuth2Api'];

    icon: Icon = 'file:../nodes/MicrosoftTeamsLite/teams.svg';

    displayName = 'Reduced Permissions Microsoft Teams OAuth2 API';

    documentationUrl = 'https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow';

    properties: INodeProperties[] = [
        {
            displayName: 'Grant Type',
            name: 'grantType',
            type: 'hidden',
            default: 'authorizationCode',
        },
        {
            displayName: 'Authorization URL',
            name: 'authUrl',
            type: 'string',
            default: 'https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/authorize',
        },
        {
            displayName: 'Access Token URL',
            name: 'accessTokenUrl',
            type: 'string',
            default: 'https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/token',
        },
        {
            displayName: 'Auth URI Query Parameters',
            name: 'authQueryParameters',
            type: 'hidden',
            default: 'response_mode=query',
        },
        {
            displayName: 'Authentication',
            name: 'authentication',
            type: 'hidden',
            default: 'body',
        },
        {
            displayName: 'Scope',
            name: 'scope',
            type: 'string',
            default:
                'Chat.Create ChatMessage.Read ChatMessage.Send Chat.ReadWrite User.Read.All Channel.Create ChannelMessage.Read.All ChannelMessage.Send TeamMember.ReadWriteNonOwnerRole.All offline_access',
            description: 'The scope for the OAuth2 request',
        },
    ];
}