import { MSGraphClientFactory } from '@microsoft/sp-http';

export interface IMeDemoProps {
  isMicrosoftTeams: boolean;
  msGraphClientFactory: MSGraphClientFactory;
}
