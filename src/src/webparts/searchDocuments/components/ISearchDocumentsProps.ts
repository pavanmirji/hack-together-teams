import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISearchDocumentsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  ctx: WebPartContext;
}
