import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IKpfasrsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext; // Добавляем контекст для доступа к SharePoint API
}