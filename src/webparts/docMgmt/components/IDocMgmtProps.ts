import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocMgmtProps {
  context: WebPartContext;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
}
