import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DynamicProperty } from "@microsoft/sp-component-base";

export interface ICostumerProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  context: WebPartContext;
  DeptTitleId: DynamicProperty<string>;
}
