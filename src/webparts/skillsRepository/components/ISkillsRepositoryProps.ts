import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AadHttpClient } from "@microsoft/sp-http";

export interface ISkillsRepositoryProps {
  /** Resolved SharePoint site URL hosting the Skills library */
  skillsSiteUrl: string;
  /** Name of the SharePoint document library containing skill files */
  skillsLibraryName: string;
  /** OneDrive folder path for Cowork/Copilot agent usage */
  coworkOneDrivePath: string;
  /** Optional sub-folder within the OneDrive path */
  coworkSubfolderName: string;
  /** Local machine path displayed as guidance after download */
  localMachinePath: string;
  /** SPFx web part context */
  context: WebPartContext;
  /** Authenticated Graph API client */
  graphClient: AadHttpClient;
  /** Whether dark theme is active */
  isDarkTheme: boolean;
  /** Current user's display name */
  userDisplayName: string;
  /** Current user's email */
  userEmail: string;
}
