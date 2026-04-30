import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { AadHttpClient } from "@microsoft/sp-http";

import * as strings from "SkillsRepositoryWebPartStrings";
import SkillsRepository from "./components/SkillsRepository";
import { ISkillsRepositoryProps } from "./components/ISkillsRepositoryProps";

const DEFAULT_SKILLS_SITE_URL =
  "/sites/OrgAISkills";
const DEFAULT_SKILLS_LIBRARY_NAME = "Skills";
const DEFAULT_COWORK_ONEDRIVE_PATH = "Documents/Cowork/Skills";
const DEFAULT_COWORK_SUBFOLDER_NAME = "";
const DEFAULT_LOCAL_MACHINE_PATH = "C:\\Users\\{username}\\.agents\\skills";

export interface ISkillsRepositoryWebPartProps {
  /** SharePoint site URL hosting the skills library */
  skillsSiteUrl: string;
  /** Name of the skills library */
  skillsLibraryName: string;
  /** OneDrive root path used for Cowork copies */
  coworkOneDrivePath: string;
  /** Optional subfolder inside the OneDrive root path */
  coworkSubfolderName: string;
  /** Local machine path hint shown to users */
  localMachinePath: string;
}

export default class SkillsRepositoryWebPart extends BaseClientSideWebPart<ISkillsRepositoryWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _graphClient!: AadHttpClient;

  public async onInit(): Promise<void> {
    await super.onInit();
    // Acquire an AAD HTTP client scoped to Microsoft Graph
    this._graphClient = await this.context.aadHttpClientFactory.getClient(
      "https://graph.microsoft.com"
    );
  }

  public render(): void {
    // Property values are editable by the web part owner and defaulted when blank.
    const effectiveConfig = {
      skillsSiteUrl:
        this.properties.skillsSiteUrl?.trim() || DEFAULT_SKILLS_SITE_URL,
      skillsLibraryName:
        this.properties.skillsLibraryName?.trim() ||
        DEFAULT_SKILLS_LIBRARY_NAME,
      coworkOneDrivePath:
        this.properties.coworkOneDrivePath?.trim() ||
        DEFAULT_COWORK_ONEDRIVE_PATH,
      coworkSubfolderName:
        this.properties.coworkSubfolderName?.trim() ||
        DEFAULT_COWORK_SUBFOLDER_NAME,
      localMachinePath:
        this.properties.localMachinePath?.trim() || DEFAULT_LOCAL_MACHINE_PATH,
    };

    const element: React.ReactElement<ISkillsRepositoryProps> =
      React.createElement(SkillsRepository, {
        ...effectiveConfig,
        context: this.context,
        graphClient: this._graphClient,
        isDarkTheme: this._isDarkTheme,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.SourceGroupName,
              groupFields: [
                PropertyPaneTextField("skillsSiteUrl", {
                  label: strings.SkillsSiteUrlLabel,
                  description: strings.SkillsSiteUrlDesc,
                  placeholder: DEFAULT_SKILLS_SITE_URL,
                }),
                PropertyPaneTextField("skillsLibraryName", {
                  label: strings.SkillsLibraryNameLabel,
                  description: strings.SkillsLibraryNameDesc,
                  placeholder: DEFAULT_SKILLS_LIBRARY_NAME,
                }),
              ],
            },
            {
              groupName: strings.DestinationGroupName,
              groupFields: [
                PropertyPaneTextField("coworkOneDrivePath", {
                  label: strings.CoworkPathLabel,
                  description: strings.CoworkPathDesc,
                  placeholder: DEFAULT_COWORK_ONEDRIVE_PATH,
                }),
                PropertyPaneTextField("coworkSubfolderName", {
                  label: strings.CoworkSubfolderLabel,
                  description: strings.CoworkSubfolderDesc,
                  placeholder: DEFAULT_COWORK_SUBFOLDER_NAME,
                }),
                PropertyPaneTextField("localMachinePath", {
                  label: strings.LocalPathLabel,
                  description: strings.LocalPathDesc,
                  placeholder: DEFAULT_LOCAL_MACHINE_PATH,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
