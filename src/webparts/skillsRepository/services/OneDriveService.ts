import { AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";
import { ISkillItem } from "../components/ISkillsRepositoryState";

interface IGraphDriveItem {
  id: string;
  name: string;
  folder?: object;
}

/**
 * Service for copying skill files into the current user's OneDrive
 * using the Microsoft Graph API (via an AAD-authenticated HTTP client).
 */
export class OneDriveService {
  private graphClient: AadHttpClient;
  private graphBaseUrl = "https://graph.microsoft.com/v1.0";

  constructor(graphClient: AadHttpClient) {
    this.graphClient = graphClient;
  }

  /**
   * Copies a skill file from SharePoint into the user's OneDrive Cowork folder.
   *
   * Strategy:
   *  1. Ensure the destination folder exists (create if missing).
   *  2. Fetch the file content from the SharePoint download URL.
   *  3. PUT the content into OneDrive via Graph.
   *
   * @param skill       The skill to copy.
   * @param fileContent Raw text content of the skill file (fetched by SkillsService).
   * @param coworkOneDrivePath  Root OneDrive folder path from config.
   * @param coworkSubfolderName Optional sub-folder from config.
   */
  public async copySkillToOneDrive(
    skill: ISkillItem,
    fileContent: string,
    coworkOneDrivePath: string,
    coworkSubfolderName: string
  ): Promise<void> {
    const destinationFolder = this.buildCoworkDestinationPath(
      coworkOneDrivePath,
      coworkSubfolderName,
      skill.category !== "General" ? skill.category : undefined
    );
    await this.ensureFolderPath(destinationFolder);
    await this.uploadFile(destinationFolder, skill.fileName, fileContent);
  }

  private buildCoworkDestinationPath(
    coworkOneDrivePath: string,
    coworkSubfolderName: string,
    skillFolderName?: string
  ): string {
    const parts = [coworkOneDrivePath];
    if (coworkSubfolderName) {
      parts.push(coworkSubfolderName);
    }
    if (skillFolderName) {
      parts.push(skillFolderName);
    }
    return parts.join("/");
  }

  /**
   * Ensures all folders in the given path exist in the user's OneDrive,
   * creating any missing segments one level at a time.
   */
  private async ensureFolderPath(folderPath: string): Promise<void> {
    const segments = folderPath.split("/").filter(Boolean);
    let currentPath = "";

    for (const segment of segments) {
      const parentPath = currentPath || "root";
      currentPath = currentPath ? `${currentPath}/${segment}` : segment;

      const exists = await this.folderExists(currentPath);
      if (!exists) {
        await this.createFolder(parentPath, segment);
      }
    }
  }

  /**
   * Checks whether a OneDrive folder path exists.
   */
  private async folderExists(folderPath: string): Promise<boolean> {
    const url = `${this.graphBaseUrl}/me/drive/root:/${encodeURIComponent(folderPath)}`;
    const response: HttpClientResponse = await this.graphClient.get(
      url,
      AadHttpClient.configurations.v1
    );
    return response.ok;
  }

  /**
   * Creates a folder as a child of parentPath in the user's OneDrive.
   * parentPath of "root" targets the drive root.
   */
  private async createFolder(
    parentPath: string,
    folderName: string
  ): Promise<IGraphDriveItem> {
    const url =
      parentPath === "root"
        ? `${this.graphBaseUrl}/me/drive/root/children`
        : `${this.graphBaseUrl}/me/drive/root:/${encodeURIComponent(parentPath)}:/children`;

    const body = JSON.stringify({
      name: folderName,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename",
    });

    const response: HttpClientResponse = await this.graphClient.post(
      url,
      AadHttpClient.configurations.v1,
      {
        headers: { "Content-Type": "application/json" },
        body,
      }
    );

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`Failed to create folder "${folderName}": ${err}`);
    }

    return response.json() as Promise<IGraphDriveItem>;
  }

  /**
   * Uploads (PUT) text content as a file into the specified OneDrive folder.
   * Uses the simple upload endpoint (files up to 4 MB).
   */
  private async uploadFile(
    folderPath: string,
    fileName: string,
    content: string
  ): Promise<void> {
    const url = `${this.graphBaseUrl}/me/drive/root:/${encodeURIComponent(
      folderPath
    )}/${encodeURIComponent(fileName)}:/content`;

    const response: HttpClientResponse = await this.graphClient.fetch(
      url,
      AadHttpClient.configurations.v1,
      {
        method: "PUT",
        headers: { "Content-Type": "text/markdown" },
        body: content,
      }
    );

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`Failed to upload "${fileName}": ${err}`);
    }
  }
}
