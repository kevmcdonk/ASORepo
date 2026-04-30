import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { ISkillItem } from "../components/ISkillsRepositoryState";

interface ISpFileItem {
  Id: number;
  FileLeafRef: string;
  FileRef: string;
  File: {
    Length: string;
    TimeLastModified: string;
    Author: { Title: string };
  };
  FileDirRef: string;
  OData__ExtendedDescription?: string;
  UniqueId: string;
}

/**
 * Service for reading skills from the SharePoint document library.
 * Uses the SharePoint REST API directly (no external dependencies).
 */
export class SkillsService {
  private context: WebPartContext;
  private siteUrl: string;
  private libraryName: string;

  constructor(
    context: WebPartContext,
    siteUrl: string,
    libraryName: string
  ) {
    this.context = context;
    this.siteUrl = this.resolveSiteUrl(siteUrl);
    this.libraryName = libraryName;
  }

  private resolveSiteUrl(siteUrl: string): string {
    const trimmed = (siteUrl || "").trim();
    const tenantOrigin = new URL(this.context.pageContext.web.absoluteUrl).origin;

    if (!trimmed) {
      return this.context.pageContext.web.absoluteUrl.replace(/\/$/, "");
    }

    if (/^https?:\/\//i.test(trimmed)) {
      return trimmed.replace(/\/$/, "");
    }

    if (trimmed.startsWith("/")) {
      return `${tenantOrigin}${trimmed}`.replace(/\/$/, "");
    }

    return `${tenantOrigin}/${trimmed}`.replace(/\/$/, "");
  }

  /**
   * Fetches all skill files from the configured SharePoint library.
   * Returns files recursively across all folders (categories).
   */
  public async getSkills(): Promise<ISkillItem[]> {
    const encodedLib = encodeURIComponent(this.libraryName);
    const selectFields = [
      "Id",
      "FileLeafRef",
      "FileRef",
      "FileDirRef",
      "UniqueId",
      "OData__ExtendedDescription",
      "File/Length",
      "File/TimeLastModified",
      "File/Author/Title",
    ].join(",");

    const url =
      `${this.siteUrl}/_api/web/lists/getbytitle('${encodedLib}')/items` +
      `?$select=${selectFields}` +
      `&$expand=File,File/Author` +
      `&$filter=FSObjType eq 0` + // files only, not folders
      `&$top=500`;

      try {
          const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(
        `Failed to load skills: ${response.status} ${response.statusText}`
      );
    }

    const data = await response.json();
    const items: ISpFileItem[] = data.value || [];

    return items
      .filter((item) => item.FileLeafRef?.toLowerCase().endsWith(".md"))
      .map((item) => this.mapToSkillItem(item));
      }
      catch (error) {
          console.error("Error fetching skills from SharePoint:", error);
          throw error;
      }
  }

  /**
   * Returns the download URL for a file given its server-relative URL.
   */
  public getDownloadUrl(serverRelativeUrl: string): string {
    return `${this.siteUrl}/_layouts/download.aspx?SourceUrl=${encodeURIComponent(serverRelativeUrl)}`;
  }

  /**
   * Triggers a browser download for the given skill file.
   */
  public async downloadFile(skill: ISkillItem): Promise<void> {
    const url = this.getDownloadUrl(skill.serverRelativeUrl);
    const a = document.createElement("a");
    a.href = url;
    a.download = skill.fileName;
    a.target = "_blank";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  /**
   * Fetches the raw text content of a skill file.
   */
  public async getFileContent(serverRelativeUrl: string): Promise<string> {
    const url = `${this.siteUrl}/_api/web/getfilebyserverrelativeurl('${encodeURIComponent(serverRelativeUrl)}')/$value`;
    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      throw new Error(`Failed to fetch file content: ${response.statusText}`);
    }
    return response.text();
  }

  private mapToSkillItem(item: ISpFileItem): ISkillItem {
    const fileName = item.FileLeafRef;
    const name = fileName.replace(/\.md$/i, "");

    // Derive category from parent folder name (last segment of FileDirRef)
    const dirParts = (item.FileDirRef || "").split("/");
    const libraryIndex = dirParts.findIndex(
      (p) => p.toLowerCase() === this.libraryName.toLowerCase()
    );
    const categoryParts =
      libraryIndex >= 0 ? dirParts.slice(libraryIndex + 1) : [];
    const category =
      categoryParts.length > 0 ? categoryParts.join(" / ") : "General";

    return {
      id: String(item.Id),
      name,
      fileName,
      downloadUrl: this.getDownloadUrl(item.FileRef),
      size: parseInt(item.File?.Length || "0", 10),
      lastModified: item.File?.TimeLastModified || "",
      author: item.File?.Author?.Title || "",
      description: item.OData__ExtendedDescription || "",
      category,
      driveItemId: item.UniqueId,
      serverRelativeUrl: item.FileRef,
    };
  }
}
