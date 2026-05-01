import { SPHttpClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISkillItem } from "../components/ISkillsRepositoryState";

export interface IPublishTargetSite {
  title: string;
  url: string;
  pathSegments: string[];
}

interface ISearchCell {
  Key: string;
  Value: string;
}

interface ISearchRow {
  Cells: ISearchCell[];
}

interface ISearchResponse {
  PrimaryQueryResult?: {
    RelevantResults?: {
      Table?: {
        Rows?: ISearchRow[];
      };
    };
  };
}

interface IFolderInfo {
  Exists: boolean;
  ServerRelativeUrl?: string;
}

/**
 * Service for discovering SharePoint sites with the Agent Assets library
 * and publishing markdown skills into that library.
 */
export class SharePointPublishService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  public async getAgentAssetSites(): Promise<IPublishTargetSite[]> {
    const queryText = `* Title:"Agent Assets" AND contentclass:STS_List_DocumentLibrary`;
    const selectProperties = ["Title", "Path", "SPSiteUrl", "SiteName"].join(",");
    const searchUrl =
      `${this.context.pageContext.web.absoluteUrl}/_api/search/query` +
      `?querytext='${encodeURIComponent(queryText)}'` +
      `&selectproperties='${encodeURIComponent(selectProperties)}'` +
      "&trimduplicates=false" +
      "&rowlimit=500";

    const response = await this.context.spHttpClient.get(
      searchUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
        },
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to discover publish sites: ${response.statusText}`);
    }

    const data = (await response.json()) as ISearchResponse;
    const rows = data.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];
    const tenantOrigin = new URL(this.context.pageContext.web.absoluteUrl).origin;
    const seen = new Set<string>();

    return rows
      .map((row) => this.mapSearchRow(row, tenantOrigin))
      .filter((site): site is IPublishTargetSite => Boolean(site))
      .filter((site) => {
        const key = site.url.toLowerCase();
        if (seen.has(key)) {
          return false;
        }
        seen.add(key);
        return true;
      })
      .sort((left, right) => left.url.localeCompare(right.url));
  }

  public async publishSkillToSite(
    skill: ISkillItem,
    fileContent: string,
    targetSiteUrl: string
  ): Promise<void> {
    const siteUrl = targetSiteUrl.replace(/\/$/, "");
    const libraryRoot = await this.getAgentAssetsRoot(siteUrl);
    const destinationFolder = await this.ensureCategoryFolder(
      siteUrl,
      libraryRoot,
      skill.category
    );
    await this.uploadFile(siteUrl, destinationFolder, skill.fileName, fileContent);
  }

  private mapSearchRow(
    row: ISearchRow,
    tenantOrigin: string
  ): IPublishTargetSite | undefined {
    const values = new Map(
      (row.Cells || []).map((cell) => [cell.Key.toLowerCase(), cell.Value])
    );
    const siteUrl = (values.get("spsiteurl") || "").replace(/\/$/, "");

    if (!siteUrl || !siteUrl.startsWith(tenantOrigin)) {
      return undefined;
    }

    const relativePath = siteUrl.slice(tenantOrigin.length).replace(/^\//, "");
    return {
      title: values.get("sitename") || values.get("title") || this.getFallbackTitle(siteUrl),
      url: siteUrl,
      pathSegments: relativePath ? relativePath.split("/").filter(Boolean) : [],
    };
  }

  private getFallbackTitle(siteUrl: string): string {
    const segments = new URL(siteUrl).pathname.split("/").filter(Boolean);
    return segments.length > 0 ? segments[segments.length - 1] : "Root site";
  }

  private async getAgentAssetsRoot(siteUrl: string): Promise<string> {
    const url =
      `${siteUrl}/_api/web/lists/getbytitle('Agent Assets')` +
      "?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder";
    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
        },
      }
    );

    if (!response.ok) {
      throw new Error(`Agent Assets is not available on ${siteUrl}.`);
    }

    const data = (await response.json()) as {
      RootFolder?: { ServerRelativeUrl?: string };
    };
    const root = data.RootFolder?.ServerRelativeUrl;

    if (!root) {
      throw new Error(`Agent Assets root folder could not be resolved for ${siteUrl}.`);
    }

    return root;
  }

  private async ensureCategoryFolder(
    siteUrl: string,
    libraryRoot: string,
    category: string
  ): Promise<string> {
    if (!category || category === "General") {
      return libraryRoot;
    }

    const categorySegments = category.split("/").map((segment) => segment.trim()).filter(Boolean);
    let currentFolder = libraryRoot;

    for (const segment of categorySegments) {
      const nextFolder = `${currentFolder}/${segment}`;
      const folder = await this.getFolder(siteUrl, nextFolder);
      if (!folder.Exists) {
        await this.createFolder(siteUrl, currentFolder, segment);
      }
      currentFolder = nextFolder;
    }

    return currentFolder;
  }

  private async getFolder(siteUrl: string, folderServerRelativeUrl: string): Promise<IFolderInfo> {
    const url =
      `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${this.escapeODataString(folderServerRelativeUrl)}')` +
      "?$select=Exists,ServerRelativeUrl";
    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
        },
      }
    );

    if (response.status === 404) {
      return { Exists: false };
    }

    if (!response.ok) {
      throw new Error(`Failed to inspect folder ${folderServerRelativeUrl}.`);
    }

    return (await response.json()) as IFolderInfo;
  }

  private async createFolder(
    siteUrl: string,
    parentFolderServerRelativeUrl: string,
    folderName: string
  ): Promise<void> {
    const url = `${siteUrl}/_api/web/folders`;
    const body = JSON.stringify({
      ServerRelativeUrl: `${parentFolderServerRelativeUrl}/${folderName}`,
    });

    const response = await this.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
        },
        body,
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to create folder \"${folderName}\": ${errorText}`);
    }
  }

  private async uploadFile(
    siteUrl: string,
    folderServerRelativeUrl: string,
    fileName: string,
    content: string
  ): Promise<void> {
    const uploadUrl =
      `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${this.escapeODataString(folderServerRelativeUrl)}')` +
      `/Files/AddUsingPath(decodedurl='${this.escapeODataString(fileName)}',overwrite=true)`;

    const response = await this.context.spHttpClient.post(
      uploadUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "text/markdown;charset=utf-8",
        },
        body: content,
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to publish \"${fileName}\": ${errorText}`);
    }
  }

  private escapeODataString(value: string): string {
    return value.replace(/'/g, "''");
  }
}