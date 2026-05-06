import * as React from "react";
import {
  ActionButton,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  Icon,
  Panel,
  PanelType,
  PrimaryButton,
  SearchBox,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import * as strings from "SkillsRepositoryWebPartStrings";
import { ISkillsRepositoryProps } from "./ISkillsRepositoryProps";
import {
  ISkillItem,
  ISkillsRepositoryState,
} from "./ISkillsRepositoryState";
import { SkillsService } from "../services/SkillsService";
import { OneDriveService } from "../services/OneDriveService";
import {
  IPublishTargetSite,
  SharePointPublishService,
} from "../services/SharePointPublishService";
import styles from "./SkillsRepository.module.css";

interface IPublishSiteTreeNode {
  title: string;
  pathSegments: string[];
  site?: IPublishTargetSite;
  children: IPublishSiteTreeNode[];
}

export default class SkillsRepository extends React.Component<
  ISkillsRepositoryProps,
  ISkillsRepositoryState
> {
  private skillsService: SkillsService;
  private oneDriveService: OneDriveService;
  private publishService: SharePointPublishService;

  constructor(props: ISkillsRepositoryProps) {
    super(props);
    this.state = {
      skills: [],
      filteredSkills: [],
      searchText: "",
      isLoading: true,
      errorMessage: "",
      operationInProgress: {},
      operationResult: {},
      selectedCategory: "",
      categories: [],
      isPublishPanelOpen: false,
      activePublishSkillId: undefined,
      publishTargetSites: [],
      isPublishTargetsLoading: false,
      publishTargetsError: "",
      publishTargetPath: ['sites'],
      publishTargetSearchText: "",
    };
    this.skillsService = new SkillsService(
      props.context,
      props.skillsSiteUrl,
      props.skillsLibraryName
    );
    this.oneDriveService = new OneDriveService(props.graphClient);
    this.publishService = new SharePointPublishService(props.context);
  }

  public async componentDidMount(): Promise<void> {
    await this.loadSkills();
  }

  public async componentDidUpdate(
    prevProps: ISkillsRepositoryProps
  ): Promise<void> {
    if (
      prevProps.skillsSiteUrl !== this.props.skillsSiteUrl ||
      prevProps.skillsLibraryName !== this.props.skillsLibraryName
    ) {
      this.skillsService = new SkillsService(
        this.props.context,
        this.props.skillsSiteUrl,
        this.props.skillsLibraryName
      );
      await this.loadSkills();
    }
  }

  // ── Data Loading ───────────────────────────────────────────────────

  private async loadSkills(): Promise<void> {
    this.setState({ isLoading: true, errorMessage: "" });
    try {
      const skills = await this.skillsService.getSkills();
      const categories = Array.from(
        new Set(skills.map((s) => s.category))
      ).sort();
      this.setState({
        skills,
        filteredSkills: skills,
        categories,
        isLoading: false,
      });
    } catch (err) {
      this.setState({
        isLoading: false,
        errorMessage: strings.ErrorLoadingSkills,
      });
      console.error("[ASORepo] Failed to load skills:", err);
    }
  }

  private async loadPublishTargets(forceRefresh: boolean = false): Promise<void> {
    if (this.state.isPublishTargetsLoading) {
      return;
    }

    if (!forceRefresh && this.state.publishTargetSites.length > 0) {
      return;
    }

    this.setState({
      isPublishTargetsLoading: true,
      publishTargetsError: "",
    });

    try {
      const publishTargetSites = await this.publishService.getAgentAssetSites();
      this.setState((prev) => ({
        publishTargetSites,
        isPublishTargetsLoading: false,
        publishTargetsError: "",
        publishTargetPath: this.coercePublishPath(
          prev.publishTargetPath,
          publishTargetSites
        ),
      }));
    } catch (err) {
      console.error("[ASORepo] Failed to load publish targets:", err);
      this.setState({
        isPublishTargetsLoading: false,
        publishTargetsError: strings.PublishPanelError,
      });
    }
  }

  // ── Filtering ──────────────────────────────────────────────────────

  private applyFilters(
    searchText: string,
    selectedCategory: string,
    skills: ISkillItem[]
  ): ISkillItem[] {
    return skills.filter((s) => {
      const matchesSearch =
        !searchText ||
        s.name.toLowerCase().includes(searchText.toLowerCase()) ||
        s.description.toLowerCase().includes(searchText.toLowerCase()) ||
        s.category.toLowerCase().includes(searchText.toLowerCase());
      const matchesCategory =
        !selectedCategory || s.category === selectedCategory;
      return matchesSearch && matchesCategory;
    });
  }

  private onSearchChange = (
    _ev: React.ChangeEvent<HTMLInputElement> | undefined,
    newValue?: string
  ): void => {
    const searchText = newValue || "";
    this.setState((prev) => ({
      searchText,
      filteredSkills: this.applyFilters(
        searchText,
        prev.selectedCategory,
        prev.skills
      ),
    }));
  };

  private onCategoryChange = (
    _ev: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    const selectedCategory = option?.key === "" ? "" : String(option?.key || "");
    this.setState((prev) => ({
      selectedCategory,
      filteredSkills: this.applyFilters(
        prev.searchText,
        selectedCategory,
        prev.skills
      ),
    }));
  };

  private onPublishSearchChange = (
    _ev: React.ChangeEvent<HTMLInputElement> | undefined,
    newValue?: string
  ): void => {
    this.setState({ publishTargetSearchText: newValue || "" });
  };

  // ── Actions ────────────────────────────────────────────────────────

  private copyToCowork = async (skill: ISkillItem): Promise<void> => {
    this.setOperationInProgress(skill.id, "copy");
    try {
      const content = await this.skillsService.getFileContent(
        skill.serverRelativeUrl
      );
      await this.oneDriveService.copySkillToOneDrive(
        skill,
        content,
        this.props.coworkOneDrivePath,
        this.props.coworkSubfolderName
      );
      this.setOperationResult(skill.id, true, strings.CopySuccessMessage);
    } catch (err) {
      const message =
        err instanceof Error ? err.message : "An error occurred.";
      this.setOperationResult(skill.id, false, message);
      console.error(`[ASORepo] Copy failed for ${skill.name}:`, err);
    } finally {
      this.setOperationInProgress(skill.id, null);
    }
  };

  private downloadLocal = async (skill: ISkillItem): Promise<void> => {
    this.setOperationInProgress(skill.id, "download");
    try {
      await this.skillsService.downloadFile(skill);
      this.setOperationResult(
        skill.id,
        true,
        `${strings.DownloadSuccessMessage} ${this.resolveLocalPath(
          this.props.userDisplayName
        )}`
      );
    } catch (err) {
      const message =
        err instanceof Error ? err.message : "An error occurred.";
      this.setOperationResult(skill.id, false, message);
    } finally {
      this.setOperationInProgress(skill.id, null);
    }
  };

  private openPublishPanel = async (skill: ISkillItem): Promise<void> => {
    this.setState({
      isPublishPanelOpen: true,
      activePublishSkillId: skill.id,
      publishTargetSearchText: "",
    });

    await this.loadPublishTargets();
  };

  private closePublishPanel = (): void => {
    this.setState({
      isPublishPanelOpen: false,
      activePublishSkillId: undefined,
      publishTargetSearchText: "",
      publishTargetsError: "",
    });
  };

  private publishToSharePoint = async (
    skill: ISkillItem,
    targetSite: IPublishTargetSite
  ): Promise<void> => {
    this.setOperationInProgress(skill.id, "publish");
    this.closePublishPanel();

    try {
      const content = await this.skillsService.getFileContent(
        skill.serverRelativeUrl
      );
      await this.publishService.publishSkillToSite(skill, content, targetSite.url);
      this.setOperationResult(
        skill.id,
        true,
        `${strings.PublishSuccessMessage} ${targetSite.title}`
      );
    } catch (err) {
      const message =
        err instanceof Error ? err.message : "An error occurred.";
      this.setOperationResult(skill.id, false, message);
      console.error(`[ASORepo] Publish failed for ${skill.name}:`, err);
    } finally {
      this.setOperationInProgress(skill.id, null);
    }
  };

  private setOperationInProgress(
    skillId: string,
    op: "copy" | "download" | "publish" | null
  ): void {
    this.setState((prev) => ({
      operationInProgress: { ...prev.operationInProgress, [skillId]: op },
    }));
  }

  private resolveLocalPath(displayName: string): string {
    const userSegment = displayName
      ? displayName.split(" ")[0].toLowerCase()
      : "username";
    return this.props.localMachinePath.replace("{username}", userSegment);
  }

  private setOperationResult(
    skillId: string,
    success: boolean,
    message: string
  ): void {
    this.setState((prev) => ({
      operationResult: {
        ...prev.operationResult,
        [skillId]: { success, message },
      },
    }));

    setTimeout(() => {
      this.setState((prev) => {
        const updated = { ...prev.operationResult };
        delete updated[skillId];
        return { operationResult: updated };
      });
    }, 6000);
  }

  private buildPublishSiteTree(
    sites: IPublishTargetSite[] = this.state.publishTargetSites
  ): IPublishSiteTreeNode {
    const root: IPublishSiteTreeNode = {
      title: strings.PublishRootLabel,
      pathSegments: [],
      children: [],
    };

    sites.forEach((site) => {
      if (site.pathSegments.length === 0) {
        root.site = site;
        root.title = site.title;
        return;
      }

      let currentNode = root;

      site.pathSegments.forEach((segment, index) => {
        const pathSegments = site.pathSegments.slice(0, index + 1);
        let childNode = currentNode.children.find(
          (child) => child.pathSegments.join("/") === pathSegments.join("/")
        );

        if (!childNode) {
          childNode = {
            title: this.formatPublishNodeLabel(segment),
            pathSegments,
            children: [],
          };
          currentNode.children.push(childNode);
        }

        if (index === site.pathSegments.length - 1) {
          childNode.site = site;
          childNode.title = site.title || childNode.title;
        }

        currentNode = childNode;
      });
    });

    return root;
  }

  private coercePublishPath(
    publishTargetPath: string[],
    sites: IPublishTargetSite[]
  ): string[] {
    const root = this.buildPublishSiteTree(sites);
    const resolvedPath: string[] = [];
    let currentNode = root;

    for (const segment of publishTargetPath) {
      const nextPath = [...resolvedPath, segment].join("/");
      const nextNode = currentNode.children.find(
        (child) => child.pathSegments.join("/") === nextPath
      );

      if (!nextNode) {
        break;
      }

      resolvedPath.push(segment);
      currentNode = nextNode;
    }

    return resolvedPath;
  }

  private getPublishNode(
    root: IPublishSiteTreeNode,
    publishTargetPath: string[]
  ): IPublishSiteTreeNode {
    let currentNode = root;
    const resolvedPath: string[] = [];

    for (const segment of publishTargetPath) {
      const nextPath = [...resolvedPath, segment].join("/");
      const nextNode = currentNode.children.find(
        (child) => child.pathSegments.join("/") === nextPath
      );

      if (!nextNode) {
        break;
      }

      resolvedPath.push(segment);
      currentNode = nextNode;
    }

    return currentNode;
  }

  private getPublishSearchResults(root: IPublishSiteTreeNode): IPublishSiteTreeNode[] {
    const allSites: IPublishSiteTreeNode[] = [];
    const queue = [...root.children];

    if (root.site) {
      allSites.push(root);
    }

    while (queue.length > 0) {
      const node = queue.shift();
      if (!node) {
        continue;
      }

      if (node.site) {
        allSites.push(node);
      }

      queue.push(...node.children);
    }

    const search = this.state.publishTargetSearchText.trim().toLowerCase();
    return allSites
      .filter((node) => {
        const site = node.site;
        return Boolean(
          site &&
            (!search ||
              node.title.toLowerCase().includes(search) ||
              site.url.toLowerCase().includes(search))
        );
      })
      .sort((left, right) => left.title.localeCompare(right.title));
  }

  private formatPublishNodeLabel(segment: string): string {
    return decodeURIComponent(segment).replace(/[-_]/g, " ");
  }

  private getNodeUrl(node: IPublishSiteTreeNode): string {
    if (node.site) {
      return node.site.url;
    }

    const tenantOrigin = new URL(this.props.context.pageContext.web.absoluteUrl).origin;
    return node.pathSegments.length > 0
      ? `${tenantOrigin}/${node.pathSegments.join("/")}`
      : tenantOrigin;
  }

  private renderPublishPanel(activeSkill?: ISkillItem): React.ReactNode {
    const {
      isPublishPanelOpen,
      isPublishTargetsLoading,
      publishTargetsError,
      publishTargetPath,
      publishTargetSites,
      publishTargetSearchText,
      operationInProgress,
    } = this.state;
    const root = this.buildPublishSiteTree();
    const currentNode = this.getPublishNode(root, publishTargetPath);
    const searchResults = this.getPublishSearchResults(root);
    const hasSearch = publishTargetSearchText.trim().length > 0;
    const browseNodes = [...currentNode.children].sort((left, right) =>
      left.title.localeCompare(right.title)
    );
    const visibleNodes = hasSearch ? searchResults : browseNodes;
    const activeSkillBusy = activeSkill
      ? operationInProgress[activeSkill.id] !== null
      : true;

    //TODO: work out the point of this agentskillbusy


    const breadcrumbNodes = publishTargetPath.reduce<IPublishSiteTreeNode[]>(
      (items, _segment, index) => {
        items.push(this.getPublishNode(root, publishTargetPath.slice(0, index + 1)));
        return items;
      },
      []
    );

    return (
      <Panel
        isOpen={isPublishPanelOpen}
        type={PanelType.medium}
        headerText={strings.PublishPanelTitle}
        closeButtonAriaLabel={strings.PublishPanelTitle}
        isLightDismiss={true}
        onDismiss={this.closePublishPanel}
      >
        <div className={styles.publishPanel}>
          {activeSkill && (
            <div className={styles.publishPanelSummary}>
              <span className={styles.publishPanelLabel}>{skillLabel(activeSkill)}</span>
              <span className={styles.publishPanelValue}>{activeSkill.fileName}</span>
            </div>
          )}

          <p className={styles.publishPanelDescription}>
            {strings.PublishPanelDescription}
          </p>

          <div className={styles.publishPanelToolbar}>
            <SearchBox
              className={styles.publishSearchBox}
              placeholder={strings.PublishPanelSearchPlaceholder}
              value={publishTargetSearchText}
              onChange={this.onPublishSearchChange}
              onClear={() => this.onPublishSearchChange(undefined, "")}
            />
            <DefaultButton
              iconProps={{ iconName: "Refresh" }}
              onClick={() => this.loadPublishTargets(true)}
              disabled={isPublishTargetsLoading}
            >
              {strings.PublishTargetRefreshLabel}
            </DefaultButton>
          </div>

          {!hasSearch && (
            <div className={styles.publishBreadcrumbWrap}>
              <span className={styles.publishBreadcrumbLabel}>
                {strings.PublishPanelCurrentLocationLabel}
              </span>
              <div className={styles.publishBreadcrumb}>
                <ActionButton
                  onClick={() => this.setState({ publishTargetPath: [] })}
                >
                  {strings.PublishRootLabel}
                </ActionButton>
                {breadcrumbNodes.map((node) => (
                  <ActionButton
                    key={node.pathSegments.join("/") || "root"}
                    onClick={() =>
                      this.setState({ publishTargetPath: node.pathSegments })
                    }
                  >
                    {node.title}
                  </ActionButton>
                ))}
              </div>
            </div>
          )}

          {!hasSearch && currentNode.site && activeSkill && (
            <div className={styles.publishCurrentSiteCard}>
              <div className={styles.publishSiteHeader}>
                <span className={styles.publishSiteName}>{currentNode.title}</span>
                <span className={styles.publishSiteUrl}>{currentNode.site.url}</span>
              </div>
              
              <PrimaryButton
                onClick={() => this.publishToSharePoint(activeSkill, currentNode.site!)}
                disabled={activeSkillBusy}
              >
                {
                
                /*activeSkillBusy
                  ? strings.PublishingLabel
                  : strings.PublishHereLabel*/
                  strings.PublishHereLabel}
              </PrimaryButton>
            </div>
          )}

          {isPublishTargetsLoading && (
            <div className={styles.stateContainer}>
              <Spinner
                size={SpinnerSize.large}
                label={strings.PublishPanelLoadingLabel}
              />
            </div>
          )}

          {!isPublishTargetsLoading && publishTargetsError && (
            <div className={styles.stateContainer}>
              <Icon iconName="ErrorBadge" className={styles.stateIcon} />
              <p className={styles.stateMessage}>{publishTargetsError}</p>
            </div>
          )}

          {!isPublishTargetsLoading &&
            !publishTargetsError &&
            publishTargetSites.length === 0 && (
              <div className={styles.stateContainer}>
                <Icon iconName="World" className={styles.stateIcon} />
                <p className={styles.stateMessage}>{strings.PublishPanelNoSites}</p>
              </div>
            )}

          {!isPublishTargetsLoading &&
            !publishTargetsError &&
            publishTargetSites.length > 0 &&
            visibleNodes.length === 0 && (
              <div className={styles.stateContainer}>
                <Icon iconName="Search" className={styles.stateIcon} />
                <p className={styles.stateMessage}>{strings.PublishPanelNoResults}</p>
              </div>
            )}

          {!isPublishTargetsLoading &&
            !publishTargetsError &&
            visibleNodes.length > 0 && (
              <div className={styles.publishSiteList}>
                {visibleNodes.map((node) => {
                  const site = node.site;
                  const canPublish = Boolean(site && activeSkill);
                  const isBusy = activeSkillBusy;
                  const key = site?.url || node.pathSegments.join("/");

                  return (
                    <div className={styles.publishSiteCard} key={key}>
                      <div className={styles.publishSiteHeader}>
                        <span className={styles.publishSiteName}>{node.title}</span>
                        <span className={styles.publishSiteUrl}>
                          {this.getNodeUrl(node)}
                        </span>
                      </div>
                      <div className={styles.publishSiteActions}>
                        {!hasSearch && node.children.length > 0 && (
                          <DefaultButton
                            onClick={() =>
                              this.setState({ publishTargetPath: node.pathSegments })
                            }
                          >
                            {strings.BrowseIntoLabel}
                          </DefaultButton>
                        )}
                        {hasSearch && (
                          <DefaultButton
                            onClick={() =>
                              this.setState({
                                publishTargetPath: node.pathSegments,
                                publishTargetSearchText: "",
                              })
                            }
                          >
                            {strings.BrowseIntoLabel}
                          </DefaultButton>
                        )}
                        {canPublish && site && (
                          
                          <PrimaryButton
                            onClick={() => this.publishToSharePoint(activeSkill!, site)}
                            disabled={
                              //isBusy
                              false
                              }
                          >
                            {
                            /*isBusy
                              ? strings.PublishingLabel
                              : strings.PublishHereLabel*/
                              strings.PublishHereLabel}
                          </PrimaryButton>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
        </div>
      </Panel>
    );
  }

  // ── Render ─────────────────────────────────────────────────────────

  public render(): React.ReactElement {
    const {
      isLoading,
      errorMessage,
      filteredSkills,
      searchText,
      selectedCategory,
      categories,
      skills,
      operationInProgress,
      operationResult,
      activePublishSkillId,
    } = this.state;
    const { coworkOneDrivePath, userDisplayName, isDarkTheme } = this.props;

    const activePublishSkill =
      skills.find((skill) => skill.id === activePublishSkillId) || undefined;

    const containerClass = [
      styles.container,
      isDarkTheme ? styles.isDarkTheme : "",
    ]
      .filter(Boolean)
      .join(" ");

    const categoryOptions: IDropdownOption[] = [
      { key: "", text: strings.AllCategories },
      ...categories.map((c) => ({ key: c, text: c })),
    ];

    return (
      <>
        <div className={containerClass}>
          <div className={styles.header}>
            <h2 className={styles.title}>{strings.AppTitle}</h2>
            <div className={styles.headerActions}>
              <ActionButton
                iconProps={{ iconName: "Refresh" }}
                onClick={() => this.loadSkills()}
                disabled={isLoading}
              >
                {strings.RefreshLabel}
              </ActionButton>
            </div>
          </div>

          <div className={styles.pathInfoBanner}>
            <div className={styles.pathInfoRow}>
              <span className={styles.pathInfoLabel}>
                <Icon iconName="OneDrive" /> {strings.CoworkPathHintLabel}:
              </span>
              <span className={styles.pathInfoValue}>
                OneDrive › {coworkOneDrivePath}
              </span>
            </div>
            <div className={styles.pathInfoRow}>
              <span className={styles.pathInfoLabel}>
                <Icon iconName="ThisPC" /> {strings.LocalPathHintLabel}:
              </span>
              <span className={styles.pathInfoValue}>
                {this.resolveLocalPath(userDisplayName)}
              </span>
            </div>
          </div>

          <div className={styles.toolbar}>
            <SearchBox
              className={styles.searchBox}
              placeholder={strings.SearchPlaceholder}
              value={searchText}
              onChange={this.onSearchChange}
              onClear={() => this.onSearchChange(undefined, "")}
            />
            <Dropdown
              className={styles.categoryDropdown}
              options={categoryOptions}
              selectedKey={selectedCategory}
              onChange={this.onCategoryChange}
            />
            {!isLoading && (
              <span className={styles.skillCount}>
                {filteredSkills.length} / {skills.length} {strings.SkillsCountLabel}
              </span>
            )}
          </div>

          {isLoading && (
            <div className={styles.stateContainer}>
              <Spinner size={SpinnerSize.large} label={strings.LoadingLabel} />
            </div>
          )}

          {!isLoading && errorMessage && (
            <div className={styles.stateContainer}>
              <Icon iconName="ErrorBadge" className={styles.stateIcon} />
              <p className={styles.stateMessage}>{errorMessage}</p>
              <DefaultButton
                iconProps={{ iconName: "Refresh" }}
                onClick={() => this.loadSkills()}
              >
                {strings.RefreshLabel}
              </DefaultButton>
            </div>
          )}

          {!isLoading && !errorMessage && filteredSkills.length === 0 && (
            <div className={styles.stateContainer}>
              <Icon iconName="Search" className={styles.stateIcon} />
              <p className={styles.stateMessage}>{strings.NoSkillsFound}</p>
            </div>
          )}

          {!isLoading && !errorMessage && filteredSkills.length > 0 && (
            <div className={styles.skillsGrid}>
              {filteredSkills.map((skill) => (
                <SkillCard
                  key={skill.id}
                  skill={skill}
                  operationInProgress={operationInProgress[skill.id] || null}
                  operationResult={operationResult[skill.id]}
                  onCopyToCowork={this.copyToCowork}
                  onDownloadLocal={this.downloadLocal}
                  onPublishToSharePoint={this.openPublishPanel}
                />
              ))}
            </div>
          )}
        </div>
        {this.renderPublishPanel(activePublishSkill)}
      </>
    );
  }
}

interface ISkillCardProps {
  skill: ISkillItem;
  operationInProgress: "copy" | "download" | "publish" | null;
  operationResult?: { success: boolean; message: string };
  onCopyToCowork: (skill: ISkillItem) => Promise<void>;
  onDownloadLocal: (skill: ISkillItem) => Promise<void>;
  onPublishToSharePoint: (skill: ISkillItem) => Promise<void>;
}

class SkillCard extends React.PureComponent<ISkillCardProps> {
  private formatSize(bytes: number): string {
    if (!bytes) return "";
    if (bytes < 1024) return `${bytes} B`;
    return `${(bytes / 1024).toFixed(1)} KB`;
  }

  private formatDate(iso: string): string {
    if (!iso) return "";
    return new Date(iso).toLocaleDateString(undefined, {
      year: "numeric",
      month: "short",
      day: "numeric",
    });
  }

  public render(): React.ReactElement {
    const { skill, operationInProgress, operationResult } = this.props;
    const isBusy = operationInProgress !== null;

    return (
      <div className={styles.skillCard}>
        <div className={styles.skillCardHeader}>
          <Icon iconName="TextDocument" className={styles.skillIcon} />
          <div>
            <p className={styles.skillName}>{skill.name}</p>
            {skill.description && (
              <p className={styles.skillDescription}>{skill.description}</p>
            )}
          </div>
        </div>

        <span className={styles.categoryBadge}>{skill.category}</span>

        <div className={styles.skillMeta}>
          {skill.author && (
            <span className={styles.skillMetaItem}>
              <Icon iconName="Contact" />
              {skill.author}
            </span>
          )}
          {skill.lastModified && (
            <span className={styles.skillMetaItem}>
              <Icon iconName="Calendar" />
              {this.formatDate(skill.lastModified)}
            </span>
          )}
          {skill.size > 0 && (
            <span className={styles.skillMetaItem}>
              <Icon iconName="PageData" />
              {this.formatSize(skill.size)}
            </span>
          )}
        </div>

        <div className={styles.cardActions}>
          <ActionButton
            iconProps={{
              iconName: operationInProgress === "copy" ? "Sync" : "OneDrive",
            }}
            disabled={isBusy}
            onClick={() => this.props.onCopyToCowork(skill)}
          >
            {operationInProgress === "copy"
              ? strings.CopyingLabel
              : strings.CopyToCoworkLabel}
          </ActionButton>
          <ActionButton
            iconProps={{
              iconName: operationInProgress === "download" ? "Sync" : "Download",
            }}
            disabled={isBusy}
            onClick={() => this.props.onDownloadLocal(skill)}
          >
            {operationInProgress === "download"
              ? strings.DownloadingLabel
              : strings.DownloadLocalLabel}
          </ActionButton>
          <ActionButton
            iconProps={{
              iconName: operationInProgress === "publish" ? "Sync" : "Share",
            }}
            disabled={isBusy}
            onClick={() => this.props.onPublishToSharePoint(skill)}
          >
            {operationInProgress === "publish"
              ? strings.PublishingLabel
              : strings.PublishToSharePointLabel}
          </ActionButton>
        </div>

        {operationResult && (
          <div
            className={[
              styles.operationResult,
              operationResult.success
                ? styles.operationResultSuccess
                : styles.operationResultError,
            ].join(" ")}
          >
            <Icon
              iconName={operationResult.success ? "CheckMark" : "ErrorBadge"}
            />
            {operationResult.message}
          </div>
        )}
      </div>
    );
  }
}

function skillLabel(skill: ISkillItem): string {
  return `${strings.PublishToSharePointLabel}: ${skill.name}`;
}
