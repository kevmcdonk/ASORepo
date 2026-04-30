import * as React from "react";
import {
  ActionButton,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  Icon,
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
import styles from "./SkillsRepository.module.css";

export default class SkillsRepository extends React.Component<
  ISkillsRepositoryProps,
  ISkillsRepositoryState
> {
  private skillsService: SkillsService;
  private oneDriveService: OneDriveService;

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
    };
    this.skillsService = new SkillsService(
      props.context,
      props.skillsSiteUrl,
      props.skillsLibraryName
    );
    this.oneDriveService = new OneDriveService(props.graphClient);
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

  private setOperationInProgress(
    skillId: string,
    op: "copy" | "download" | null
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
    // Auto-clear result after 6 seconds
    setTimeout(() => {
      this.setState((prev) => {
        const updated = { ...prev.operationResult };
        delete updated[skillId];
        return { operationResult: updated };
      });
    }, 6000);
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
    } = this.state;
    const { coworkOneDrivePath, userDisplayName, isDarkTheme } = this.props;

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
      <div className={containerClass}>
        {/* Header */}
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

        {/* Path info banner */}
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

        {/* Toolbar */}
        <div className={styles.toolbar}>
          <SearchBox
            className={styles.searchBox}
            placeholder={strings.SearchPlaceholder}
            value={searchText}
            onChange={this.onSearchChange}
            onClear={() =>
              this.onSearchChange(undefined, "")
            }
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

        {/* Body */}
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
              />
            ))}
          </div>
        )}
      </div>
    );
  }
}

// ── SkillCard ──────────────────────────────────────────────────────────

interface ISkillCardProps {
  skill: ISkillItem;
  operationInProgress: "copy" | "download" | null;
  operationResult?: { success: boolean; message: string };
  onCopyToCowork: (skill: ISkillItem) => Promise<void>;
  onDownloadLocal: (skill: ISkillItem) => Promise<void>;
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
        {/* Card header */}
        <div className={styles.skillCardHeader}>
          <Icon iconName="TextDocument" className={styles.skillIcon} />
          <div>
            <p className={styles.skillName}>{skill.name}</p>
            {skill.description && (
              <p className={styles.skillDescription}>{skill.description}</p>
            )}
          </div>
        </div>

        {/* Category badge */}
        <span className={styles.categoryBadge}>{skill.category}</span>

        {/* Meta */}
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

        {/* Actions */}
        <div className={styles.cardActions}>
          <ActionButton
            iconProps={{
              iconName:
                operationInProgress === "copy" ? "Sync" : "OneDrive",
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
              iconName:
                operationInProgress === "download"
                  ? "Sync"
                  : "Download",
            }}
            disabled={isBusy}
            onClick={() => this.props.onDownloadLocal(skill)}
          >
            {operationInProgress === "download"
              ? strings.DownloadingLabel
              : strings.DownloadLocalLabel}
          </ActionButton>
        </div>

        {/* Operation result */}
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
