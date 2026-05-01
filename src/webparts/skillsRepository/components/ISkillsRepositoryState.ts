export interface ISkillItem {
  /** Unique identifier from SharePoint */
  id: string;
  /** Skill name (file name without extension) */
  name: string;
  /** Full file name including extension */
  fileName: string;
  /** URL to download the file directly */
  downloadUrl: string;
  /** File size in bytes */
  size: number;
  /** Last modified date */
  lastModified: string;
  /** Author / creator */
  author: string;
  /** Description from file properties */
  description: string;
  /** Folder/category the skill belongs to */
  category: string;
  /** SharePoint drive item ID for Graph operations */
  driveItemId: string;
  /** SharePoint site-relative URL */
  serverRelativeUrl: string;
}

export interface ISkillsRepositoryState {
  /** All skills loaded from SharePoint */
  skills: ISkillItem[];
  /** Skills visible after search/filter */
  filteredSkills: ISkillItem[];
  /** Current search text */
  searchText: string;
  /** Whether data is being fetched */
  isLoading: boolean;
  /** General error message */
  errorMessage: string;
  /** Map of skill ID → operation in progress */
  operationInProgress: {
    [skillId: string]: "copy" | "download" | "publish" | null;
  };
  /** Map of skill ID → last operation result message */
  operationResult: {
    [skillId: string]: { success: boolean; message: string };
  };
  /** Active category filter */
  selectedCategory: string;
  /** Sorted list of unique categories */
  categories: string[];
  /** Whether the publish panel is open */
  isPublishPanelOpen: boolean;
  /** Skill currently selected for publish */
  activePublishSkillId?: string;
  /** Agent Assets-enabled SharePoint sites */
  publishTargetSites: Array<{
    title: string;
    url: string;
    pathSegments: string[];
  }>;
  /** Whether publish targets are loading */
  isPublishTargetsLoading: boolean;
  /** Publish target discovery error */
  publishTargetsError: string;
  /** Current path in the publish site browser */
  publishTargetPath: string[];
  /** Search text within the publish site browser */
  publishTargetSearchText: string;
}
