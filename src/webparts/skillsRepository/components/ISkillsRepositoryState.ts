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
  operationInProgress: { [skillId: string]: "copy" | "download" | null };
  /** Map of skill ID → last operation result message */
  operationResult: {
    [skillId: string]: { success: boolean; message: string };
  };
  /** Active category filter */
  selectedCategory: string;
  /** Sorted list of unique categories */
  categories: string[];
}
