declare interface ISkillsRepositoryWebPartStrings {
  PropertyPaneDescription: string;
  SourceGroupName: string;
  DestinationGroupName: string;
  SkillsSiteUrlLabel: string;
  SkillsSiteUrlDesc: string;
  SkillsLibraryNameLabel: string;
  SkillsLibraryNameDesc: string;
  CoworkPathLabel: string;
  CoworkPathDesc: string;
  CoworkSubfolderLabel: string;
  CoworkSubfolderDesc: string;
  LocalPathLabel: string;
  LocalPathDesc: string;

  // UI strings
  AppTitle: string;
  SearchPlaceholder: string;
  AllCategories: string;
  CopyToCoworkLabel: string;
  DownloadLocalLabel: string;
  CopyingLabel: string;
  DownloadingLabel: string;
  CopySuccessMessage: string;
  DownloadSuccessMessage: string;
  LoadingLabel: string;
  NoSkillsFound: string;
  ErrorLoadingSkills: string;
  LastModifiedLabel: string;
  AuthorLabel: string;
  SizeLabel: string;
  CategoryLabel: string;
  SkillsCountLabel: string;
  CoworkPathHintLabel: string;
  LocalPathHintLabel: string;
  RefreshLabel: string;
}

declare module "SkillsRepositoryWebPartStrings" {
  const strings: ISkillsRepositoryWebPartStrings;
  export = strings;
}
