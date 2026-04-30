define([], function () {
  return {
    // Property pane
    PropertyPaneDescription:
      "Configure the ASORepo \u2014 AI Skills Org Repository settings for this web part instance.",
    SourceGroupName: "Source (SharePoint Skills Library)",
    DestinationGroupName: "Destination Paths",
    SkillsSiteUrlLabel: "Skills Site URL",
    SkillsSiteUrlDesc:
      "Full URL of the SharePoint site hosting the Skills library.",
    SkillsLibraryNameLabel: "Skills Library Name",
    SkillsLibraryNameDesc:
      "Name of the document library containing skill .md files.",
    CoworkPathLabel: "OneDrive Cowork Path",
    CoworkPathDesc:
      "OneDrive folder path where skills are copied for Cowork / VS Code Copilot agent use.",
    CoworkSubfolderLabel: "OneDrive Cowork Subfolder (Optional)",
    CoworkSubfolderDesc:
      "Optional subfolder inside the OneDrive Cowork path. Leave empty to copy directly to the root Cowork path.",
    LocalPathLabel: "Local Machine Path",
    LocalPathDesc:
      "Local path shown as guidance after downloading a skill file.",

    // UI
    AppTitle: "ASORepo \u2014 AI Skills Org Repository",
    SearchPlaceholder: "Search skills\u2026",
    AllCategories: "All Categories",
    CopyToCoworkLabel: "Copy to Cowork",
    DownloadLocalLabel: "Download",
    CopyingLabel: "Copying\u2026",
    DownloadingLabel: "Downloading\u2026",
    CopySuccessMessage: "Copied to your OneDrive Cowork folder.",
    DownloadSuccessMessage:
      "Downloaded. Place the file in your local skills folder.",
    LoadingLabel: "Loading skills\u2026",
    NoSkillsFound: "No skills match your search.",
    ErrorLoadingSkills:
      "Could not load skills. Check the site URL and library name in the web part settings.",
    LastModifiedLabel: "Last modified",
    AuthorLabel: "Author",
    SizeLabel: "Size",
    CategoryLabel: "Category",
    SkillsCountLabel: "skills",
    CoworkPathHintLabel: "OneDrive Cowork path",
    LocalPathHintLabel: "Local path",
    RefreshLabel: "Refresh",
  };
});
