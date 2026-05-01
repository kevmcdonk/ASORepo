# ASORepo — AI Skills Org Repository

A SharePoint Framework (SPFx 1.22) web part that provides a central, browsable catalogue of AI agent skill files (`.md`) stored in a SharePoint document library. Users can:

- **Copy to Cowork** — copy a skill directly into their OneDrive in the configured Cowork/Copilot agents folder (synced locally by OneDrive for Business).
- **Download** — trigger a browser download to save the skill file to their local machine for use with local VS Code.

---

## Quick-start

### 1. Install dependencies

```bash
cd ASORepo
npm install
```

### 2. Configure web part settings

The web part stores its settings as editable web part properties. Defaults are already provided in the manifest and web part code.

When you add the web part to a page, open the property pane and set:

| Property | Description | Example |
|---|---|---|
| `skillsSiteUrl` | SharePoint site URL hosting the Skills library | `https://contoso.sharepoint.com/sites/OrgAISkills` |
| `skillsLibraryName` | Document library name | `Skills` |
| `coworkOneDrivePath` | OneDrive folder path for Cowork use | `Copilot/Agents/Skills` |
| `coworkSubfolderName` | Optional sub-folder within the above | `""` or `"Avanade"` |
| `localMachinePath` | Local hint path after download (`{username}` is substituted) | `C:\Users\{username}\.agents\skills` |

If you leave a field blank, the built-in default is used.

### 3. Set up the SharePoint Skills library

1. Create a document library called `Skills` (or whatever matches your config) on the target site.
2. Optionally create sub-folders to group skills by category — sub-folder names become the category labels in the UI.
3. Upload `.md` skill files. The file's **Description** column (internal name: `_ExtendedDescription`) is shown on the card.

### 4. Build and package

```bash
# Build + test + package (production)
npm run build
```

The `.sppkg` package is output to `sharepoint/solution/asorepo.sppkg`.

### Local testing (recommended before packaging)

Because this web part calls SharePoint REST and Microsoft Graph, use the SharePoint hosted workbench for realistic testing.

1. Start the local dev server:

```bash
npm run start
```

2. In your browser, open the hosted workbench in your tenant (replace with your site):

```text
https://yourtenant.sharepoint.com/sites/OrgAISkills/_layouts/15/workbench.aspx?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js
```

3. Add the AI Skills Repo web part.
4. Open the web part property pane and confirm settings:
	- Skills site URL
	- Skills library name
	- OneDrive Cowork path
	- Optional subfolder
	- Local machine path hint
5. Verify end-to-end behavior:
	- Skills load from the configured library.
	- Category filter reflects library sub-folders.
	- Copy to Cowork creates folders/files in OneDrive.
	- Download saves the file and shows the local-path hint.

Optional: You can open the local workbench at `https://localhost:4321/temp/workbench.html` for basic rendering checks, but data operations may not behave the same as hosted SharePoint context.

### 5. Deploy

1. Upload `asorepo.sppkg` to your **SharePoint App Catalog**.
2. Approve the Graph API permission requests (`Files.ReadWrite`, `Sites.Read.All`) in the **SharePoint Admin Center → Advanced → API access**.
3. Add the **AI Skills Repo** web part to any SharePoint page.

---

## Project structure

```
skills-repository/
├── src/
│   └── webparts/skillsRepository/
│       ├── components/
│       │   ├── SkillsRepository.tsx  ← main React component + SkillCard
│       │   ├── SkillsRepository.module.css
│       │   ├── ISkillsRepositoryProps.ts
│       │   └── ISkillsRepositoryState.ts
│       ├── services/
│       │   ├── SkillsService.ts      ← reads skills from SharePoint REST API
│       │   └── OneDriveService.ts    ← copies files to OneDrive via Graph API
│       ├── loc/
│       │   ├── en-us.js              ← UI strings
│       │   └── mystrings.d.ts
│       ├── SkillsRepositoryWebPart.ts
│       └── SkillsRepositoryWebPart.manifest.json
├── config/
│   ├── config.json
│   ├── package-solution.json
│   ├── rig.json
│   ├── sass.json
│   ├── serve.json
│   ├── typescript.json
│   └── write-manifests.json
├── package.json
├── tsconfig.json
└── .yo-rc.json
```

---

## How it works

### Copy to Cowork
1. Fetches the raw `.md` content from SharePoint via REST API.  
2. Ensures the destination folder exists in the user's OneDrive (`me/drive/root:/{path}`) via Microsoft Graph, creating missing folders if needed.  
3. PUTs the file content using the Graph simple-upload endpoint.  
4. The file is then available locally once OneDrive for Business syncs.

### Download to local
Triggers a browser download via the SharePoint `/_layouts/download.aspx` endpoint. After downloading, the UI shows the configured local path as a hint of where to place the file.

### Category detection
Sub-folders within the Skills library are automatically detected and used as category labels. The UI provides a category dropdown to filter by folder/category.

---

## Requirements

| Requirement | Version |
|---|---|
| SPFx | 1.22.2 |
| Node.js | 18.x |
| SharePoint Online | Any |
| Microsoft Graph permissions | `Files.ReadWrite`, `Sites.Read.All` |
