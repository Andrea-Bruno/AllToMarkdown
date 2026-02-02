# .NET 9.0 Upgrade Plan

## Execution Steps

Execute steps below sequentially one at a time in the order they are listed.

1. Validate that an .NET 9.0 SDK required for this upgrade is installed on the machine and if not, help to get it installed.
2. Ensure that the SDK version specified in global.json files is compatible with the .NET 9.0 upgrade.
3. Upgrade AllToMarkdown.csproj

## Settings

This section contains settings and data used by execution steps.

### Excluded projects

Table below contains projects that do belong to the dependency graph for selected projects and should not be included in the upgrade.

| Project name                                   | Description                 |
|:-----------------------------------------------|:---------------------------:|
|                                               |                            |

### Aggregate NuGet packages modifications across all projects

NuGet packages used across all selected projects or their dependencies that need version update in projects that reference them.

| Package Name                        | Current Version | New Version | Description                                   |
|:------------------------------------|:---------------:|:-----------:|:----------------------------------------------|
| DocSharp.SystemDrawing              |   0.17.0        |             | Non è stata trovata alcuna versione supportata |

### Project upgrade details
This section contains details about each project upgrade and modifications that need to be done in the project.

#### AllToMarkdown.csproj modifications

Project properties changes:
  - Target framework should be changed from `netstandard2.0` to `net9.0`

NuGet packages changes:
  - DocSharp.SystemDrawing should be updated from `0.17.0` to  (Non è stata trovata alcuna versione supportata)

Feature upgrades:

Other changes: