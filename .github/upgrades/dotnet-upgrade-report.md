# .NET 9 Upgrade Report

## Project target framework modifications

| Project name                                   | Old Target Framework    | New Target Framework         | Commits                   |
|:-----------------------------------------------|:-----------------------:|:----------------------------:|---------------------------|
| AllToMarkdown.csproj                           |   netstandard2.0        | net9.0                       |                           |

## NuGet Packages

| Package Name                        | Old Version | New Version | Commit Id                                 |
|:------------------------------------|:-----------:|:-----------:|-------------------------------------------|
| DocSharp.SystemDrawing              |   0.17.0    |             |                                           |

## All commits

| Commit ID              | Description                                |
|:-----------------------|:-------------------------------------------|

## Project feature upgrades

Contains summary of modifications made to the project assets during different upgrade stages.

### AllToMarkdown.csproj

Here is what changed for the project during upgrade:

- Replaced the missing SupportedFileExtensions with a runtime-generated list of supported extensions based on the SupportedFileFormat enum, preserving the original logic.

## Next steps