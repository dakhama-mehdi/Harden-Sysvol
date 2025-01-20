### HardenSysvol - Advanced Options

#### Ignoring Specific Extensions

The `-ignoreextensions` flag allows you to skip specific file extensions during the audit. This is particularly useful for avoiding false positives from files like `.adm`, `.admx`, or `.adml`.

##### Example Usage:

```powershell
Invoke-HardenSysvol -ignoreextensions adm,admx,adml
````

In this example, HardenSysvol will ignore all files with .adm, .admx, or .adml extensions.

Limiting Files by Size
-maxfilesize
Use this flag to list files exceeding a specific size (in MB). For example:

Example Usage:
powershell
Copier
Modifier
Invoke-HardenSysvol -maxfilesize 1
This command will list all files larger than 1 MB.

-maxsizebinary
This flag lists binary files exceeding a specific size (in MB). For example:

Example Usage:
```powershell
Invoke-HardenSysvol -maxsizebinary 5
````
This command will list all binary files larger than 5 MB.

Combined Example
To run an audit that:

Ignores .adm, .admx, and .adml files,
Lists files larger than 1 MB,
And lists binary files larger than 5 MB:
```powershell
Invoke-HardenSysvol -ignoreextensions adm,admx,adml -maxfilesize 1 -maxsizebinary 5
````
This ensures the audit focuses on relevant files while filtering out unnecessary or false-positive entries.
