### HardenSysvol - Custom Pattern Support

#### Adding Custom Patterns

It is possible and useful to add custom patterns and regular expressions to the **HardenSysvol** command to further refine the audit. This can be done using the `-addpattern` flag, followed by a comma-separated list of patterns.

#### Example Usage:

```powershell
Invoke-HardenSysvol -addpattern admin,administrator,log,ssh
````
