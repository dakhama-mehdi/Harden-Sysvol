### HardenSysvol - Custom Magic Number Definition

#### Adding Your Own Magic Number

It is possible to define your own **magic number** in the `extensions.json` file used by the PowerShell module. A magic number helps identify file types based on their binary header rather than relying on file extensions alone.

#### Steps to Add a Magic Number

1. Open the `extensions.json` file located in the module directory.
2. Add your custom magic number using the following format:

```json
{
  "magic": ["4D53465402000100"],
  "extensions": ["tlb"]
}
