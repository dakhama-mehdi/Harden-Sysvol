### HardenSysvol - Regular Expression Support

#### Adding Regular Expressions

In addition to simple patterns, **HardenSysvol** allows you to add **regular expressions** for more advanced matching. This is done using the same `-addpattern` flag, but specifying regular expressions as patterns.

#### Example Usage with Regular Expressions

You can specify regular expressions to match specific formats, such as UPNs (User Principal Names), email addresses, or credit card numbers. Here's an example:

```powershell
Invoke-HardenSysvol -addpattern "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b,\b\d{16}\b"
````

This example includes:

Email format: \b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b
Matches typical email addresses like user@example.com.

Credit card format: \b\d{16}\b
Matches a 16-digit sequence, commonly used for credit card numbers.

Example for UPNs:
To detect UPNs (User Principal Names), which resemble email addresses but are specific to Active Directory:

```powershell
Invoke-HardenSysvol -addpattern "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\b"
````
