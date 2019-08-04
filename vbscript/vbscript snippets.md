# VBScript Snippets

## Get Account Name from SID

```vbscript
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set objAccount = objWMIService.Get _
    ("Win32_SID.SID='S-1-5-20'")
Wscript.Echo objAccount.AccountName
Wscript.Echo objAccount.ReferencedDomainName
```
