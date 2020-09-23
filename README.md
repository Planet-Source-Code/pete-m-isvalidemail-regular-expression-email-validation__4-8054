<div align="center">

## IsValidEmail \- Regular Expression Email Validation


</div>

### Description

Checks the inputed email address consisting of periods, dashes, underscores and alphanumeric characters. It will also check for the order of characters (no dashes at the begining of the address, etc..). Allows domain extensions between 2 and 7 letters.
 
### More Info
 
If you use this function on your site, please send me an email with any comments you may have.

If the email is valid, IsValidEmail will return a boolean value of True.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Pete M\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pete-m.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Validation/ Processing](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/validation-processing__4-16.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pete-m-isvalidemail-regular-expression-email-validation__4-8054/archive/master.zip)





### Source Code

```
<%
Option Explicit
Function IsValidEmail(strEAddress)
 Dim objRegExpr
 Set objRegExpr = New RegExp
 objRegExpr.Pattern = "^[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]@[\w-\.]*[a-zA-Z0-9]\.[a-zA-Z]{2,7}$"
 objRegExpr.Global = True
 objRegExpr.IgnoreCase = False
 IsValidEmail = objRegExpr.Test(strEAddress)
 Set objRegExpr = Nothing
End Function
If IsValidEmail("my.code@p-s-c.com") = True Then
 Response.Write("Valid Email Address")
Else
 Response.Write("Invalid Email Address")
End If
%>
```

