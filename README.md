<div align="center">

## Generic Email Form Handler using CDONTS


</div>

### Description

Email any form from your site using CDONTS (IIS's built-in smtp). Just 10 lines of code handles any size form! Email will display message in the form of fieldname: fieldvalue in proper tab order, with line breaks between each name/value pair.
 
### More Info
 
Required form fields: Email, Subject. Email is the sender's email that will appear in the "From" field on the email. Both fields may hidden inputs if desired. Any number of other inputs may be added as desired.

Basic knowledge of html to set up the form is required. This code does not do anything after sending email, so you may want to redirect to another page or display a message that the mail was sent. Be sure to change "you@yourdomain.com" to your actual email address!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[BebeZed](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bebezed.md)
**Level**          |Beginner
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bebezed-generic-email-form-handler-using-cdonts__4-6085/archive/master.zip)





### Source Code

```
<%
for i=1 to request.form.count
  strMessage = strMessage & request.form.key(i) & ": " & request.form.item(i) & vbCrLf
Next
Set objMail = CreateObject("CDONTS.Newmail")
objMail.From = request.form("Email")
objMail.To = "you@yourdomain.com"
objMail.Subject = request.form("Subject")
objMail.Body = strMessage
objMail.Send
Set objMail = Nothing
%>
```

