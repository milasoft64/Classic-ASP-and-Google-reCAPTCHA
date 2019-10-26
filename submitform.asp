<%@LANGUAGE=VBSCRIPT%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>Contact</title>
</head>
    <script src="https://www.google.com/recaptcha/api.js" async defer></script>

   <body>
<!-- #include file="aspJSON.asp"-->

<center>

<form name="form1" method="post" action="submitform.asp">
<br><br>
    <div align="center"> 
      <table width="50%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><b><font color=white>E-mail address</b></td>
          <td><input class="black" type="text" name="txtName" value="<%=request.form("txtName")%>">
          </td></tr>
        <tr> 
          <td><font color=white><b>Subject</b></td>
          <td><input class="black" type="text" name="txtEmail" value="<%=request.form("txtEmail")%>">
          </td></tr>
        <tr> 
          <td><font color=white><b>Comment</b></td>
        <td><textarea class="black" name="txtFeedback" cols="40" rows="7"><%=request.form("txtFeedback")%></textarea>
</td></tr>
<tr><td>
  <div class="g-recaptcha" data-sitekey="**** KEY HERE ****"></div>
            <br />
            <input type="submit" value="Submit">   
</td></tr></table>
 


<%
' Based upon code from https://stackoverflow.com/questions/30711884/how-to-implement-google-recaptcha-2-0-in-asp-classic
' This version by Milasoft64 (2019)
' -Retains form fields if errors occur
' -Form field validation
' -Sends form details by email
'
' replace secret key and site keys with your own Captcha v2 keys
err = 0 ' no errors in form

    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
        Dim recaptcha_secret, sendstring, objXML
        ' Secret key
        recaptcha_secret = "**** GOOGLE RECAPTCHA SECRET KEY HERE ****"

        sendstring = "https://www.google.com/recaptcha/api/siteverify?secret=" & recaptcha_secret & "&response=" & Request.form("g-recaptcha-response")

        Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
        objXML.Open "GET", sendstring, False

        objXML.Send


    result = (objXML.responseText)
    Set objXML = Nothing


 Set oJSON = New aspJSON
    oJSON.loadJSON(result)

    success = oJSON.data("success")
    if success <> "True" then
response.write "You have to complete the reCaptcha"
err = 1
    end if

    Set objXML = Nothing

end if

' has the form been submitted? check fields now...

if request.form  <> "" then


'receive the form values
Dim sName, sEmail, sFeedback
sName=Request.Form("txtName")
sEmail=Request.Form("txtEmail")
sFeedback=Request.Form("txtFeedback")


if sname = "" then
response.write "<br><br><center><h4>No name entered."
err = 1
end if

if InStr(sname, "@") = 0 or InStr(sname, ".") = 0 then
response.write "<br><br><center><h4>Impropertly formatted e-mail address"
err = 1
end if

if semail = "" then
response.write "<br><br><center><h4>No subject entered."
err = 1
end if

if sfeedback = "" then
response.write "<br><br><center><h4>No message entered."
err = 1
end if

if err = 1 then response.end

' safeguards
sname = replace(sname,"union","")
sname = replace(sname,"%","")
sname = replace(sname,"'","''")
sname = replace(sname,";","")
sname = replace(sname,"<?","")

semail = replace(semail,"union","")
semail = replace(semail,"%","")
semail = replace(semail,"'","''")
semail = replace(semail,";","")
semail = replace(semail,"<?","")

sfeedback = replace(sfeedback,"%","")
sfeedback = replace(sfeedback,"'","''")
sfeedback = replace(sfeedback,"UNION","")
sfeedback = replace(sfeedback,"select","")
sfeedback = replace(sfeedback,";","")
sfeedback = replace(sfeedback,"php","")
sfeedback = replace(sfeedback,"<%","")
sfeedback = replace(sfeedback,"echo","")
sfeedback = replace(sfeedback,"<?","")



' create the HTML formatted email text
Dim sEmailText
 
sEmailText = sEmailText & "<html>"
sEmailText = sEmailText & "<head>"
sEmailText = sEmailText & "<title>HTML Email</title>"
sEmailText = sEmailText & "</head>"
sEmailText = sEmailText & "<body>"
sEmailText = sEmailText & "Subject: " & semail & "<br>"
sEmailText = sEmailText & "Message from: " & sName & "<br>"
sEmailText = sEmailText & "Message:" & sFeedback & "<br>"

sDate = Now()
sEmailText = sEmailText & "Date & Time:" & sNow & "<br>"
sEmailText = sEmailText & "IP :" & Request.ServerVariables("REMOTE_ADDR")
sEmailText = sEmailText & "</body>"
sEmailText = sEmailText & "</html>"

'create the mail object 

 set objMessage = createobject("cdo.message") 
 set objConfig = createobject("cdo.configuration") 
 Set Flds = objConfig.Fields 

 Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
 Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="127.0.0.1" 
 Flds.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 

 Flds.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0 
 Flds.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="noreply@yourdomain.com" 


 Flds.update 
 Set objMessage.Configuration = objConfig 
 objMessage.To = "youremailaddress@domain.com"
 objMessage.From = "noreply@yourdomain.com" 
 objMessage.Subject = "New Contact Message"
Objmessage.HTMLBody = sEmailtext
 objMessage.fields.update 



Objmessage.Send
Set ObjSendMail = Nothing 

Response.write "<div align='center'><br><br><font size=3>Thank you for sending your message.<br></div>"

End If 

%>
</body>
