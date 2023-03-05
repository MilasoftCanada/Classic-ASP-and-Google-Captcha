<%@LANGUAGE=VBSCRIPT%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>Sample Code</title>

<script src='https://www.google.com/recaptcha/api.js'></script>


   <body>
<!-- #include file="aspJSON.asp"-->


<form name="form1" method="post" action="index.asp">
  <div class="g-recaptcha" data-sitekey="6Le3U7cjAAAAABkNnxsprpAdIKBJgbvFIv7Tan9u"></div>
            <input type="submit" value="Submit">   

<%
' when the Captcha is entered properly, the form submits and then index2.asp is loaded
' you could also pass a variable through index2.asp such as index2.asp?verified=yes
' to ensure form is submitted

    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
        Dim recaptcha_secret, sendstring, objXML
        ' Secret key
        recaptcha_secret = "SECRET KEY HERE"

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
response.end
    end if

    Set objXML = Nothing

end if

if request.form  <> "" then
str = "index2.asp"
response.redirect str
end if

%>
</body>
