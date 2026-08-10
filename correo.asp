<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" 
NAME="CDO for Windows Library" -->
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" 
NAME="ADODB Type Library" -->

<%@ CodePage=65001 Language="VBScript"%> 

<%
Dim ObjectSendMail 
Dim Conf 
Dim Fields

Set ObjectSendMail = Server.CreateObject("CDO.Message") 
Set Conf = Server.CreateObject("CDO.Configuration")

strPara = trim(request("strPara"))
'txtApellido = request("txtApellido")
'txtEmpresa = request("txtEmpresa")
'txtTelefono = request("txtTelefono")
strAsunto = request("strAsunto")
'txtMotivo = request("txtMotivo")
strMensaje = request("strMensaje")
idCaso = request("idCaso")

msj="ok"

Set Fields = Conf.Fields 
With Fields 
.Item(cdoSendUsingMethod) = 2
.Item(cdoSMTPServer) = "mail.euroamericanassistance.com" 
.Item(cdoSMTPServerPort) = 25 
.Item(cdoSMTPconnectiontimeout) = 10
.Item(cdoSMTPAuthenticate) = cdoBasic
.Item(cdoSendUserName) = "webmaster@euroamericanassistance.com"
.Item(cdoSendPassword) = "webmaster#1"
.Update 
End With

Set ObjectSendMail.Configuration = Conf
ObjectSendMail.Subject = strAsunto
ObjectSendMail.From = "contactoweb@networkcareassist.com"
ObjectSendMail.To = strPara

'ObjectSendMail.AddAttachment 

'ObjectSendMail.TextBodyPart.Charset = "utf-8" 
ObjectSendMail.BodyPart.Charset = "utf-8" 
'ObjectSendMail.HTMLBodyPart.Charset = "utf-8"

ObjectSendMail.HTMLBody = strMensaje 
ObjectSendMail.HTMLBodyPart.Charset = "utf-8"
ObjectSendMail.TextBody = strMensaje
ObjectSendMail.TextBodyPart.Charset = "utf-8" 

on error resume next

ObjectSendMail.Send

Set ObjectSendMail = Nothing 
Set Conf = Nothing 
Set Fields = Nothing

If Err.Number <> 0 Then
	msj="ko"
	  Response.Write (Err.Description& "<br><br>")   
	  response.End()
End if

response.Redirect("EnviarCorreo.asp?msj="&msj&"&idCaso="&idCaso&"")
%>