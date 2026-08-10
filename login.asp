<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

Mensaje=""
cod_usuario = Request("cod_usuario")
des_pass 	= Request("des_pass")

correcto = 1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>::: Innova Doctors :::</title>
<LINK href="admin_login.css" type=text/css rel=stylesheet>
<LINK href="estilo.css" type=text/css rel=stylesheet>
<script language="javascript" type="text/javascript">
	function enviar(user,tipo) {
	    document.loginForm.cod_usuario.value = user;
		document.loginForm.tip_usuario.value = tipo;
		document.loginForm.action = "home.asp";
		document.loginForm.submit();
	}
</script>
</head>
<body>
<form name="loginForm" method="post">
<input name="cod_usuario" type="hidden" value="<%=cod_usuario%>" />
<input name="tip_usuario" type="hidden" value="<%=tip_usuario%>" />
<%
if Request.ServerVariables("REQUEST_METHOD") = "POST" then

pwd_db ="XxX"

sql = " select des_pass,tip_usuario from usuario where cod_usuario='"& cod_usuario &"' and est_activo='S'"
'response.Write(sql)
'response.End()
cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
		pwd_db = trim(rs("des_pass"))
		tip_usuario = trim(rs("tip_usuario"))
		
			if des_pass = pwd_db then
			session("codUsr") = cod_usuario
%>			
			<script>
			  enviar('<%=cod_usuario%>','<%=tip_usuario%>');
			</script> 
<%			else
				Mensaje = "Contraseña no valida!"	
				response.Redirect("index.asp?mensaje="&Mensaje&"&cod_usuario="&cod_usuario&"")		
			end if
	else
		Mensaje = "Usuario no valido!"			
		response.Redirect("index.asp?mensaje="&Mensaje&"&cod_usuario="&cod_usuario&"")				
	end if
	rs.close()
cn.close
end if
%>
</form>
</body>
</html>
