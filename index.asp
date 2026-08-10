<%@LCID = 1034%>
<%
Mensaje = request("Mensaje")
cod_usuario = request("cod_usuario")

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Innova Doctors </title>
<LINK href="admin_login.css" type=text/css rel=stylesheet>
<LINK href="estilo.css" type=text/css rel=stylesheet>
<script language="javascript" type="text/javascript">
	function setFocus() {
		document.loginForm.cod_usuario.select();
		document.loginForm.cod_usuario.focus();
	}
</script>
</head>
<body onLoad="setFocus();">
<div id="wrapper">
    <div id="header">
           <div ><span class="style1">Sistema de </span> <span class="style3">Gestión de Casos</span></div>
  	</div>
    <div id="divider" style="width:100%"><img src="images/div.png" alt="" border="0" width="100%" height="5" /></div>
</div>
<div id="ctr" align="center">
	<div class="login">
		<div class="login-form">
			<div class="ctr"><img src="logo.png" /></div>
       	  <form action="login.asp" method="post" name="loginForm" id="loginForm">
			<div class="form-block">
	        	<div class="inputlabel"> Usuario</div>
		    	<div><input name="cod_usuario" type="text" class="inputbox" size="15" value="<%=cod_usuario%>" /></div>
	        	<div class="inputlabel"> Contraseña </div>
		    	<div><input name="des_pass" type="password" class="inputbox" size="15" /></div>
				<br />                                
	        	<div align="center"><input type="submit" name="submit" class="button" value="Login" /><br />
<p>Utilice un nombre de usuario y contraseña validos para poder accesar al modulo de administrador. </p>
	        	</div>
        	</div>
		  </form>
   	  </div>
<div class="login-text">
			
   	  
	  </div>
		<div class="clr">
        <font color="#FF0000"><%=Mensaje%></font></div>
	</div>
</div>
<div id="break"></div>


</body>

</html>
