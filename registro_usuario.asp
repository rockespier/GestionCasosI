<%@LCID = 1034%>
<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")
	Mensaje 		= request("Mensaje")
	'id_usuario  	= request("id_usuario") 
    des_usuario  	= request("des_usuario")
	des_pass  		= request("des_pass") 
	est_activo  	= SiVacio(request("est_activo"),"N")	
	des_email    	= request("des_email")
	des_cargo  		= request("des_cargo")
	pais_acceso  	= request("pais_acceso")	
	est_alerta		= SiVacio(request("est_alerta"),"N")
	cod_usuario 	= request("cod_usuario")
	idUsuarioCreador = request("idUsuarioCreador")
		tip_usuario_usr = request("tip_usuario_usr")

if Request.ServerVariables("REQUEST_METHOD") = "POST"  and request("txtModo")<>"" then
	
	strSql = " insert into usuario(cod_usuario,des_usuario,des_pass,"
	strSql = strSql & " est_activo,des_email,"
	strSql = strSql & " des_cargo,des_pais,"
	strSql = strSql & " tip_usuario) values("
	strSql = strSql & "'"& cod_usuario &"','" & des_usuario & "','" & des_pass & "','"& est_activo &"',"
	strSql = strSql & "'"& des_email &"','"& des_cargo &"','"& pais_acceso &"','"&tip_usuario_usr&"'"		
	strSql = strSql & ")"
	'response.Write(strSql)
	'response.End()
	cn.open strUrl
	cn.execute(strSql)
	cn.close
	response.Redirect("lista_usuarios.asp?idUsuarioCreador="& idUsuarioCreador &"")
		
end if

		
'cn.close()
set rs = nothing
set cn = nothing	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>::: Network Care Assist ::: </title>
<LINK href="estilo.css" type=text/css rel=stylesheet>
<script language="javascript">
function grabar()
{
frm.submit();
}
function goback(){
history.back();
}

function cuenta(){ 
       document.forms[0].caracteres.value=document.forms[0].des_detalle.value.length 
} 

</script>
</head>

<body>
<center>
<form name="frm" method="post" action="registro_usuario.asp">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left"><table width="100%" border="0" cellspacing="1" cellpadding="1">
      <tr>
        <td colspan="2" valign="top" class="titulo">Actualizar Usuario</td>
        <td align="right" rowspan="3">&nbsp;
        <table width="100%" border="0" align="left" cellpadding="0" cellspacing="1" >
     
      <tr>
        <td class="style5">Nombres y Apellidos</td>
        <td align="left" class="style4">&nbsp;
          <input name="des_usuario" type="text" value="<%=des_usuario%>" size="35" maxlength="50"></td>
          <td width="36%">&nbsp;</td>
        </tr>
      

      <tr>
        <td class="style5">Email</td>
        <td align="left" class="style4">&nbsp;
        <input name="des_email" type="text" value="<%=des_email%>" size="35" maxlength="50"></td>
        </tr>
      <tr>
        <td class="style5">Cargo</td>
        <td align="left" class="style4" >&nbsp;
        <input name="des_cargo" type="text" value="<%=des_cargo%>" size="35" maxlength="50">
        </td>
        </tr>
         <tr>
        <td width="26%" class="style5">Pais</td>
        <td width="38%" align="left" class="style4">&nbsp;
                <input name="pais_acceso" type="text" value="<%=pais_acceso%>" size="35" maxlength="50"> </td>
        
      </tr>
     <tr>
        <td class="style5">Activo</td>
        <td align="left">&nbsp;
        <select name="est_activo" id="est_activo">
        <option value="S" <%if est_activo="S" then%> selected="selected" <%end if%> >Si</option>
        <option value="N" <%if est_activo="N" then%> selected="selected" <%end if%> >No</option>        
        </select></td>

      </tr>
       <tr>
              <td class="style5">Tipo usuario</td>
              <td align="left" class="style4">&nbsp;
                <select name="tip_usuario_usr" id="tip_usuario_usr">
                  <option value="A" <%if tip_usuario_usr="A" then%> selected="selected" <%end if%> >Administrador</option>
                  <option value="U" <%if tip_usuario_usr="U" then%> selected="selected" <%end if%> >Usuario</option>        
                  </select>
                &nbsp;&nbsp;&nbsp;&nbsp;
                &nbsp;&nbsp;&nbsp;</td>
              </tr> 
      <tr>
        <td class="style5">ID Usuario</td>
        <td align="left">&nbsp;
        <input type="text" name="cod_usuario" size="35" value="<%=cod_usuario%>" maxlength="15"></td>

      </tr>
       <tr>
        <td class="style5">Contrase&ntilde;a</td>
        <td align="left">&nbsp;
        <input type="text" name="des_pass" size="35" value="<%=des_pass%>" maxlength="15"></td>

      </tr>
      <tr>
        <td>&nbsp;</td>
        <td align="center">
        <table width="50%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><input type="submit" name="guardar" id="guardar" value="Grabar" class="boton"  /></td>
            <td><input type="button" name="retornar" id="retornar" value="Retornar" class="boton" onclick="goback();" /></td>
          </tr>
        </table></td>
        <td>&nbsp;</td>
      </tr>
    </table>
        </td>
        </tr>
<tr><td>&nbsp;</td><td>&nbsp;</td></tr>
<tr><td></td><td>&nbsp;</td></tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<input name="idUsuarioCreador" type="hidden" value="<%=idUsuarioCreador%>" />
 <input type="hidden" name="txtModo" value="OK" />
</form>
</center>
</body>
</html>
