<%@LCID = 1034%>
<!--#include file="main.asp"-->
<%
on error resume next

Set rs = Server.CreateObject("ADODB.RECORDSET")
	Mensaje 	= request("Mensaje")
	'id_usuario  = request("id_usuario") 
    des_usuario = request("des_usuario")
	des_pass  	= request("des_pass") 
	est_activo = SiVacio(request("est_activo"),"N")	
	des_email   = request("des_email")
	des_cargo  	= request("des_cargo")
	pais_acceso = request("pais_acceso")	
	est_alerta	= request("est_alerta")
		cod_usuario = request("cod_usuario")
		idUsuarioCreador = request("idUsuarioCreador")
		tip_usuario_usr = request("tip_usuario_usr")
		des_mensaje = request("des_mensaje")

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	
	strSql = " update usuario set des_usuario='"& des_usuario &"'"
	strSql = strSql & " ,des_pass='" & des_pass & "'"	
	strSql = strSql & " ,est_activo='"& est_activo &"'"
	strSql = strSql & " ,des_email='" & des_email & "'"
	strSql = strSql & " ,des_cargo='"& des_cargo &"'"
	strSql = strSql & " ,des_pais='"& pais_acceso &"'"
	strSql = strSql & " ,tip_usuario='"& tip_usuario_usr & "'"
	strSql = strSql & " ,des_mensaje='"& des_mensaje & "'"
	strSql = strSql & " where cod_usuario='"& cod_usuario &"'"	
	'response.Write(strSql)
	'response.End()
	cn.open strUrl
	cn.execute(strSql)
	cn.close
		
	'Grabar Accesos
	strSql = " delete from accesos where id_usuario='"& cod_usuario &"'"
	cn.open strUrl
	cn.execute(strSql)
	cn.close
	
	arrOpciones = split(request("id_opcion"),",")
	if isArray(arrOpciones) then
		for i=0 to ubound(arrOpciones)
		strOpcion = arrOpciones(i)
		strSql = " insert into accesos values('"& trim(cod_usuario) &"','" & trim(strOpcion) &"',DATEADD(hour,-5,GETUTCDATE()))"
		cn.open strUrl
		cn.execute(strSql)
		cn.close
		next
	end if

	response.Redirect("lista_usuarios.asp?idUsuarioCreador="& idUsuarioCreador &"")
else
	strSql = " select des_usuario,des_pass,"
	strSql = strSql & " est_activo,des_email,"
	strSql = strSql & " des_cargo,des_pais,tip_usuario,des_mensaje from usuario"
	strSql = strSql & " where rtrim(cod_usuario)='"& trim(cod_usuario) &"'"
	
	'response.Write(strSql)
	
	cn.open strUrl	
	rs.open strSql, cn
	if not rs.eof then

	des_usuario   = rs(0)
    des_pass  = rs(1)
	est_activo  = rs(2) 
	des_email = rs(3)
	des_cargo  = rs(4)
	pais_acceso = rs(5)
	tip_usuario = rs(6)
	des_mensaje = rs(7)
	
	end if
	rs.close()			
	cn.close
	strSql = " select a.id_opcion,o.des_opcion"
	strSql = strSql & " from accesos a,opciones o"
	strSql = strSql & " where a.id_opcion=o.id_opcion"
	strSql = strSql & " and id_usuario='"& cod_usuario &"'"	
	strSql = strSql & " order by o.num_orden"
	cn.open strUrl
	rs.open strSql, cn
	if not rs.eof then
	arrAccesos = rs.GetRows()
	end if
	rs.close()
	cn.close
	
	strSql = " select o.id_opcion,o.des_opcion from opciones o"
	strSql = strSql & " where o.id_opcion not in(select a.id_opcion"
	strSql = strSql & " from accesos a where id_usuario='"& cod_usuario &"')"	
	strSql = strSql & " order by o.num_orden"
	cn.open strUrl
	rs.open strSql, cn
	if not rs.eof then
	arrOpciones = rs.GetRows()
	end if
	rs.close()
	cn.close
		
end if

'cn.close()
set rs = nothing
set cn = nothing	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Alas</title>
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
<form name="frm" method="post" action="actualizar_usuario.asp">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left"><table width="100%" border="0" cellspacing="1" cellpadding="1">
      <tr>
        <td colspan="2" valign="top" class="titulo">Actualizar Usuario</td>
        <td align="right" rowspan="3">&nbsp;
          <table width="100%" border="0" align="left" cellpadding="0" cellspacing="1" >
                      
            <tr>
              <td class="style5"> Nombres y Apellidos</td>
              <td align="left" class="style4">&nbsp;
                <input type="text" name="des_usuario" size="35" value="<%=des_usuario%>"></td>
              <td width="36%">&nbsp;</td>
              </tr>
            
            
            <tr>
              <td class="style5">Email</td>
              <td align="left" class="style4">&nbsp;
                <input type="text" name="des_email" size="35" value="<%=des_email%>"></td>
              </tr>
            <tr>
              <td class="style5">Cargo</td>
              <td align="left" class="style4" >&nbsp;
                <input type="text" name="des_cargo" size="35" value="<%=des_cargo%>">        </td>
              </tr>
            <tr>
              <td width="26%" class="style5">Pais</td>
              <td align="left">&nbsp;
                <input type="text" name="pais_acceso" size="35" value="<%=pais_acceso%>"> 
  </td>
              </tr>
            <tr>
              <td class="style5">Activo?</td>
              <td align="left" class="style4">&nbsp;
                <select name="est_activo" id="est_activo">
                  <option value="S" <%if est_activo="S" then%> selected="selected" <%end if%> >Si</option>
                  <option value="N" <%if est_activo="N" then%> selected="selected" <%end if%> >No</option>        
                  </select>
                &nbsp;&nbsp;&nbsp;&nbsp;
                &nbsp;&nbsp;&nbsp;</td>
              </tr>
             
              <tr>
              <td class="style5">Tipo usuario</td>
              <td align="left" class="style4">&nbsp;
                <select name="tip_usuario_usr" id="tip_usuario_usr">
                  <option value="A" <%if tip_usuario="A" then%> selected="selected" <%end if%> >Administrador</option>
                  <option value="U" <%if tip_usuario="U" then%> selected="selected" <%end if%> >Usuario</option>        
                  </select>
                &nbsp;&nbsp;&nbsp;&nbsp;
                &nbsp;&nbsp;&nbsp;</td>
              </tr> 
  <tr>
              <td class="style5">ID Usuario</td>
              <td align="left">&nbsp;
                <input type="text" name="cod_usuario" size="35" value="<%=cod_usuario%>" readonly="readonly"></td>
              </tr>
            <tr>
              <td class="style5">Contrase&ntilde;a</td>
              <td align="left">&nbsp;
                <input type="text" name="des_pass" size="35" value="<%=des_pass%>"></td>
              </tr>
              <tr>
              <td class="style5">Mensaje (Máx.144)</td>
              <td align="left">&nbsp;
                <textarea name="des_mensaje" cols="40" rows="4"><%=des_mensaje%></textarea>
                </td>                
              </tr>
            <tr>
              <td>&nbsp;</td>
              <td align="center"><table width="50%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><input type="submit" name="guardar" id="guardar" value="Grabar" class="boton"  /></td>
                  <td><input type="button" name="retornar" id="retornar" value="Retornar" class="boton" onclick="goback();" /></td>
                  </tr>
                
                </table>        </td>
              <td>&nbsp;</td>
              </tr>
            <tr>
              <td class="style3" style="border-bottom:solid 1px #000000">Accesos</td>
              <td colspan="2" align="center" style="border-bottom:solid 1px #000000">&nbsp;</td>
              </tr>
            <tr>
              <td>&nbsp;</td>
              <td colspan="2" align="center">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <%if isArray(arrAccesos) then 
		  	contador = 0
				for i=0 to ubound(arrAccesos,2)
				contador = contador + 1					
				%>
                  <td align="center">&nbsp;<input name="id_opcion" type="checkbox" value="<%=arrAccesos(0,i)%>" checked /></td>
                    <td align="left" class="style6"><%=arrAccesos(1,i)%></td>
                    
                    <%	if contador mod 3=0 then response.write("</tr><tr>")
		  
		  next
		  end if%>
                    <%if isArray(arrOpciones) then 
		  	contador = 0
				for i=0 to ubound(arrOpciones,2)
				contador = contador + 1					
				%>
                    <td align="center">&nbsp;<input name="id_opcion" type="checkbox" value="<%=arrOpciones(0,i)%>" /></td>
                    <td align="left"  class="style6"><%=arrOpciones(1,i)%></td>
                    
                    <%	if contador mod 3=0 then response.write("</tr><tr>")
		  
		  next
		  end if%> 
                    </table>         </td>
              </tr>
            </table>
        </td>
        </tr>
<tr><td width="50">&nbsp;</td><td width="150">&nbsp;</td></tr>
<tr><td></td><td>&nbsp;</td></tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<input name="idUsuarioCreador" type="hidden" value="<%=idUsuarioCreador%>" />
</form>
</center>
</body>
</html>
