<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

idUsuarioCreador = request("idUsuarioCreador")
tip_usuario = request("tip_usuario")

proveedorservicioId = request("proveedorservicioId")
proveedorservicioRazonSocial = request("proveedorservicioRazonSocial")
proveedorservicioCiudadAtencion = request("proveedorservicioCiudadAtencion")
proveedorservicioContactoTelefono = request("proveedorservicioContactoTelefono")
proveedorservicioContactoFax = request("proveedorservicioContactoFax")
proveedorservicioContactoWeb = request("proveedorservicioContactoWeb")
proveedorservicioContactoCorreo = request("proveedorservicioContactoCorreo")
proveedorservicioHorario = request("proveedorservicioHorario")
proveedorservicioCalificacion = request("proveedorservicioCalificacion")
proveedorservicioEspecialidad = request("proveedorservicioEspecialidad")
proveedorservicioActivo = SiVacio(request("proveedorservicioActivo"),"0")
proveedorservicioComentario = request("proveedorservicioComentario")

if Request.ServerVariables("REQUEST_METHOD") = "POST" then

	if proveedorservicioId <>"" then
		'Editar
		sql = "update proveedorservicio set proveedorservicioRazonSocial='"& proveedorservicioRazonSocial &"',"
		sql = sql & " proveedorservicioCiudadAtencion='"& proveedorservicioCiudadAtencion &"',"
		sql = sql & " proveedorservicioContactoTelefono='"& proveedorservicioContactoTelefono &"',"
		sql = sql & " proveedorservicioContactoFax='"& proveedorservicioContactoFax &"',"
		sql = sql & " proveedorservicioContactoWeb='"& proveedorservicioContactoWeb &"',"
		sql = sql & " proveedorservicioContactoCorreo='"& proveedorservicioContactoCorreo &"',"
		sql = sql & " proveedorservicioHorario='"& proveedorservicioHorario &"',"
		sql = sql & " proveedorservicioCalificacion="& proveedorservicioCalificacion &","
		sql = sql & " proveedorservicioEspecialidad='"& proveedorservicioEspecialidad &"',"
		sql = sql & " proveedorservicioActivo='"& proveedorservicioActivo &"',"
		sql = sql & " proveedorservicioComentario='"& proveedorservicioComentario &"'"
		sql = sql & " where proveedorservicioId="& proveedorservicioId 
		cn.open strUrl
		cn.execute(sql)
		cn.close	
	
	else
		'Nuevo
		sql = "insert into proveedorservicio(proveedorservicioRazonSocial,"
		sql = sql & "proveedorservicioCiudadAtencion,proveedorservicioContactoTelefono,"
		sql = sql & "proveedorservicioContactoFax ,"
		sql = sql & "proveedorservicioContactoWeb ,"
		sql = sql & "proveedorservicioContactoCorreo ,"
		sql = sql & "proveedorservicioHorario ,"
		sql = sql & "proveedorservicioCalificacion ,"
		sql = sql & "proveedorservicioEspecialidad ,"
		sql = sql & "proveedorservicioActivo,proveedorservicioComentario "
		sql = sql & ") values("
		sql = sql & "'"&proveedorservicioRazonSocial&"',"
		sql = sql & "'"&proveedorservicioCiudadAtencion&"',"
		sql = sql & "'"&proveedorservicioContactoTelefono&"',"
		sql = sql & "'"&proveedorservicioContactoFax&"',"
		sql = sql & "'"&proveedorservicioContactoWeb&"',"
		sql = sql & "'"&proveedorservicioContactoCorreo&"',"
		sql = sql & "'"&proveedorservicioHorario&"',"
		sql = sql & ""&proveedorservicioCalificacion&","
		sql = sql & "'"&proveedorservicioEspecialidad&"',"
		sql = sql & "'"&proveedorservicioActivo&"','"& proveedorservicioComentario &"'"
		sql = sql & ")"
		'response.Write(sql)
		'response.End()
		cn.open strUrl
		cn.execute(sql)
		cn.close	
	
	end if
	
	'Limpiar variables
	proveedorservicioId = ""
	proveedorservicioRazonSocial = ""
	proveedorservicioCiudadAtencion = ""
	proveedorservicioContactoTelefono = ""
	proveedorservicioContactoFax = ""
	proveedorservicioContactoWeb = ""
	proveedorservicioContactoCorreo = ""
	proveedorservicioHorario = ""
	proveedorservicioCalificacion = ""
	proveedorservicioEspecialidad = ""
	proveedorservicioActivo = ""
	
else

	if proveedorservicioId <>"" then
		sql = "select proveedorservicioRazonSocial,"
		sql = sql & " proveedorservicioCiudadAtencion,"
		sql = sql & " proveedorservicioContactoTelefono,"
		sql = sql & " proveedorservicioContactoFax,"
		sql = sql & " proveedorservicioContactoWeb,"
		sql = sql & " proveedorservicioContactoCorreo,"
		sql = sql & " proveedorservicioHorario,"
		sql = sql & " proveedorservicioCalificacion,"
		sql = sql & " proveedorservicioEspecialidad,"
		sql = sql & " proveedorservicioActivo,proveedorservicioComentario"
		sql = sql & " from proveedorservicio  where proveedorservicioId="& proveedorservicioId

		cn.open strUrl
		rs.open sql, cn
		if not rs.eof then
			proveedorservicioRazonSocial = trim(rs.fields(0))		
			proveedorservicioCiudadAtencion = trim(rs.fields(1))
			proveedorservicioContactoTelefono = trim(rs.fields(2))
			proveedorservicioContactoFax = trim(rs.fields(3))
			proveedorservicioContactoWeb = trim(rs.fields(4))
			proveedorservicioContactoCorreo = trim(rs.fields(5))
			proveedorservicioHorario = trim(rs.fields(6))
			proveedorservicioCalificacion = trim(rs.fields(7))
			proveedorservicioEspecialidad = trim(rs.fields(8))
			proveedorservicioActivo = rs.fields(9)
			proveedorservicioComentario = rs.fields(10)
		end if
		rs.close()
		cn.close
		
	end if
	
end if

		sql = "select proveedorservicioId,proveedorservicioRazonSocial,"
		sql = sql & " proveedorservicioCiudadAtencion,"
		sql = sql & " proveedorservicioContactoTelefono,"
		sql = sql & " proveedorservicioContactoFax,"
		sql = sql & " proveedorservicioContactoWeb,"
		sql = sql & " proveedorservicioContactoCorreo,"
		sql = sql & " proveedorservicioHorario,"
		sql = sql & " proveedorservicioCalificacion,"
		sql = sql & " proveedorservicioEspecialidad,"
		sql = sql & " proveedorservicioActivo" ',proveedorservicioComentario"
		sql = sql & " from proveedorservicio"
		sql = sql & " order by proveedorservicioRazonSocial"

		cn.open strUrl
		rs.open sql, cn
		if not rs.eof then
		arrDatos = rs.GetRows()
		end if
		rs.close()
		cn.close

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>::: Med Global Assist ::: </title>
<LINK href="estilo.css" type=text/css rel=stylesheet>
</head>
<body bgcolor="#e9e9e9">
<center>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="titulo" align="left">
    <br />
    Registro de Proveedor
    <br />
	<br />
</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" align="left">
    <fieldset>
            <legend class="titulo_grupo" >Proveedores de Servicio:</legend> 
            <div style="width:100%; height:260px; overflow:scroll">
            <table border="0" align="center" cellpadding="1" cellspacing="1"  width="100%">
              <tr bgcolor="#FF4935">
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid;border-left:#000 1px solid;">
                Id</td>
                <td align="center" class="titulo_tabla" width="40%" style="border-right:#000 1px solid">Razon Social</td>
                <td align="center" class="titulo_tabla" width="30%" style="border-right:#000 1px solid">Ciudad</td>
                <td align="center" class="titulo_tabla" width="20%" style="border-right:#000 1px solid">Telefono</td>
                <!--<td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">Fax</td>          
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">Web</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">Correo</td>-->
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">Activo?</td>
              </tr>
           <%if isArray(arrDatos) then
		   colorfila = 0
		   		for i=0 to ubound(arrDatos,2)
			if colorfila = 0 then
					color= "#FFFFFF"
			'colorfila=1
'			else
'					color="#COCOCO"
'					colorfila=0
			end if
			%>   
              <tr class="contenido_tabla1"  bgcolor="<%=color%>">
                <td align="center">
                <a href="lista_prov-servicio.asp?proveedorservicioId=<%=arrDatos(0,i)%>&idUsuarioCreador=<%=idUsuarioCreador%>&tip_usuario=<%=tip_usuario%>"><%=arrDatos(0,i)%>               	 </a>
                </td>
                <td align="center">&nbsp;<%=ucase(arrDatos(1,i))%></td>               
                <td align="center">&nbsp;<%=ucase(arrDatos(2,i))%></td>               
                <td align="center">&nbsp;<%=ucase(arrDatos(3,i))%></td>               
<!--                <td align="center">&nbsp;<%=ucase(arrDatos(4,i))%></td>               
                <td align="center">&nbsp;<%=ucase(arrDatos(5,i))%></td>               
                <td align="center">&nbsp;<%=ucase(arrDatos(6,i))%></td> -->              
                <td align="center">&nbsp;<%if arrDatos(10,i)="1" then%>SI<%else%>NO<%end if%></td>
              </tr>
              	<%
			 	next
			 end if	
			 %> 
            </table>
</div>
            <br />
           
     </fieldset>       
    	
    </td>
  </tr>
  <tr>
    
    <td align="center" colspan="2">
    <fieldset style="width:97%">
            <legend class="titulo_grupo" >Crear / Actualizar Proveedor:</legend> 
        <form name="frm" method="post" action="lista_prov-servicio.asp"  >
            <table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="contenido_tabla1">
            <tr>
            	<td>Razon Social:</td>
                <td colspan="3"><input name="proveedorservicioRazonSocial" type="text" value="<%=proveedorservicioRazonSocial%>" size="90" maxlength="80" required /></td>
            </tr>
            <tr>
            	<td>Ciudad de Atención:</td>
                <td colspan="2"><input name="proveedorservicioCiudadAtencion" type="text" value="<%=proveedorservicioCiudadAtencion%>" size="60" maxlength="255" /></td>
                <td>Para agregar más de una ciudad utilice la coma.</td>
            </tr>            
            <tr>
	            <td>Telefono:</td>
                <td><input name="proveedorservicioContactoTelefono" type="text" value="<%=proveedorservicioContactoTelefono%>" size="30" maxlength="40" /></td>
                <td>Fax:</td>
                <td><input name="proveedorservicioContactoFax" type="text" value="<%=proveedorservicioContactoFax%>" size="40" maxlength="40" /></td>
            </tr>
            <tr>
	            <td>Web:</td>
                <td><input name="proveedorservicioContactoWeb" type="text" value="<%=proveedorservicioContactoWeb%>" size="30" maxlength="30" /></td>
                <td>Correo:</td>
                <td><input name="proveedorservicioContactoCorreo" type="text" value="<%=proveedorservicioContactoCorreo%>" size="40" maxlength="40" /></td>
            </tr>
            <tr>
	            <td>Horario:</td>
                <td><input name="proveedorservicioHorario" type="text" value="<%=proveedorservicioHorario%>" size="30" maxlength="80" /></td>
                <td>Calificación:</td>
                <td>
                	<select name="proveedorservicioCalificacion" required>
                    	<option value="">[Seleccione]</option>
                        <option value="5" <%if proveedorservicioCalificacion="5" then%> selected="selected" <%end if%> >Excelente</option>
                        <option value="4" <%if proveedorservicioCalificacion="4" then%> selected="selected" <%end if%>>Bueno</option>
                        <option value="3" <%if proveedorservicioCalificacion="3" then%> selected="selected" <%end if%>>Regular</option>
                        <option value="2" <%if proveedorservicioCalificacion="2" then%> selected="selected" <%end if%>>Malo</option>
                        <option value="1" <%if proveedorservicioCalificacion="1" then%> selected="selected" <%end if%>>Pesimo</option>
                    </select>
                </td>
            </tr>
            <tr>
	            <td>Especialidad:</td>
                <td colspan="2"><input name="proveedorservicioEspecialidad" type="text" value="<%=proveedorservicioEspecialidad%>" size="60" maxlength="255" /></td>
                <td>Para agregar más de una especialidad utilice la coma.</td>
            </tr>
            <tr>
            	<td>Comentarios:</td>
                <td colspan="3"><textarea name="proveedorservicioComentario" cols="30" rows="2" style="width:100%"><%=proveedorservicioComentario%></textarea></td>
            </tr>
            <tr>
            	<td>Activo?</td>
                <td><input type="checkbox" name="proveedorservicioActivo" value="1" <%if proveedorservicioActivo="1" then%> checked="checked" <%end if%> /></td>
                <td><input type="submit" name="btnEnviar" value="Grabar Proveedor" /></td>
                <td></td>                                
            </tr> 
          </table>        
        
         <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
         <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
         <input type="hidden" name="proveedorservicioId" value="<%=proveedorservicioId%>" />
                    
        </form>
        </fieldset>
    </td>
  </tr>
</table>
</center>
</body>
</html>
