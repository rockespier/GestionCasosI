<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

idUsuarioCreador = request("idUsuarioCreador")
tip_usuario = request("tip_usuario")

servicioId = request("servicioId")
servicioNombre = request("servicioNombre")
servicioTipoServicioId = request("servicioTipoServicioId")
servicioActivo = SiVacio(request("servicioActivo"),"0")

if Request.ServerVariables("REQUEST_METHOD") = "POST" then

	if servicioId <>"" then
		'Editar
		sql = "update serviciocaso set servicioNombre='"& servicioNombre &"',"
		sql = sql & " servicioTipoServicioId="& servicioTipoServicioId &","
		sql = sql & " servicioActivo="& servicioActivo &""
		sql = sql & " where servicioId="& servicioId 
		cn.open strUrl
		cn.execute(sql)
		cn.close	
	
	else
		'Nuevo
		sql = "insert into serviciocaso(servicioNombre,"
		sql = sql & "servicioTipoServicioId,"
		sql = sql & "servicioActivo "
		sql = sql & ") values("
		sql = sql & "'"&servicioNombre&"',"
		sql = sql & ""&servicioTipoServicioId&","
		sql = sql & ""&servicioActivo&""
		sql = sql & ")"
		cn.open strUrl
		cn.execute(sql)
		cn.close	
	
	end if
	
	'Limpiar variables
	servicioId = ""
	servicioNombre = ""
	servicioTipoServicioId = ""
	servicioActivo = ""
	
else

	if servicioId <>"" then
		sql = "select servicioNombre,"
		sql = sql & " servicioTipoServicioId,"
		sql = sql & " servicioActivo"
		sql = sql & " from serviciocaso  where servicioId="& servicioId

		cn.open strUrl
		rs.open sql, cn
		if not rs.eof then
			servicioNombre = trim(rs.fields(0))		
			servicioTipoServicioId = trim(rs.fields(1))
			servicioActivo = trim(rs.fields(2))
		end if
		rs.close()
		cn.close
		
	end if
	
end if

		sql = "select servicioId,servicioNombre,"
		sql = sql & " servicioTipoServicioId,"
		sql = sql & " desValor,"
		sql = sql & " servicioActivo"
		sql = sql & " from serviciocaso, tipo"
		sql = sql & " where servicioTipoServicioId = desTipo"
		sql = sql & " and desColumna='TipoServicioId' and estRegistro='S'"
		sql = sql & " order by servicioNombre"

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
    Registro de Servicio
    <br />
	<br />
</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" align="center">
    <fieldset>
            <legend class="titulo_grupo" >Servicios:</legend> 
            <div style="width:100%; height:260px; overflow-y:auto;">
            <table border="0" align="center" cellpadding="1" cellspacing="1" class="contenido_tabla1" width="98%">
              <tr bgcolor="#FF4935">
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid;border-left:#000 1px solid;">
                Id</td>
                <td align="center" class="titulo_tabla" width="50%" style="border-right:#000 1px solid">Descripción</td>
                <td align="center" class="titulo_tabla" width="30%" style="border-right:#000 1px solid">Tipo</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">Activo?</td>
              </tr>
           <%if isArray(arrDatos) then
		   colorfila = 0
		   		for i=0 to ubound(arrDatos,2)
					if colorfila = 0 then
							color= "#FFFFFF"
					end if
			%>   
              <tr class="contenido_tabla1"  bgcolor="<%=color%>">
                <td align="center">
                <a href="lista_servicio.asp?servicioId=<%=arrDatos(0,i)%>&idUsuarioCreador=<%=idUsuarioCreador%>&tip_usuario=<%=tip_usuario%>"><%=arrDatos(0,i)%>               	 </a>
                </td>
                <td align="center">&nbsp;<%=ucase(arrDatos(1,i))%></td>               
                <td align="center">&nbsp;<%=ucase(arrDatos(3,i))%></td>               
                <td align="center">&nbsp;<%if arrDatos(4,i)="1" then%>SI<%else%>NO<%end if%></td>
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
            <legend class="titulo_grupo" >Crear / Actualizar Servicio:</legend> 
        <form name="frm" method="post" action="lista_servicio.asp"  >
            <table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="contenido_tabla1">
            <tr>
            	<td>Descripción:</td>
                <td colspan="3"><input name="servicioNombre" type="text" value="<%=servicioNombre%>" size="90" maxlength="40" required /></td>
            </tr>
            
            <tr>	          
                <td>Tipo:</td>
                <td>
                	<select name="servicioTipoServicioId">
                    	<option value="">[Seleccione]</option>
                        <option value="1" <%if servicioTipoServicioId="1" then%> selected="selected" <%end if%> >Simple</option>
                        <option value="2" <%if servicioTipoServicioId="2" then%> selected="selected" <%end if%>>Middle</option>
                        <option value="3" <%if servicioTipoServicioId="3" then%> selected="selected" <%end if%>>Complex</option>                        
                    </select>
                </td>
            </tr>    
            <tr>
            	<td>Activo?</td>
                <td><input type="checkbox" name="servicioActivo" value="1" <%if servicioActivo="1" then%> checked="checked" <%end if%> /></td>
                <td><input type="submit" name="btnEnviar" value="Grabar Servicio" /></td>
                <td></td>                                
            </tr> 
          </table>        
        
         <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
         <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
         <input type="hidden" name="servicioId" value="<%=servicioId%>" />
                    
        </form>
        </fieldset>
    </td>
  </tr>
</table>
</center>
</body>
</html>
