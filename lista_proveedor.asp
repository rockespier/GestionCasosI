<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

desBusqueda = request("desBusqueda")

idUsuarioCreador = request("idUsuarioCreador")
tip_usuario = request("tip_usuario")
idProveedor = request("idProveedor")
desProveedor = request("desProveedor")
idProveedorPadre = request("idProveedorPadre")
idPadre = request("idPadre")

if Request.ServerVariables("REQUEST_METHOD") = "POST" and desProveedor<>"" then

	if idProveedor<>"" then
	
		'borrar
		sql = " delete from proveedorxespecialidad  where id_proveedor=" & idProveedor
		cn.open strUrl
		cn.execute(sql)
		cn.close
'	
		for i=0 to cint(SiVacio(Request.Form("txtNumReg"),0))
		
		idEspecialidadPro  = Request.Form("idEspecialidadPro"& i &"")		
		textoEspecialidad  = Request.Form("textoEspecialidad"& i &"")			
		
			sql = "insert into proveedorxespecialidad "
			sql = sql & " values("&idProveedor&",'"&idEspecialidadPro&"','"& textoEspecialidad &"','"& idUsuarioCreador &"',DATEADD(hour,-5,GETUTCDATE()))" 
			cn.open strUrl
			cn.execute(sql)
			cn.close
'response.Write(sql) & "<br>"
			
		Next 
		
		idEspecialidadPro = request.Form("idEspecialidadPro")
		if idEspecialidadPro<>"" then
			sql = "insert into proveedorxespecialidad "
			sql = sql & " values("&idProveedor&",'"&idEspecialidadPro&"','','"& idUsuarioCreador &"',DATEADD(hour,-5,GETUTCDATE()))" 
			cn.open strUrl
			cn.execute(sql)
			cn.close
			
		end if
	
		'Editar
		sql = "update proveedor set desProveedor='"& desProveedor &"',"
		sql = sql & "idProveedorPadre='"&idProveedorPadre&"'"
		sql = sql & " where idProveedor="& idProveedor 
		cn.open strUrl
		cn.execute(sql)
		cn.close	
	
	else
	'Nuevo
	sql = "insert into proveedor(desProveedor,idProveedorPadre,desArchivo,"
	sql = sql & "estRegistro,cod_usuario,fchCreacion) "
	sql = sql & " values('"&desProveedor&"','"&idProveedorPadre&"','','S','"&idUsuarioCreador&"',DATEADD(hour,-5,GETUTCDATE()))"
	cn.open strUrl
	cn.execute(sql)
	cn.close	
	
	end if
	
	
	'Limpiar variables
	idProveedor = ""
	desProveedor = ""
	idProveedorPadre = ""
else

	if idProveedor<>"" then
		sql = "select desProveedor,idProveedorPadre"
		sql = sql & " from proveedor where idProveedor="& idProveedor

		cn.open strUrl
		rs.open sql, cn
		if not rs.eof then
			desProveedor = rs.fields(0)
			idProveedorPadre = rs.fields(1)	
		end if
		rs.close()
		cn.close
		
	end if
	
end if

sql = "select a.idProveedor,a.desProveedor"
sql = sql & " from proveedor a where a.idProveedorPadre=0"
if desBusqueda<>"" then
	sql = sql & " and a.desProveedor like '%" & desBusqueda & "%'"
	sql = sql & " or a.idProveedor in (select b.idProveedorPadre from proveedor b"
	sql = sql & " where b.desProveedor like '%" & desBusqueda & "%')"
end if
sql = sql & " order by a.desProveedor"

	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
	arrProveedorPadre = rs.GetRows()
	end if
	rs.close()
	cn.close
	
sql = "select id_especialidad,des_especialidad "
sql = sql & " from especialidades where estregistro='S'"
sql = sql & "  order by 2"

	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
	arrEspecialidad = rs.GetRows()
	end if
	rs.close()
	cn.close	

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>::: Network Care Assist ::: </title>
<LINK href="estilo.css" type=text/css rel=stylesheet>
<SCRIPT language="JavaScript">
function popitup(url)
{
	newwindow=window.open(url,'name','height=270,width=400,top=200,left=100,scrollbars=yes');
	if (window.focus) {newwindow.focus()}
	//return false;
}
</script>
</head>
<body bgcolor="#e9e9e9">
<center>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="titulo" align="left">
    <br />
    Registro de proveedores
    <br />
	<br />
</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="60%" align="Left">
    <form name="frm" action="lista_proveedor.asp" method="post">
    	<fieldset style="width:500px">
            <legend class="titulo_grupo" >Buscar por:</legend>            
            <table width="400" border="0" cellspacing="0" cellpadding="0" class="contenido_tabla1">
              <tr>
                <td width="200">Descripcion (Pais/Ciudad):</td>
                <td width="100"><input type="text" name="desBusqueda" size="20" value="<%=desBusqueda%>" /></td>
                <td  align="left"><input type="submit" value="Buscar" /></td>
              </tr>
            </table>          
        </fieldset>
        <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
     	<input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
     </form>    
    </td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" align="left">
    <fieldset>
            <legend class="titulo_grupo" >Proveedores:</legend> 
            <div style="width:100%; height:260px; overflow:scroll-y">
            <table border="0" align="left" cellpadding="1" cellspacing="1" class="contenido_tabla1">
              
           <%if isArray(arrProveedorPadre) then
		   colorfila = 0
		   		for i=0 to ubound(arrProveedorPadre,2)
			if colorfila = 0 then
					color= "#FFFFFF"
			'colorfila=1
'			else
'					color="#COCOCO"
'					colorfila=0
			end if
			%>   
              <tr class="contenido_tabla1">
                <td align="left">
    <a href="lista_proveedor.asp?idPadre=<%=arrProveedorPadre(0,i)%>&idProveedor=<%=arrProveedorPadre(0,i)%>&idUsuarioCreador=<%=idUsuarioCreador%>&tip_usuario=<%=tip_usuario%>">
               	<img src="images/carpeta1.png" width="30" height="22" border="0" /> </a>
                </td>
                <td align="left">
<a href="lista_proveedor.asp?idPadre=<%=arrProveedorPadre(0,i)%>&idProveedor=<%=arrProveedorPadre(0,i)%>&idUsuarioCreador=<%=idUsuarioCreador%>&tip_usuario=<%=tip_usuario%>"><%=ucase(arrProveedorPadre(1,i))%></a></td>
              </tr>
<%				if trim(idPadre) = trim(arrProveedorPadre(0,i)) then
					sql = "select idProveedor,desProveedor,desArchivo,idProveedorPadre"
					sql = sql & " from proveedor where idProveedorPadre="& idPadre 'arrProveedorPadre(0,i)
					sql = sql & " order by desProveedor"
					
						cn.open strUrl
						rs.open sql, cn
						if not rs.eof then
							arrProveedorHijo = rs.GetRows()
						else
							arrProveedorHijo = ""
						end if
						rs.close()
						cn.close
						
					if isArray(arrProveedorHijo) then	
						for j=0 to ubound(arrProveedorHijo,2)
							if ubound(arrProveedorHijo,2) = 0 then
								imagen = "file_abajo.png"
							elseif j= ubound(arrProveedorHijo,2) then
								imagen = "file_abajo1.png"	
							else
								imagen = "file_centro1.png"	
							end if
				%>
                	<tr>
                    	<td style="padding-left:15px">
                         <a href="lista_proveedor.asp?idPadre=<%=arrProveedorHijo(3,j)%>&idProveedor=<%=arrProveedorHijo(0,j)%>&idUsuarioCreador=<%=idUsuarioCreador%>&tip_usuario=<%=tip_usuario%>">
                        <img src="images/<%=imagen%>" width="17" height="21" border="0" /></a>
                        </td>
                        <td align="left">&nbsp; <a href="lista_proveedor.asp?idPadre=<%=arrProveedorHijo(3,j)%>&idProveedor=<%=arrProveedorHijo(0,j)%>&idUsuarioCreador=<%=idUsuarioCreador%>&tip_usuario=<%=tip_usuario%>">
							<%=arrProveedorHijo(1,j)%></a>
                        </td>
                    </tr>
                
             <%			next
			 		end if	
					
			end if		%>
                
			 <%	next
			 end if	
			 %> 
            </table>
</div>
            <br />
           
     </fieldset>       
    	
    </td>
  </tr>
  <tr>
    
    <td align="left" colspan="2">
    <fieldset style="width:97%">
            <legend class="titulo_grupo" >Crear / Actualizar Proveedor:</legend> 
            <div style="width:100%; height:120; overflow:scroll-y">
        <form name="frm" id="frm" method="post" action="lista_proveedor.asp"  >
            <table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="contenido_tabla1">
    		<tr>
            	<td>Descripcion:</td>
                <td><input name="desProveedor" type="text" value="<%=desProveedor%>" size="20" maxlength="30" /></td>
                <td>Carpeta:</td>
                <td>
                	<select name="idProveedorPadre">
                    	<option value="0">[Seleccione]</option>
                        <%if isArray(arrProveedorPadre) then
							for i=0 to ubound(arrProveedorPadre,2)%>
                            <option value="<%=arrProveedorPadre(0,i)%>" <%if idProveedorPadre=arrProveedorPadre(0,i) then%> selected="selected" <%end if%> >
							<%=arrProveedorPadre(1,i)%></option>
                        <%  next
						  end if%>    
                    </select>
                </td>              
                <td>&nbsp;</td><td>&nbsp;</td>                
                <td><input type="submit" name="btnEnviar" value="Grabar proveedor" />    </td>
            </tr>        
           <%if idProveedor<>"" then
		   		sql = " select a.id_especialidad,b.des_especialidad,des_texto "
				sql = sql & " from proveedorxespecialidad a, especialidades b where a.id_especialidad  = b.id_especialidad "
				sql = sql & " and a.id_proveedor = "& idProveedor
				sql = sql & " order by 2"
						cn.open strUrl
						rs.open sql, cn
						if not rs.eof then
							arrProveedorEspecialidad = rs.GetRows()
						else
							arrProveedorEspecialidad = ""
						end if
						rs.close()
						cn.close	
			 end if						
		   %> 
           <%
		   if isArray(arrProveedorEspecialidad) then
		   	contador = ubound(arrProveedorEspecialidad,2) 
		   		for j=0 to ubound(arrProveedorEspecialidad,2)%>
           <tr>
           		<td><%=arrProveedorEspecialidad(1,j)%>
                <input type="hidden" name="idEspecialidadPro<%=j%>" value="<%=arrProveedorEspecialidad(0,j)%>" />
                </td>
                <td colspan="4">
                <textarea name="textoEspecialidad<%=j%>" cols="65" rows="5"><%=arrProveedorEspecialidad(2,j)%></textarea>
                </td>
               
                <td></td>
                <td></td>
           </tr>
           <%	next
		   end if%> 
           <tr>
           	<td>Agregar especialidad:</td>
            <td colspan="4">
            	<select name="idEspecialidadPro">
                <option value="">[Seleccione]</option>
                <%if isArray(arrEspecialidad) then
					for i=0 to ubound(arrEspecialidad,2)
				%>
                	<option value="<%=arrEspecialidad(0,i)%>" ><%=arrEspecialidad(1,i)%></option>
                <%
					next
				end if	
				%>
                </select>
                &nbsp;
                <input type="submit" value="+" title="Agregar Especialidad"  />
            </td>             
                <td></td>
                <td></td>
           </tr>
          </table>
         <input type="hidden" name="txtNumReg" value="<%=contador%>" />     
         <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
         <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
         <input type="hidden" name="idProveedor" value="<%=idProveedor%>" />
         <input type="hidden" name="idPadre" value="<%=idPadre%>" />
                    
        </form>
        </div>
        </fieldset>
    </td>
  </tr>
</table>
</center>
</body>
</html>
