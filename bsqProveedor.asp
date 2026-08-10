<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

idUsuarioCreador = request("idUsuarioCreador")
tip_usuario = request("tip_usuario")
idCaso = request("idCaso")

'if Request.ServerVariables("REQUEST_METHOD") = "POST" then

'Ciudad del Caso
sql ="select des_ciudad from caso where idcaso=" & idCaso
		cn.open strUrl
		rs.open sql, cn
		if not rs.eof then
			strDesCiudad = rs(0)
		end if
		rs.close()
		cn.close


sql = "select proveedorservicioId,proveedorservicioRazonSocial,"
		sql = sql & " proveedorservicioCiudadAtencion,"
		sql = sql & " proveedorservicioContactoTelefono,"
		sql = sql & " proveedorservicioContactoFax,"
		sql = sql & " proveedorservicioContactoWeb,"
		sql = sql & " proveedorservicioContactoCorreo,"
		sql = sql & " proveedorservicioHorario,"
		sql = sql & " proveedorservicioCalificacion,"
		sql = sql & " proveedorservicioEspecialidad,"
		sql = sql & " proveedorservicioActivo"
		sql = sql & " from proveedorservicio"
		sql = sql & " where proveedorservicioActivo=1"
		if strDesCiudad <>"" then
		sql = sql & " and proveedorservicioCiudadAtencion like '%"& trim(strDesCiudad) &"%'"		
		end if
		sql = sql & " order by proveedorservicioRazonSocial"

		cn.open strUrl
		rs.open sql, cn
		if not rs.eof then
			arrDatos = rs.GetRows()
		end if
		rs.close()
		cn.close
'End if


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>::: Med Global Assist ::: </title>
<LINK href="estilo.css" type=text/css rel=stylesheet>
<SCRIPT> 
//creamos una variable de tipo array para recuperar y devolver los datos 


//aqui recuperamos los datos de la ventana padre 
/*function RecuperaDatos(){ 
datos=dialogArguments; 
DatoHijo1.value=datos[0]; 
DatoHijo2.value=datos[1]; 
DatoHijo3.value=datos[2]; 
} */

//aqui le devolvemos los datos a la ventana padre 
function Devuelve(nombre,telefono,web,codigo){ 
var datos=new Array(); 

datos[0]=nombre; 
datos[1]=telefono; 
datos[2]=web; 
datos[3]=codigo; 

returnValue = datos; 
window.close(); 
} 
</SCRIPT> 

</head>
<body bgcolor="#e9e9e9">
<form name="frm" id="frm" method="post" action="bsqProveedor.asp">
 <table border="0" align="center" cellpadding="0" cellspacing="0" class="contenido_tabla1" width="100%">
 <tr>
 	<td colspan="2" align="center">Para buscar proveedores de otra ciudad, debe modificar el caso.</td>
	<td>&nbsp;</td>
 </tr>
 </table>
 <br />
<fieldset>
            <legend class="titulo_grupo" >Proveedores de Servicio (<%=ucase(strDesCiudad)%>):</legend> 
           
            <table border="0" align="center" cellpadding="0" cellspacing="0" class="contenido_tabla1" width="100%">
              <tr bgcolor="#FF4935">
                <td align="center" class="titulo_tabla" width="26%" style="border-right:#000 1px solid;border-left:#000 1px solid;">Razon Social</td>
                <td align="center" class="titulo_tabla" width="30%" style="border-right:#000 1px solid">Ciudad</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">Telefono</td>
                <td align="center" class="titulo_tabla" width="17%" style="border-right:#000 1px solid">Web</td>
                <td align="center" class="titulo_tabla" width="17%" style="border-right:#000 1px solid">Correo</td>
              </tr>
           </table>
            <div style="width:100%; height:120px; overflow:scroll-y">   
            <table border="0" align="center" cellpadding="0" cellspacing="0" class="contenido_tabla1" width="100%">
           <%if isArray(arrDatos) then
				   colorfila = 0
		   		for i=0 to ubound(arrDatos,2)
			if colorfila = 0 then
					color= "#FFFFFF"
			end if
			%>   
              <tr class="contenido_tabla1"  bgcolor="<%=color%>">
                <td align="center" width="26%" >&nbsp;<a href="Javascript:Devuelve('<%=ucase(arrDatos(1,i))%>','<%=arrDatos(3,i)%>','<%=arrDatos(6,i)%>','<%=arrDatos(0,i)%>')">
				<%=ucase(arrDatos(1,i))%></td>               
                <td align="center" width="30%" >&nbsp;<%=ucase(arrDatos(2,i))%></td>               
                <td align="center" width="10%" >&nbsp;<%=ucase(arrDatos(3,i))%></td>               
                <td align="center" width="17%" >&nbsp;<%=arrDatos(5,i)%></td>               
                <td align="center" width="17%" >&nbsp;<%=arrDatos(6,i)%></td>               
              </tr>
              	<%
			 	next
			 end if	
			 %> 
            </table>
			</div>
            <br />           
     </fieldset>
</form>
</body>
</html>
