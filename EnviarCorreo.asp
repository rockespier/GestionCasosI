<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

idUsuarioCreador = request("idUsuarioCreador")
tip_usuario = request("tip_usuario")
idCaso = request("idCaso")



msj = request("msj")
	if msj="ok" then
		EnvioResultado="Correo Enviado Correctamente!"
	else
		if msj="ko" then
		EnvioResultado="Ocurrio un error al enviar el correo. Intente nuevamente luego."
		end if
	end if


		sql = ""
		sql = sql & " select s.idServicio,desProveedor,fchFactura,ImpCostoDolares,"
		sql = sql & " desserviciorec,"
		sql = sql & " (select servicioNombre from serviciocaso sc where sc.servicioId = s.servicioIdRecibido) des_servicioIdRecibido,idRefExterna,c.desnombre"
		sql = sql & "  from servicio s, caso c"
		sql = sql & " where Caso_idCaso="& idCaso
		sql = sql & "  and s.Caso_idCaso = idcaso and fchFactura is not null order by fchFactura"
		
		cn.open strUrl
		rs.open sql, cn
		if not rs.eof then
			arrDatos = rs.GetRows()
		end if
		rs.close()
		cn.close
		
		strMensaje = ""
		if isArray(arrDatos) then
				strMensaje = " " & vbCrLf
			for i=0 to ubound(arrDatos,2)
				strMensaje = strMensaje & "Servicio: " & arrDatos(0,i) & vbCrLf
				strMensaje = strMensaje & "Descripción: " & SiVacio(arrDatos(5,i),arrDatos(4,i)) & vbCrLf
				strMensaje = strMensaje & "Proveedor: " & arrDatos(1,i) & vbCrLf
				strMensaje = strMensaje & "Recepción Factura: " & FormateaFecha(arrDatos(2,i),"ddmmyyyy") & vbCrLf
				strMensaje = strMensaje & "Costo USD: " & arrDatos(3,i) & vbCrLf
				strMensaje = strMensaje & "Ref.Externa:" & arrDatos(6,i) & vbCrLf				
				strMensaje = strMensaje & "Pasajero:" & arrDatos(7,i) & vbCrLf
				strMensaje = strMensaje & "----------------------------" & vbCrLf
			next
		end if
		strPara="roberto@rtres.net"
		strAsunto="Facturas Caso NCA-"& idCaso
		
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>::: Network Care Assist ::: </title>
<LINK href="estilo.css" type=text/css rel=stylesheet>
<style>
.mensaje{ color:#F00}
</style>
</head>

<body bgcolor="#e9e9e9" style="font-size:11px;">


<form name="frm" id="frm" method="post" action="correo.asp">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>Para:&nbsp;</td>
  </tr>
  <tr>  
    <td>&nbsp;<input name="strPara" type="text" size="60" value="<%=strPara%>"  style="width:100%" /></td>
  </tr>
  <tr>
    <td>Asunto:&nbsp;</td>
  </tr>  
  <tr>
    <td>&nbsp;<input name="strAsunto" type="text" size="60" value="<%=strAsunto%>"  style="width:100%" /></td>
  </tr>
  <tr>
    <td>Mensaje:&nbsp;</td>
  </tr>
  <tr>  
    <td>
	<textarea name="strMensaje" cols="45" rows="10">
	<%=strMensaje%>
    </textarea>
    </td>
  </tr>
  <tr>
    <td>
    <input type="hidden" name="idCaso" value="<%=idCaso%>" />
    </td>
  </tr>
  <tr>  
    <td>&nbsp;<input name="btnEnviar" type="submit" value="Enviar" /></td>
  </tr>
  <tr>
  	<td colspan="2">
    	<span class="mensaje">	<%=EnvioResultado%> </span>
    </td>
  </tr>
</table>
</form>

</body>
</html>
