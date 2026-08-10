<%@LCID = 1034%>
<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

idUsuarioCreador = request("idUsuarioCreador")
idCaso = request("idCaso")
idServicio = request("idServicio")

if instr(idCaso,"-")>0 then
	idCaso_bd = clng(mid(idCaso,5))
else
	idCaso_bd = clng(idCaso)
end if


if Request.ServerVariables("REQUEST_METHOD") = "POST" then

	impmontogarantia = SiVacio(request("impmontogarantia"),0)
	impMontoDeducible = SiVacio(request("impMontoDeducible"),0)
	desproveedor = request("desproveedor")
	impcosto = SiVacio(request("impcosto"),0)
	fchatencion = request("fchatencion")
	desserviciosol = request("desserviciosol")
	desserviciorec = request("desserviciorec")
	fchInformeMedico = request("fchInformeMedico")
	fchFactura = request("fchFactura")
	'MndDeducible =  SiVacio(request("MndDeducible"),0)
'	MndGarantia  =  SiVacio(request("MndGarantia"),0)
'	MndCosto     =  SiVacio(request("MndCosto"),0)
	MonedaDeducible = request("MonedaDeducible")
	MonedaGarantia = request("MonedaGarantia")
	MonedaCosto = request("MonedaCosto")
	anioServicio = request("anioServicio")

servicioId = request("servicioId")
servicioIdRecibido = request("servicioIdRecibido")
ImpCostoDolares = request("ImpCostoDolares")	
proveedorservicioId = request("proveedorservicioId")
	
		sql = " update servicio set impmontogarantia="&impmontogarantia&","
		sql = sql & "desproveedor='"&desproveedor&"',impcosto="&impcosto&","
		sql = sql & "fchatencion='"&FormateaFecha(fchatencion,"yyyymmdd") &"'"
		sql = sql & ",desserviciosol='"&desserviciosol&"'"
		sql = sql & ",desserviciorec='"&desserviciorec&"'"
		if fchInformeMedico <> "" then
		sql = sql & ",fchInformeMedico='"&FormateaFecha(fchInformeMedico,"yyyymmdd")&"'"
		else
		sql = sql & ",fchInformeMedico=NULL"
		end if
		if fchFactura <> "" then
		sql = sql & ",fchFactura='"&FormateaFecha(fchFactura,"yyyymmdd")&"'" 
		else
		sql = sql & ",fchFactura=NULL"
		end if
		sql = sql & ",MonedaDeducible='"& MonedaDeducible &"'"
		sql = sql & ",MonedaGarantia='"& MonedaGarantia & "'"
		sql = sql & ",MonedaCosto='"& MonedaCosto & "'"
		sql = sql & ",impMontoDeducible="&impMontoDeducible
		sql = sql & ",servicioId=" & SiVacio(servicioId,0)
		sql = sql & ",servicioIdRecibido=" & SiVacio(servicioIdRecibido,0)
		sql = sql & ",ImpCostoDolares=" & SiVacio(ImpCostoDolares,0)
		sql = sql & ",proveedorservicioId=" & SiVacio(proveedorservicioId,0)
		sql = sql & " where idServicio="& idServicio
		
		'response.write(sql)
		'response.end
		
		cn.open strUrl
		cn.execute(sql)
		cn.close
		
		idCaso = "00000" & idCaso
		idCaso = Right(idCaso,5)
		
'sql = "insert into suceso(dessuceso,codusuario,idCaso) values('Servicio modificado:"& idServicio & "-" & anioServicio &"','"&idUsuarioCreador&"','NCA-"&idCaso&"')"
'		cn.open strUrl
'		cn.execute(sql)
'		cn.close()
		
		impmontogarantia= ""
		desproveedor = ""
		impcosto = ""
		fchatencion = ""
		desserviciosol = ""
		desserviciorec = ""
		fchInformeMedico = ""
		fchFactura = ""
		idServicio = ""	
		anioServicio = ""
		MndDeducible =  ""
		MndGarantia  =  ""
		MndCosto     =  ""		
		anioServicio = ""
		servicioId = ""
		servicioIdRecibido = ""
		ImpCostoDolares = ""
	
	response.Redirect("registro_servicio.asp?idCaso="&idCaso_bd&"&idUsuarioCreador="&idUsuarioCreador&"")
	
end if

	sql = "select idServicio,fchatencion,desproveedor,fchInformeMedico,fchFactura,"
	sql = sql & "caso_idcaso,fchcreacion,desserviciorec,codUsuario,MonedaDeducible,impMontoDeducible,RIGHT(year(fchcreacion),2)"
	sql = sql & " ,s.servicioId,s.servicioIdRecibido,s.ImpCostoDolares," '12 13 14
	sql = sql & " (select servicioNombre from serviciocaso sc where sc.servicioId = s.servicioId ) des_servicioId,"
	sql = sql & " (select servicioNombre from serviciocaso sc where sc.servicioId = s.servicioIdRecibido) des_servicioIdRecibido"
	sql = sql & " from servicio s"
	sql = sql & " where caso_idcaso=" & idCaso_bd

	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
	arrServicios = rs.GetRows()
	end if
	rs.close()
	cn.close
	
	if idServicio <> "" then
	
		sql = "select idServicio,codUsuario,impMontoGarantia,desproveedor,impCosto,fchAtencion,fchInformeMedico,fchFactura,"
		sql = sql & "desServicioSol,desserviciorec,caso_idcaso,fchcreacion,impMontoDeducible,MonedaDeducible,MonedaGarantia,MonedaCosto,RIGHT(year(fchcreacion),2)"
		sql = sql & " ,s.servicioId,s.servicioIdRecibido,s.ImpCostoDolares," '17 18 19		
		sql = sql & " (select proveedorservicioContactoTelefono from proveedorservicio p where s.proveedorservicioId = p.proveedorservicioId) proveedorservicioContactoTelefono,"
		sql = sql & " (select proveedorservicioContactoCorreo from proveedorservicio p where s.proveedorservicioId = p.proveedorservicioId) proveedorservicioContactoCorreo,"			
		sql = sql & " (select servicioNombre from serviciocaso sc where sc.servicioId = s.servicioId ) des_servicioId,"
		sql = sql & " (select servicioNombre from serviciocaso sc where sc.servicioId = s.servicioIdRecibido) des_servicioIdRecibido"
		sql = sql & " from servicio s"
		sql = sql & " where idServicio=" & idServicio
	
		cn.open strUrl
		rs.open sql, cn
		if not rs.eof then
			
			
			
			impMontoGarantia = rs.fields(2)
			desproveedor = rs.fields(3)
			impCosto = rs.fields(4)
			fchAtencion = FormateaFecha(rs.fields(5),"ddmmyyyy")
			fchInformeMedico = FormateaFecha(rs.fields(6),"ddmmyyyy")
			fchFactura = FormateaFecha(rs.fields(7),"ddmmyyyy")
			if rs.fields(22)<>"" then
				desServicioSol = rs.fields(22)
			else
				desServicioSol = rs.fields(8)
			end if	
			if rs.fields(23)<>"" then
				desserviciorec = rs.fields(23)
			else
				desserviciorec = rs.fields(9)
			end if							
			impMontoDeducible = rs.fields(12)
			MonedaDeducible =  rs.fields(13)
			MonedaGarantia  =  rs.fields(14)
			MonedaCosto	 = rs.fields(15)
			anioServicio = rs.fields(16)
			servicioId			= rs.fields(17)
			servicioIdRecibido = rs.fields(18)
			ImpCostoDolares = rs.fields(19)
			proveedorservicioContactoTelefono = rs.fields(20)
			proveedorservicioContactoCorreo = rs.fields(21)
			
		end if
		rs.close()
		cn.close
	end if
	
	sql = "select idMoneda,desMoneda"
	sql = sql & " from moneda"
	sql = sql & " order by 2" 

	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
	arrMoneda = rs.GetRows()
	end if
	rs.close()
	cn.close	
	
		sql = "select servicioId,servicioNombre"
	sql = sql & " from serviciocaso"
	sql = sql & " order by 2" 

	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
	arrServicio = rs.GetRows()
	end if
	rs.close()
	cn.close	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>::: Network Care Assist ::: </title>
<LINK href="datePicker.css" type=text/css rel=stylesheet>
<LINK href="estilo.css" type=text/css rel=stylesheet>
<link type="text/css"rel="stylesheet"
     href="http://jquery-ui.googlecode.com/svn/tags/1.7/themes/redmond/jquery-ui.css" />   
  <script type="text/javascript"
     src="http://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.min.js"></script>
  <script type="text/javascript"
     src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.7.2/jquery-ui.min.js"></script>
  <script type="text/javascript">
      $(document).ready(function() {
          $('#fchAtencion').datepicker(
            {   dateFormat: 'dd/mm/yy',
                minDate: '-1Y',				
                maxDate: '+1Y',				
                changeMonth: true,
                changeYear: true,
                numberOfMonths: 1,
				required: true,  
                dayNamesMin: ['Do', 'Lu', 'Ma', 'Mi', 'Ju', 'Vi', 'Sa'],
                monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo',
                    'Junio', 'Julio', 'Agosto', 'Septiembre',
                    'Octubre', 'Noviembre', 'Diciembre'],
                monthNamesShort: ['Ene', 'Feb', 'Mar', 'Abr',
                    'May', 'Jun', 'Jul', 'Ago',
                    'Sep', 'Oct', 'Nov', 'Dic'] 
            });
		  $('#fchInformeMedico').datepicker(
            {   dateFormat: 'dd/mm/yy',
                minDate: '-1Y',				
                maxDate: '+1Y',				
                changeMonth: true,
                changeYear: true,
                numberOfMonths: 1,
                dayNamesMin: ['Do', 'Lu', 'Ma', 'Mi', 'Ju', 'Vi', 'Sa'],
                monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo',
                    'Junio', 'Julio', 'Agosto', 'Septiembre',
                    'Octubre', 'Noviembre', 'Diciembre'],
                monthNamesShort: ['Ene', 'Feb', 'Mar', 'Abr',
                    'May', 'Jun', 'Jul', 'Ago',
                    'Sep', 'Oct', 'Nov', 'Dic'] 
            });
			$('#fchFactura').datepicker(
            {   dateFormat: 'dd/mm/yy',
                minDate: '-1Y',				
                maxDate: '+1Y',				
                changeMonth: true,
                changeYear: true,
                numberOfMonths: 1,
                dayNamesMin: ['Do', 'Lu', 'Ma', 'Mi', 'Ju', 'Vi', 'Sa'],
                monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo',
                    'Junio', 'Julio', 'Agosto', 'Septiembre',
                    'Octubre', 'Noviembre', 'Diciembre'],
                monthNamesShort: ['Ene', 'Feb', 'Mar', 'Abr',
                    'May', 'Jun', 'Jul', 'Ago',
                    'Sep', 'Oct', 'Nov', 'Dic'] 
            });
      });
	  
		  //creamos una variable de tipo array para pasar y recuperar los datos 
var datos=new Array(); 

function AbreVentanaModal(){ 
/*datos[0]=DatoPadre1.value; 
datos[1]=DatoPadre2.value; 
datos[2]=DatoPadre3.value; */

//aqui paso los datos a la ventana hija 
datos=showModalDialog('bsqProveedor.asp?idCaso=<%=idCaso%>',datos,'dialogWidth:900px;dialogHeight:350px;status:no;resizable:yes'); 

//aqui recuepero datos de la ventana hija 
frm.desProveedor.value=datos[0]; 
frm.proveedorContactoTelefono.value=datos[1]; 
frm.proveedorContactoCorreo.value=datos[2]; 
frm.proveedorservicioId.value= datos[3];
} 	  
  </script> 
</head>
<body bgcolor="#e9e9e9" style="font-size:11px;" leftmargin="0">
<form name="frm" method="post" action="actualizar_servicio.asp">    
 <table width="100%" border="0" cellspacing="1" cellpadding="1" class="contenido_tabla1">
  <tr>
  	<td>
    	<fieldset style="width:98%">
          <legend class="titulo_grupo" >Datos del servicio:</legend>  
<table border="0" cellspacing="1" cellpadding="1" class="contenido_tabla1">
 <tr>
 <td>#Servicio
     </td>
     <td nowrap="nowrap"> <input name="idServicio" type="text" size="4" readonly="readonly" value="<%=idServicio%>" />
      <input name="guionServicio" type="text" size="1" readonly="readonly" value="-"  style="width:10px" />
      <input name="anioServicio" type="text" size="2" readonly="readonly" value="<%=anioServicio%>"  style="width:20px" />&nbsp;</td>
      <td>Proveedor</td>
      <td><input name="desProveedor" type="text" size="20"  value="<%=desProveedor%>" /><input type="button" value="..." id="btnProveedor" onclick="AbreVentanaModal()" /></td>
      <td>Contacto:</td>
      <td><input name="proveedorContactoTelefono" value="<%=proveedorservicioContactoTelefono%>" type="text" size="10" readonly="readonly" /></td>
      <td><input name="proveedorContactoCorreo" value="<%=proveedorservicioContactoCorreo%>" type="text" size="20" readonly="readonly" /></td>
 </tr>
            <tr>
      <td>Deducible</td>
    <td><input type="text" name="impMontoDeducible" size="5" value="<%=replace(SiVacio(impMontoDeducible,0), ",", ".")%>"/>
      <input type="text" name="MonedaDeducible" size=5 value="<%=MonedaDeducible%>" maxlength="10" /></td>
    <td>Garantia</td>
    <td><input type="text" name="impMontoGarantia" size="5"  value="<%=replace(impMontoGarantia, ",", ".")%>"/>
      <input type="text" name="MonedaGarantia" value="<%=MonedaGarantia%>" size="5" maxlength="10" /></td>
   
    
    <td>Costo </td>
    <td><input type="text" name="impCosto" size="5" value="<%=replace(impCosto, ",", ".")%>"/>
    		<input type="text" name="MonedaCosto" size="5" value="<%=MonedaCosto%>" maxlength="10" />
    </td>
  </tr>
  <tr>
    <td>F. Atenci&oacute;n</td>
    <td><input type="text" name="fchAtencion" id="fchAtencion" size="8"  value="<%=fchAtencion%>"/></td>
    <td>Informe Medico</td>
    <td><input type="text" name="fchInformeMedico" id="fchInformeMedico" size="10" value="<%=fchInformeMedico%>"/></td>
    <td>Recepci&oacute;n Factura</td>
    <td><input type="text" name="fchFactura" id="fchFactura" size="10" value="<%=fchFactura%>"/></td>
    <td>&nbsp;</td>
    
  </tr>
 <tr>
    <td>Servicio Solicitado</td>   
    <td>
    <input name="desServicioSol" type="text" value="<%=desServicioSol%>" size="25" maxlength="40"/>    	
    </td>
    <td>Servicio Recibido</td>
    <td><input name="desServicioRec" type="text" value="<%=desServicioRec%>" size="25" maxlength="40"/>
    </td>
    <td>Costo USD</td>
    <td><input type="text" name="ImpCostoDolares" size="5" value="<%=ImpCostoDolares%>"/></td>
    <td><input type="submit" name="btnGrabar" value="Grabar Servicio" /></td>    
  </tr>
</table>  
        </fieldset>    
    </td>
  </tr>
  <tr>
  	<td>
    <fieldset style="width:98%">
            <legend class="titulo_grupo" >Servicios:</legend> 
             <div style="width:100%; height:180px; overflow-y:auto;">
            <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr bgcolor="#FF4935">
                <td align="center" class="titulo_tabla" width="16%" style="border-right:#000 1px solid;border-left:#000 1px solid;">&nbsp;Id Servicio</td>      
                <td align="center" class="titulo_tabla" width="9%" style="border-right:#000 1px solid">&nbsp;Atenci&oacute;n</td>
                <td align="center" class="titulo_tabla" width="21%" style="border-right:#000 1px solid">&nbsp;Proveedor</td>
                <td align="center" class="titulo_tabla" width="25%" style="border-right:#000 1px solid">&nbsp;Servicio</td>
				<td align="center" class="titulo_tabla" width="11%" style="border-right:#000 1px solid">&nbsp;Deducible</td>				
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">&nbsp;Inf. Medico</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">&nbsp;Factura</td>
              </tr>
           <%if isArray(arrServicios) then
		   colorfila = 0
		   		for i=0 to ubound(arrServicios,2)
			if colorfila = 0 then
					color= "#FFFFFF"
			'colorfila=1
'			else
'					color="#COCOCO"
'					colorfila=0
			end if
			
			%>   
              <tr class="contenido_tabla1" bgcolor="<%=color%>">
                <td align="center">
         <a href="actualizar_servicio.asp?idServicio=<%=arrServicios(0,i)%>&idCaso=<%=arrServicios(5,i)%>&idUsuarioCreador=<%=idUsuarioCreador%>">
                	<%=arrServicios(0,i)%>-<%=arrServicios(11,i)%>
                </a>
                </td>
                <td>&nbsp;<%=FormateaFecha(arrServicios(1,i),"ddmmyyyy")%></td>
                <td>&nbsp;<%=ucase(arrServicios(2,i))%></td>
                <td>&nbsp;<%=SiVacio(ucase(arrServicios(16,i)),ucase(arrServicios(7,i)))%></td>   
				<td>&nbsp;<%=ucase(arrServicios(9,i))%>&nbsp;<%=FormatNumber(arrServicios(10,i),2)%></td>					
                <td>&nbsp;<%=FormateaFecha(arrServicios(3,i),"ddmmyyyy")%></td>
                <td>&nbsp;<%=FormateaFecha(arrServicios(4,i),"ddmmyyyy")%></td>

              </tr>
             <%
			 	next
			 end if	
			 %> 
            </table>
            <br />
           </div>
     </fieldset>   
    </td>
  </tr>
  </table>
   <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
   <input type="hidden" name="idCaso" value="<%=idCaso%>" />
      <input type="hidden" name="proveedorservicioId" value="<%=proveedorservicioId%>" />
   
   
  </form>
</body>
</html>
