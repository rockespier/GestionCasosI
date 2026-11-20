<%@LCID = 1034%>
<!--#include file="main.asp"-->
<!--#include file="combos.asp"-->
<%
on error resume next

Set rs = Server.CreateObject("ADODB.RECORDSET")

idUsuarioCreador = request("idUsuarioCreador")
idCaso = request("idCaso")
empresacasoserviciomax = request("empresacasoserviciomax")
cantidadservicio = SiVacio(request("cantidadservicio"),0)
resultado = 0

if instr(idCaso,"-")>0 then
	idCaso_bd = clng(mid(idCaso,5))
else
	idCaso_bd = clng(idCaso)
end if

'MndDeducible =  SiVacio(request("MndDeducible"),1)
'MndGarantia  =  SiVacio(request("MndGarantia"),1)
'MndCosto     =  SiVacio(request("MndCosto"),1)

if Request.ServerVariables("REQUEST_METHOD") = "POST" then

impmontogarantia = SiVacio(request("impmontogarantia"),0)
desproveedor = request("desproveedor")
impcosto = SiVacio(request("impcosto"),0)
fchatencion = request("fchatencion")
desserviciosol = request("desserviciosol")
desserviciorec = request("desserviciorec")
fchInformeMedico = request("fchInformeMedico")
fchFactura = request("fchFactura")

'MndDeducible =  SiVacio(request("MndDeducible"),0)
'MndGarantia  =  SiVacio(request("MndGarantia"),0)
'MndCosto     =  SiVacio(request("MndCosto"),0)

MonedaDeducible = request("MonedaDeducible")
MonedaGarantia = request("MonedaGarantia")
MonedaCosto = request("MonedaCosto")
impMontoDeducible = SiVacio(request("impMontoDeducible"),0)

servicioId = request("servicioId")
servicioIdRecibido = request("servicioIdRecibido")
ImpCostoDolares = request("ImpCostoDolares")
proveedorservicioId = request("proveedorservicioId")


	sql = "insert into servicio(caso_idcaso,impmontogarantia,"
	sql = sql & "desproveedor,impcosto,fchatencion,desserviciosol,"
	sql = sql & "desserviciorec,codusuario,fchcreacion,fchInformeMedico,fchFactura,"
	sql = sql & "MonedaDeducible,MonedaGarantia,MonedaCosto,impMontoDeducible, " 
	sql = sql & "servicioId,servicioIdRecibido,ImpCostoDolares,proveedorservicioId"
	sql = sql & ")"
	sql = sql & "values ("&idCaso_bd&","&impmontogarantia&",'"&desproveedor&"',"&impcosto&",'"&FormateaFecha(fchatencion,"yyyymmdd") &"','"&desserviciosol
	sql = sql &"','"&desserviciorec&"','"&idUsuarioCreador&"',DATEADD(hour,-5,GETUTCDATE())"
	if fchInformeMedico <> "" then
	sql = sql & ",'"&FormateaFecha(fchInformeMedico,"yyyymmdd")&"'"
	else
	sql = sql & ",NULL"
	end if
	if fchFactura <> "" then
	sql = sql & ",'"& FormateaFecha(fchFactura,"yyyymmdd")&"'"
	else
	sql = sql & ",NULL"	
	end if
	sql = sql & ",'" & MonedaDeducible &"','"&MonedaGarantia&"','"&MonedaCosto&"',"&impMontoDeducible
	sql = sql & "," & Sivacio(servicioId,0) & "," & SiVacio(servicioIdRecibido,0) & "," & SiVacio(ImpCostoDolares,0)
	sql = sql & "," & Sivacio(proveedorservicioId,0) & ")"
	
	'response.write(sql)
	'response.end
	
	cn.open strUrl
	cn.execute(sql)
	cn.close

'sql = "SELECT SCOPE_IDENTITY()"	
'	cn.open strUrl
'	rs.open sql, cn
'	if not rs.eof then
'		idServicio = rs.fields(0)
'	end if
'	rs.close()
'	cn.close
		idCaso = "00000" & idCaso
		idCaso = Right(idCaso,5)
'		sql = "insert into suceso(dessuceso,codusuario,idCaso) values('Servicio registrado:"& idServicio & "-" & right(year(date()),2) &"','"&idUsuarioCreador&"','NCA-"&idCaso&"')"

		sql= "exec TipoCaso_Obtener "& idCaso_bd
		cn.open strUrl
		cn.execute(sql)
		cn.close()
	
end if

	sql = "select s.idServicio,s.fchatencion,s.desproveedor,s.fchInformeMedico,s.fchFactura,"
	sql = sql & "s.caso_idcaso,s.fchcreacion,s.desserviciorec,s.codUsuario,s.MonedaDeducible,s.impMontoDeducible,RIGHT(year(s.fchcreacion),2),"
	sql = sql & "empresacasodiasmax,empresacasoserviciomax," '12 13
    sql = sql & "DATEDIFF ( day , c.fchcreacion , DATEADD(hour,-5,GETUTCDATE()) ) dif_day,"
	sql = sql & " s.servicioId,s.servicioIdRecibido,s.ImpCostoDolares," '15 16 17
	sql = sql & " (select servicioNombre from serviciocaso sc where sc.servicioId = s.servicioId ) des_servicioId,"
	sql = sql & " (select servicioNombre from serviciocaso sc where sc.servicioId = s.servicioIdRecibido) des_servicioIdRecibido"
	sql = sql & " from servicio s, caso c, empresa e "
	sql = sql & " where caso_idcaso="& idCaso_bd
	sql = sql & " and s.caso_idcaso = c.idCaso"
	sql = sql & " and c.idEmpresa = e.idEmpresa"



	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
	arrServicios = rs.GetRows()
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

	sql = ""
	sql = sql & " select id_opcion from accesos"
	sql = sql & " where id_opcion='correo_fact'"
	sql = sql & " and id_usuario='"& idUsuarioCreador &"'"
	cn.open strUrl
	rs.open sql, cn

	if not rs.eof then
		arrAcceso = rs.GetRows()
	end if
	rs.close()
	cn.close	
	
'	sql = "select idMoneda,desMoneda"
'	sql = sql & " from moneda"
'	sql = sql & " order by 2" 
'
'	cn.open strUrl
'	rs.open sql, cn
'	if not rs.eof then
'	arrMoneda = rs.GetRows()
'	end if
'	rs.close()
'	cn.close	

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>::: Med Global Assist ::: </title>
<LINK href="estilo.css" type=text/css rel=stylesheet>
<link type="text/css"rel="stylesheet"
     href="http://jquery-ui.googlecode.com/svn/tags/1.7/themes/redmond/jquery-ui.css" />   
  <script type="text/javascript"
     src="http://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.min.js"></script>
  <script type="text/javascript"
     src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.7.2/jquery-ui.min.js"></script>
     
<script type="text/javascript" src="js/apprise-1.5.min.js"></script>
<link rel="stylesheet" type="text/css" href="js/apprise.min.css"/>
<link rel="stylesheet" type="text/css" href="datePicker.css"/>
     
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


$(document).ready(function(){  
	$('#frm').submit(function(e) {



		 if( $("#grabar").val() == '1')
        {          
          //alert("No se permite registrar mas servicios. \nContacte a al administrador.");
		  apprise("No se permite registrar más servicios. Contacte al administrador.");
   	 	  e.preventDefault();
        }		
		
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

function EnviarCorreo()
{ window.open("EnviarCorreo.asp?idCaso=<%=idCaso%>","Correo","toolbar=no, scrollbars=no, resizable=no, top=200, left=500, width=400, height=400");
	}

  </script> 
</head>
<body bgcolor="#e9e9e9" style="font-size:11px;" leftmargin="0">
<form id="frm" name="frm" method="post" action="registro_servicio.asp" >    
 <table width="100%" border="0" cellspacing="1" cellpadding="1" class="contenido_tabla1">

  <tr>
  	<td align="left">
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
      <td><input name="desProveedor" type="text" size="20"/><input type="button" value="..." id="btnProveedor" onclick="AbreVentanaModal()" /></td>
      <td>Contacto:</td>
      <td><input name="proveedorContactoTelefono" type="text" size="10" readonly="readonly" /></td>
      <td><input name="proveedorContactoCorreo" type="text" size="20" readonly="readonly" /></td>
 </tr>
            <tr>
       
    <td>Deducible</td>
    <td><input type="text" name="impMontoDeducible" size="5" value="0"/> <%ComboMoneda "MonedaDeducible", ""%></td>
    <td>Garantia</td>
    <td><input type="text" name="impMontoGarantia" size="5" required />
    <%ComboMoneda "MonedaGarantia", ""%>
      </td>
   
    
    <td>Costo </td>
    <td><input type="text" name="impCosto" size="5"/>
    <%ComboMoneda "MonedaCosto", ""%>
    </td>
  </tr>
  <tr>
    
   
    <td>F. Atenci&oacute;n</td>
    <td><input type="text" name="fchAtencion" id="fchAtencion" size="8" required /></td>
    <td>Informe Medico</td>
    <td><input type="text" name="fchInformeMedico" id="fchInformeMedico" size="10"/></td>
    <td>Recepci&oacute;n Factura</td>
    <td><input type="text" name="fchFactura" id="fchFactura" size="10"/></td>
    <td><%if isArray(arrAcceso) then%>
    	<a href="JavaScript:EnviarCorreo()"><img src="images/mail_send.png" width="25" height="25" />[Enviar]</a>
        <%end if%>
    </td>
    
  </tr>
  <tr>
    <td>Servicio Solicitado</td>   
    <td>
    	<%ComboServicioSolicitado "desServicioSol", desServicioSol%>
    </td>
    <td>Servicio Recibido</td>
    <td>
      <input name="desServicioRec" type="text" value="<%=desServicioRec%>" size="25" maxlength="40"/>
    </td>
    <td>Costo USD</td>
    <td><input type="text" name="ImpCostoDolares " size="5"/></td>
    <td><input type="submit" name="btnGrabar" value="Grabar Servicio" /></td>    
  </tr>
</table>

        </fieldset>    
    </td>
  </tr>
  <tr>
  	<td align="left">
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
				
				empresacasodiasmax = SiVacio(arrServicios(12,0),0)
				empresacasoserviciomax = SiVacio(arrServicios(13,0),0)
				difday = arrServicios(14,0)
				
		   		for i=0 to ubound(arrServicios,2)
					cantidadservicio = cantidadservicio + 1 
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
                <td>&nbsp;<%=SiVacio(ucase(arrServicios(19,i)),ucase(arrServicios(7,i)))%></td>    
				<td>&nbsp;<%=ucase(arrServicios(9,i))%>&nbsp;<%=FormatNumber(arrServicios(10,i),2)%></td>				
                <td>&nbsp;<%=FormateaFecha(arrServicios(3,i),"ddmmyyyy")%></td>
                <td>&nbsp;<%=FormateaFecha(arrServicios(4,i),"ddmmyyyy")%></td>

              </tr>
             <%
			 	next
				
					if clng(difday) > clng(empresacasodiasmax) then
					
					resultado = 1
					
					else
									
						if clng(cantidadservicio) >= clng(empresacasoserviciomax) then
							resultado = 1
						else
							resultado = 0
						end if
						
					end if	
				
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
   <input type="hidden" name="empresacasoserviciomax" value="<%=empresacasoserviciomax%>" />
   <input type="hidden" name="cantidadservicio" value="<%=cantidadservicio%>" />
   <input type="hidden" id="grabar" name="grabar" value="<%=resultado%>" />
         <input type="hidden" name="proveedorservicioId" value="<%=proveedorservicioId%>" />
  </form>
</body>
</html>
