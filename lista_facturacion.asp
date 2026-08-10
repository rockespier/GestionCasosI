<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

fchInicio = request("fchInicio")
fchFin = request("fchFin")

if Request.ServerVariables("REQUEST_METHOD") = "POST" then

sql = ""
sql = sql &" exec FacturacionCaso_Obtener "
sql = sql &" '"&FormateaFecha(fchInicio,"yyyymmdd") &"','"&FormateaFecha(fchFin,"yyyymmdd") &"',NULL"

'response.Write(sql)
'response.End()

	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
	arrCasos = rs.GetRows()
	end if
	rs.close()
	cn.close
	
end if	


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>::: Innova Doctors ::: </title>
<LINK href="estilo.css" type=text/css rel=stylesheet>
<link type="text/css"rel="stylesheet"
     href="http://jquery-ui.googlecode.com/svn/tags/1.7/themes/redmond/jquery-ui.css" />   
  <script type="text/javascript"
     src="http://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.min.js"></script>
  <script type="text/javascript"
     src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.7.2/jquery-ui.min.js"></script>
     
<script type="text/javascript" src="js/apprise-1.5.min.js"></script>
<link rel="stylesheet" type="text/css" href="js/apprise.min.css"/>
     
  <script type="text/javascript">
      $(document).ready(function() {
          $('#fchInicio').datepicker(
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
		  $('#fchFin').datepicker(
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
</script>	  
</head>
<body bgcolor="#e9e9e9" style="font-size:11px;" leftmargin="0">
 <center>
    <form name="frm" action="lista_facturacion.asp" method="post">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="2" class="titulo" align="left">
    <br />
    Consulta de facturación
    <br />
	<br />
</td>
</tr>
  <tr>
    <td colspan="2" >

    	<fieldset >
            <legend class="titulo_grupo" >Fecha de Cierre:</legend>            
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
              	<td>Desde:</td>
                <td>
					<input type="text" name="fchInicio" id="fchInicio" size="8" value="<%=fchInicio%>" required />					                  
                </td>
                <td>Hasta:</td>
                <td>
					<input type="text" name="fchFin" id="fchFin" size="8" value="<%=fchFin%>" required />					                  
                </td>
                <td>&nbsp; </td>
                <td width="10%" align="left"><input type="submit" value="Buscar" /></td>
              </tr>
            </table>          
        </fieldset>
        <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
        <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
      
    </td>    
  </tr>
  <tr>
  	<td colspan="2">
    	<fieldset>
            <legend class="titulo_grupo" >Casos:</legend> 
            <div style="width:100%; height:260px; overflow-y:auto;">
            <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr bgcolor="#FF4935">
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid;border-left:#000 1px solid;">&nbsp;Id</td>
                <td align="center" class="titulo_tabla" width="30%" style="border-right:#000 1px solid">&nbsp;Pasajero</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">&nbsp;Ref. Externa</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">&nbsp;Tipo</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">&nbsp;Fee</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">&nbsp;Transf. Banco</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">&nbsp;Costo Medico</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">&nbsp;Total</td>                
              </tr>
           <%if isArray(arrCasos) then
		   		nro_registro = ubound(arrCasos,2) + 1						   
				for i=0 to ubound(arrCasos,2)
					total = cdbl(total) + cdbl(arrCasos(4,i)) + cdbl(arrCasos(5,i)) + cdbl(arrCasos(6,i))
			%>   
              <tr class="contenido_tabla1" bgcolor="#fff">
                <td>&nbsp;<a href="invoice_xls.asp?idCaso=<%=arrCasos(0,i)%>&fchInicio=<%=fchInicio%>&fchFin=<%=fchFin%>" target="_blank">INV-<%=arrCasos(0,i)%></a></td>
                <td>&nbsp;<%=ucase(arrCasos(1,i))%></td>
                <td>&nbsp;<%=arrCasos(2,i)%></td>
                <td>&nbsp;<%=ucase(arrCasos(3,i))%></td>                
                <td align="right">&nbsp;<%=formatnumber(arrCasos(4,i),2)%></td>
                <td align="right">&nbsp;<%=formatnumber(arrCasos(5,i),2)%></td>
                <td align="right">&nbsp;<%=formatnumber(arrCasos(6,i),2)%></td>
                <td align="right">&nbsp;<%=formatnumber(total,2)%></td>                                         
              </tr>
             <%
			 	total = 0
			 	next
			 end if	
			 %> 
            </table>
            <br />
           </div>
     </fieldset>
    </td>
  </tr>
   <tr>
    <td class="contenido_tabla1">&nbsp;# Registros:<%=nro_registro%>    
         <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
          <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
    </td>
    <td  class="contenido_tabla1" align="right">
    <%if cint(nro_registro) > 0 then%>
   	  <a href="lista_facturacion_xls.asp?fchInicio=<%=fchInicio%>&fchFin=<%=fchFin%>" target="_blank">
      <img src="images/excel.png" width="27" height="25" /> [Exportar]</a>
    <%end if%>  
    </td>
  </tr>
</table>
</form>
</body>
</html>
