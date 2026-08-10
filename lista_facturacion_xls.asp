<%
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=facturacion"&date()&".xls" 
%>
<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

fchInicio = request("fchInicio")
fchFin = request("fchFin")

sql = ""
sql = sql &" exec FacturacionCaso_Obtener "
sql = sql &" '"&FormateaFecha(fchInicio,"yyyymmdd") &"','"&FormateaFecha(fchFin,"yyyymmdd") &"',NULL"

	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
	arrCasos = rs.GetRows()
	end if
	rs.close()
	cn.close
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>::: Innova Doctors ::: </title>
<LINK href="http://gestiondecasos.innovadoctors.com/estilo.css" type=text/css rel=stylesheet>   
</head>
<body>
 <center>
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
  	<td colspan="2">
            <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr bgcolor="#DF7401">
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
              <tr class="contenido_tabla1" bgcolor="<%=color%>">
                <td>&nbsp;INV-<%=arrCasos(0,i)%></td>
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
    </td>
  </tr>
</table>
</body>
</html>
