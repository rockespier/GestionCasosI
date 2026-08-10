<%
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=invoice-NCA"&idCaso& "-"&date()&".xls" 
%>
<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

fchInicio = request("fchInicio")
fchFin = request("fchFin")
idCaso = request("idCaso")

sql = ""
sql = sql &" exec FacturacionCaso_Obtener "
sql = sql &" '"&FormateaFecha(fchInicio,"yyyymmdd") &"','"&FormateaFecha(fchFin,"yyyymmdd") &"',"& idCaso

	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
		idCaso 			= rs(0)
		desNombre 		= rs(1)
		idRefExterna 	= rs(2)
		destipoCaso 	= rs(3)
		Fee 			= SiVacio(rs(4),0)
		trans 			= SiVacio(rs(5),0)
		costo 			= SiVacio(rs(6),0)
		desEmpresa 		= rs(7)
		desDireccion 	= rs(8)
		desRuc 			= rs(9)
		desGiro 		= rs(10)
	end if
	rs.close()
	cn.close
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>::: Network Care Assist ::: </title>
<LINK href="http://networkcareassist.com/gestion-de-casos/estilo.css" type=text/css rel=stylesheet>   
<style>
.izquierda{
	border-left:solid 2px #000;
}
.derecha{
	border-right:solid 2px #000;
}
.abajo{
	border-bottom:solid 2px #000;
}
.abaizq{
	border-bottom:solid 2px #000;
	border-left:solid 2px #000;
	}
.abader{
	border-bottom:solid 2px #000;
	border-right:solid 2px #000;
	}	
</style>
</head>
<body>
 <center>
 	<table width="720" border="0" cellpadding="0" cellspacing="0" class="contenido_tabla1">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="100">&nbsp;</td>
    <td width="100">&nbsp;</td>
    <td width="100">&nbsp;</td>
    <td width="100">&nbsp;</td>
    <td width="100">&nbsp;</td>
    <td width="100">&nbsp;</td>
    <td width="100">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td colspan="4" rowspan="8"><img src="http://networkcareassist.com/gestion-de-casos/images/logo1intregalsupporttraveler.jpg" width="272" height="153" /></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td class="titulo">INVOICE</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>INVOICE MK#</td>
    <td style="text-align:left"><%=IdCaso%></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td colspan="2">TE: <%=FormatDateTime(Date(),1)%></td>
    </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>TO:</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>FOR:</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td colspan="3"><%=desEmpresa%></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td colspan="2">Service Agrement</td>
    </tr>
  <tr>
    <td>&nbsp;</td>
    <td colspan="3"><%=desDireccion%></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td colspan="3"><%=desRuc%> &nbsp; <%=desGiro%></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td colspan="5" style="text-align: center; border:solid 2px #000" class="titulo_grupo">DESCRIPTION</td>
    <td colspan="2" style="text-align: center; border:solid 2px #000; border-left:none" class="titulo_grupo">AMOUNT</td>
    </tr>
  <tr>
    <td>&nbsp;</td>
    <td style="border-left:solid 2px #000;">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td colspan="3" class="izquierda" style="background-color:#EEE"><%=desNombre%></td>
    <td>Ref.</td>
    <td class="derecha"><%=idRefExterna%></td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="izquierda">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="izquierda">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="izquierda">Fee</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha" style="text-align:right">$&nbsp;<%=Fee%></td>    
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td colspan="3" class="izquierda">Medical Expenses</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha" style="text-align:right">$&nbsp;<%=Costo%></td> 
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="izquierda">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="izquierda">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="izquierda">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td colspan="3" class="izquierda">Bank an transfer charges</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha" style="text-align:right">$&nbsp;<%=Trans%></td> 
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="izquierda">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="izquierda">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
    <td>&nbsp;</td>
    <td class="derecha">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="abaizq">&nbsp;</td>
    <td class="abajo">&nbsp;</td>
    <td class="abajo">&nbsp;</td>
    <td class="abajo">&nbsp;</td>
    <td class="abader">&nbsp;</td>
    <td class="abajo">&nbsp;</td>
    <td class="abader">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>Total</td>
    <td class="derecha">&nbsp;</td>
    <td class="abajo">&nbsp;</td>
    <td class="abader" style="text-align:right">$&nbsp;<%=cdbl(Fee) + cdbl(costo) + cdbl(trans)%></td> 
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="8" style="text-align:center">10880 Nw. 8th. St. Pembroke Pines FL 330265 USA
Phone/Fax: (954) 237-1905</td>
  </tr>
</table>

 </center>
</body>
</html>
