<%@LCID = 1034%>
<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")
titulo_grupo = request("tip_filtro_busqueda")
des_filtro	= request("des_filtro")
idUsuarioCreador = request("idUsuarioCreador")
tip_usuario = request("tip_usuario")
idStatus = request("idStatus")
idCaso = request("idCaso")

fchInicio  	= SiVacio(Request("fchInicio"),DateAdd("d",-0,date()))
fchFin   	= SiVacio(Request("fchFin"),DateAdd("d",1,date()))

sql = "select fchCreacion,codusuario,dessuceso"
sql = sql & " from suceso a "
sql = sql & " where codusuario <> '' and idsuceso >= 3144 "

if trim(idCaso) <>"" then
	sql = sql & " and a.idcaso ='"& idCaso &"'"
else
	if fchInicio <> "" then
	fchInicio1 = fchInicio
	fchFin1 = fchFin
		sql = sql & " and a.fchCreacion > '"& FormateaFecha(fchInicio1,"yyyymmdd") &"' and a.fchCreacion < '" & FormateaFecha(fchFin1,"yyyymmdd") & "'"
	end if
end if

	sql = sql & " order by fchCreacion desc"
	
	'response.Write(sql)
	
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
<title>::: Network Care Assist ::: </title>
<LINK href="estilo.css" type=text/css rel=stylesheet>
 <link type="text/css"rel="stylesheet"
     href="http://jquery-ui.googlecode.com/svn/tags/1.7/themes/redmond/jquery-ui.css" />   
  <script type="text/javascript"
     src="http://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.min.js"></script>
  <script type="text/javascript"
     src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.7.2/jquery-ui.min.js"></script>
  <script type="text/javascript">
      $(document).ready(function() {
          $('#fchInicio').datepicker(
            {   dateFormat: 'dd/mm/yy',
                
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
		 $('#fchFin').datepicker(
            {   dateFormat: 'dd/mm/yy',
                minDate: '+0D',
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
<body bgcolor="#e9e9e9" style="font-size:11px;">
<center>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="titulo" align="left">
    <br />
    Consulta de sucesos
    <br />
	<br />
</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="70%"  align="left">
    <form name="frm" action="lista_suceso.asp" method="post">
    	<fieldset style="width:500px">
            <legend class="titulo_grupo" >Busqueda:</legend>            
            <table width="500" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="80" class="titulo_grupo">Caso:</td>               
                <td><input type="text" name="idCaso" id="idCaso" value="<%=idCaso%>" size="10" /></td>            
                <td>&nbsp;</td>
                <td width="100" class="titulo_grupo">Fechas:</td>
                <td><input type="text" name="fchInicio" id="fchInicio" value="<%=FormateaFecha(fchInicio,"ddmmyyyy")%>" size="10" readonly /></td>
                <td><input type="text" name="fchFin" id="fchFin" value="<%=FormateaFecha(fchFin,"ddmmyyyy")%>" size="10" readonly /></td>   
                <td width="50"><input type="submit" value="Buscar" /></td>
              </tr>
            </table>          
        </fieldset>
        <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
        <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
     </form>    
    </td>
    <td align="left" style="padding-right:5px">&nbsp;
      
    </td>
  </tr>
  <tr>
    <td colspan="2">
    <fieldset>
            <legend class="titulo_grupo" >Sucesos:</legend> 
            <div style="width:100%; height:260px; overflow-y:auto;">
            <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr bgcolor="#DF7401">
                <td align="center" class="titulo_tabla" width="25%" style="border-right:#000 1px solid;border-left:#000 1px solid;">&nbsp;Fecha</td>
                <td align="center" class="titulo_tabla" width="20%" style="border-right:#000 1px solid">&nbsp;Usuario</td>
                <td align="center" class="titulo_tabla" width="55%" style="border-right:#000 1px solid">&nbsp;Suceso</td>
               
              </tr>
           <%if isArray(arrCasos) then
		   
		   		colorfila = 0
		   		nro_registro = ubound(arrCasos,2) + 1
		   
		   		for i=0 to ubound(arrCasos,2)
			if colorfila = 0 then
					color= "#FFFFFF"
			'colorfila=1
'			else
'					color="#COCOCO"
'					colorfila=0
			end if
			%>   
              <tr class="contenido_tabla1" bgcolor="<%=color%>">
                <td align="center"><%=arrCasos(0,i)%></td>
                <td>&nbsp;<%=arrCasos(1,i)%></td>
                <td>&nbsp;<%=arrCasos(2,i)%></td>               
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
  <tr>
    <td class="contenido_tabla1">&nbsp;# Registros:<%=nro_registro%></td>
    <td align="center"><br />
    	
    </td>
  </tr>
</table>
</center>
</body>
</html>
