<%@LCID = 1034%>
<!--#include file="main.asp"-->
<%
' RT-112: que el navegador no muestre una version guardada de la lista
Response.Expires = -1
Response.AddHeader "Cache-Control", "no-cache, no-store, must-revalidate"
Response.AddHeader "Pragma", "no-cache"

Set rs = Server.CreateObject("ADODB.RECORDSET")

idUsuarioCreador = request("idUsuarioCreador")
idCaso = request("idCaso")

if instr(idCaso,"-")>0 then
	idCaso_bd = clng(mid(idCaso,5))
else
	idCaso_bd = clng(idCaso)
end if

	sql = " select tip_usuario from usuario where cod_usuario='"& idUsuarioCreador &"' and est_activo='S'"
	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
		tip_usuario = rs("tip_usuario")
	end if	
	cn.close

if Request.ServerVariables("REQUEST_METHOD") = "POST" then

' Se escapan las comillas simples en el servidor (antes se borraban en el navegador)
desIncidencia = trim(replace(request("desIncidencia"),"'","''"))
fhrIncidencia = request("fhrIncidencia")
horIncidencia = SiVacio(request("horIncidencia") ,0)
minIncidencia = SiVacio(request("minIncidencia") ,0)

if desIncidencia <> "" then

'sql = "insert into incidencia(Caso_idCaso,fhrIncidencia,"
'sql = sql & "desIncidencia,codUsuario,fchCreacion,horIncidencia,minIncidencia)"
'sql = sql & "values ("&idCaso_bd&",'"&FormateaFecha(fhrIncidencia,"yyyymmdd")&"',"
'sql = sql & "'"&desIncidencia&"','"&idUsuarioCreador&"',getdate(),"&horIncidencia&","&minIncidencia&")"

sql = "insert into incidencia(Caso_idCaso,"
sql = sql & "desIncidencia,codUsuario,horIncidencia,minIncidencia,fchCreacion)"
sql = sql & "values ("&idCaso_bd&","
sql = sql & "'"&desIncidencia&"','"&idUsuarioCreador&"',"&horIncidencia&","&minIncidencia&",DATEADD(hour,-5,GETUTCDATE()))"

cn.open strUrl
cn.execute(sql)
cn.close
end if

'		sql = "insert into suceso(dessuceso,codusuario,idCaso) values('Incidencia registrada Caso NCA-"& idCaso &"','"&idUsuarioCreador&"','NCA-"&idCaso&"')"
'		cn.open strUrl
'		cn.execute(sql)
'		cn.close()

	' RT-112: redirigir (POST/Redirect/GET) para que la lista se recargue limpia
	' y un F5 no vuelva a enviar el formulario.
	response.Redirect("registro_incidencia.asp?idCaso="&idCaso_bd&"&idUsuarioCreador="&Server.URLEncode(idUsuarioCreador))

end if

	sql = "select idIncidencia,fhrIncidencia,cast(desIncidencia as nvarchar(4000)) ,horIncidencia,minIncidencia,fchCreacion,codUsuario"
	sql = sql & " from incidencia"
	sql = sql & " where Caso_idCaso=" & idCaso_bd
	sql = sql & " order by fchCreacion desc"

	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
	arrIncidencia = rs.GetRows()
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
<link type="text/css"rel="stylesheet"
     href="http://jquery-ui.googlecode.com/svn/tags/1.7/themes/redmond/jquery-ui.css" />   
  <script type="text/javascript"
     src="https://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.min.js"></script>
  <script type="text/javascript"
     src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.7.2/jquery-ui.min.js"></script>
  <script type="text/javascript">
      $(document).ready(function() {
          $('#fhrIncidencia').datepicker(
            {   dateFormat: 'dd/mm/yy',
                minDate: '+0D',
				//minDate: '-100Y',
                maxDate: '+1Y',
				//maxDate: '+0D',
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

function limpiar(input)
{
s = input.value;
filteredValues = "'";     // Characters stripped out
var i;
var returnString = "";
for (i = 0; i < s.length; i++) {  // Search through string and append to unfiltered values to returnString.
var c = s.charAt(i);
if (filteredValues.indexOf(c) == -1) returnString += c;
}
input.value = returnString;
}

function replacecarriagereturn(textarea,replaceWith)
{ 
textarea.value = escape(textarea.value); //encode all characters in text area to find carriage return character

	for(i=0;i<=15;i++)
	{ 
	//loop through string, replacing carriage return encoding with HTML break tag
		if(textarea.value.indexOf("%0D%0A") > -1)
		{ 
		//Windows encodes returns as \r\n hex
		textarea.value=textarea.value.replace("%0D%0A",replaceWith)
		}
		else if(textarea.value.indexOf("%0A") > -1)
		{ 
		//Unix encodes returns as \n hex
		textarea.value=textarea.value.replace("%0A",replaceWith)
		}
		else if(textarea.value.indexOf("%0D") > -1)
		{ 
		//Maclngosh encodes returns as \r hex
		textarea.value=textarea.value.replace("%0D",replaceWith)
		}
	}
textarea.value=unescape(textarea.value); //decode all characters in text area back
}
  </script> 
</head>
<body bgcolor="#e9e9e9" style="font-size:11px;" leftmargin="0">
<form name="frm" method="post" action="registro_incidencia.asp">    
 <table width="99%" border="0" cellspacing="1" cellpadding="1" class="contenido_tabla1">
  <tr>
  	<td>
    	<fieldset style="width:99%">
          <legend class="titulo_grupo" >Incidencias:</legend>   
          <table width="98%" border="0" cellspacing="1" cellpadding="1" class="contenido_tabla1">
  <tr>
    <td width="66%" valign="top" rowspan="2">
     <div style="width:100%; height:280px; overflow-y:auto;">
    	<table width="600" border="0" align="center" cellpadding="1" cellspacing="0" class="tabla_incidencias">
              <tr bgcolor="#FF4935">
                <td align="center" class="titulo_tabla" width="60" style="border-right:#000 1px solid;border-left:#000 1px solid;">&nbsp;Id</td>      
                <td align="center" class="titulo_tabla" width="70" style="border-right:#000 1px solid">&nbsp;Fecha</td>
                <td align="center" class="titulo_tabla" width="70" style="border-right:#000 1px solid">&nbsp;Usuario</td>       
                <td align="center" class="titulo_tabla" width="470" style="border-right:#000 1px solid">&nbsp;Detalle de Incidencia</td>
              </tr>
           <%if isArray(arrIncidencia) then
		   		for i=0 to ubound(arrIncidencia,2)
					if i mod 2 = 0 then
						claseFila = "fila_par"
					else
						claseFila = "fila_impar"
					end if
					horas = right("00" & arrIncidencia(3,i),2)
					minutos = right("00" & arrIncidencia(4,i),2)

			%>
              <tr class="contenido_tabla1 <%=claseFila%>">
                <td align="center">
        <%if trim(tip_usuario)="A" then%>
		<a href="actualizar_incidencia.asp?idIncidencia=<%=arrIncidencia(0,i)%>&idCaso=<%=idCaso%>&idUsuarioCreador=<%=idUsuarioCreador%>">
            <%=arrIncidencia(0,i)%></a>
		<%else%>
			<%=arrIncidencia(0,i)%>
		<%end if%>
                </td>
                <!-- <td>&nbsp;<%=FormateaFecha(arrIncidencia(1,i),"ddmmyyyy")%>&nbsp; <%=horas%>:<%=minutos%></td> -->
                <td><%=arrIncidencia(5,i)%></td>
                <td><%=arrIncidencia(6,i)%></td>
                <td class="detalle_incidencia"><%=TextoHTML(arrIncidencia(2,i))%></td>
              </tr>
             <%
			 	next
			 end if	
			 %> 
            </table>
            </div>
    </td>
   
    <td colspan="2" valign="top"><textarea name="desincidencia" cols="40" rows="6" style="width:100%"></textarea></td>
  </tr>
  <tr>
    
    <td align="left" nowrap="nowrap">
    <!--
    Fecha:
    <input type="text" id="fhrIncidencia" name="fhrIncidencia" size="6" /><select name="horIncidencia">
    <%for i=0 to 23%>
    <option value="<%=i%>"><%=i%></option>
    <%next%>
    </select> : <select name="minIncidencia">
    <%for i=0 to 59%>
    <option value="<%=i%>"><%=i%></option>
    <%next%>
    </select>
    -->
    </td>
    <td align="right">
    <input type="submit" name="btnGrabar" value="Grabar Incidencia" />&nbsp;&nbsp;</td>
  </tr>  
</table>

        </fieldset>    
    </td>
  </tr>
</table>
   <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
   <input type="hidden" name="idCaso" value="<%=idCaso%>" />
   <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
   
  </form>
</body>
</html>
