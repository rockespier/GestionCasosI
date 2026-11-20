<%@LCID = 1034%>
<!--#include file="main.asp"-->
<%
Response.Expires = -1
Response.AddHeader "Cache-Control", "no-cache, no-store, must-revalidate"
Response.AddHeader "Pragma", "no-cache"

Set rs = Server.CreateObject("ADODB.RECORDSET")

idUsuarioCreador = request("idUsuarioCreador")
idCaso = request("idCaso")
idIncidencia = request("idIncidencia")

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

' RT-119: eliminar incidencia (borrado definitivo, solo administradores).
' El perfil se valida con el usuario de la sesion, no con el de la URL.
if Request.ServerVariables("REQUEST_METHOD") = "POST" and request("accion") = "eliminar" then
	usuarioSesion = trim(session("codUsr") & "")
	esAdmin = false
	if usuarioSesion <> "" then
		sql = " select tip_usuario from usuario where cod_usuario='"& replace(usuarioSesion,"'","''") &"' and est_activo='S'"
		cn.open strUrl
		rs.open sql, cn
		if not rs.eof then
			esAdmin = (trim(rs("tip_usuario")) = "A")
		end if
		rs.close()
		cn.close
	end if

	if not esAdmin then
		response.Write("No autorizado. Vuelva a ingresar al sistema con un usuario administrador.")
		response.End()
	end if

	if isNumeric(idIncidencia) and idIncidencia <> "" then
		sql = "delete from incidencia where idIncidencia=" & clng(idIncidencia) & " and Caso_idCaso=" & idCaso_bd
		cn.open strUrl
		cn.execute(sql)
		cn.close
	end if

	response.Redirect("registro_incidencia.asp?idCaso="&idCaso_bd&"&idUsuarioCreador="&idUsuarioCreador&"")
end if

if Request.ServerVariables("REQUEST_METHOD") = "POST" then

desIncidencia = trim(replace(request("desIncidencia"),"'","''"))
fhrIncidencia = request("fhrIncidencia")
horIncidencia = SiVacio(request("horIncidencia"),0) 
minIncidencia = SiVacio(request("minIncidencia"),0)

'	sql = "update incidencia set fhrIncidencia='"& FormateaFecha(fhrIncidencia,"yyyymmdd") &"',desIncidencia='"&desIncidencia&"',"
'	sql = sql & " horIncidencia="&horIncidencia&",minIncidencia="&minIncidencia&" where idIncidencia=" & idIncidencia

	sql = "update incidencia set desIncidencia='"&desIncidencia&"',"
	sql = sql & " horIncidencia="&horIncidencia&",minIncidencia="&minIncidencia&" where idIncidencia=" & idIncidencia

	cn.open strUrl
	cn.execute(sql)
	cn.close
	
	idCaso = "00000" & idCaso
	idCaso = Right(idCaso,5)

'		sql = "insert into suceso(dessuceso,codusuario,idCaso) values('Incidencia modificada Caso NCA-"& idCaso &"','"&idUsuarioCreador&"','NCA-"&idCaso&"')"
'		'response.Write(sql)
'		'response.End()
'		cn.open strUrl
'		cn.execute(sql)
'		cn.close()
	
	desIncidencia = ""
	fhrIncidencia = ""
	horIncidencia = "" 
	minIncidencia = ""	
	
	response.Redirect("registro_incidencia.asp?idCaso="&idCaso_bd&"&idUsuarioCreador="&idUsuarioCreador&"")
	
end if

	sql = "select idIncidencia,fhrIncidencia,cast(desIncidencia as nvarchar(4000)),horIncidencia,minIncidencia,fchCreacion,codUsuario"	
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
	
	if idIncidencia<>"" then
		sql = "select cast(desIncidencia as nvarchar(4000)) ,horIncidencia,minIncidencia,fhrIncidencia,fchCreacion,codUsuario"	
		sql = sql & " from incidencia"
		sql = sql & " where idIncidencia=" & idIncidencia
	
		cn.open strUrl
		rs.open sql, cn
		if not rs.eof then
			desIncidencia = rs.fields(0)
			horIncidencia = rs.fields(1)
			minIncidencia = rs.fields(2)
			fhrIncidencia = rs.fields(3)
			fchCreacion = rs.fields(4)
			
		end if
		rs.close()
		cn.close
	end if

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
  </script> 
</head>
<body bgcolor="#e9e9e9" style="font-size:11px;" leftmargin="0">
<form name="frm" method="post" action="actualizar_incidencia.asp">    
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
               <!--<td>&nbsp;<%=FormateaFecha(arrIncidencia(1,i),"ddmmyyyy")%>&nbsp; <%=horas%>:<%=minutos%></td>-->
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
    
    <td colspan="2" valign="top"><textarea name="desincidencia" cols="40" rows="6" style="width:100%" ><%=TextoTextarea(desIncidencia)%></textarea></td>
  </tr>
  <tr>
   
    <td align="left" nowrap="nowrap">
    <!--
    Fecha:
    <input type="text" id="fhrIncidencia" name="fhrIncidencia" size="8" value="<%=fhrIncidencia%>" /><select name="horIncidencia">
    <%for i=0 to 23%>
    <option value="<%=i%>" <%if i=horIncidencia then%> selected="selected" <%end if%> ><%=i%></option>
    <%next%>
    </select> : <select name="minIncidencia">
    <%for i=0 to 59%>
    <option value="<%=i%>" <%if i=minIncidencia then%> selected="selected" <%end if%>><%=i%></option>
    <%next%>
    </select>
    -->
    </td>
    <td align="right">
    <%if trim(tip_usuario)="A" and idIncidencia<>"" then%>
    <input type="button" name="btnEliminar" value="Eliminar Incidencia" onclick="eliminarIncidencia();" />&nbsp;
    <%end if%>
    <input type="submit" name="btnGrabar" value="Grabar Incidencia" />&nbsp;&nbsp;</td>
  </tr>  
</table>

        </fieldset>    
    </td>
  </tr>
</table>
   <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
   <input type="hidden" name="idCaso" value="<%=idCaso%>" />
   
      <input type="hidden" name="idIncidencia" value="<%=idIncidencia%>" />
      <input type="hidden" name="accion" value="" />
  </form>
<script type="text/javascript">
function eliminarIncidencia() {
	if (confirm('\u00bfEliminar definitivamente la incidencia <%=idIncidencia%>? Esta acci\u00f3n no se puede deshacer.')) {
		document.frm.accion.value = 'eliminar';
		document.frm.submit();
	}
}
</script>
</body>
</html>
