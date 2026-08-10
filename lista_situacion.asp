<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

idUsuarioCreador = request("idUsuarioCreador")
tip_usuario = request("tip_usuario")
idStatus = request("idStatus")
desStatus = request("desStatus")
desIcono = request("desIcono")
numOrden = request("numOrden")
estRegistro = SiVacio(request("estRegistro"),"N")

if Request.ServerVariables("REQUEST_METHOD") = "POST" then

	if idStatus<>"" then
		'Editar
		sql = "update Status set desStatus='"& desStatus &"',"
		sql = sql & " desIcono='"&desIcono&"',numOrden="& numOrden & ",estRegistro='"& estRegistro &"'"
		sql = sql & " where idStatus="& idStatus 
		cn.open strUrl
		cn.execute(sql)
		cn.close	
	
	else
		'Nuevo
		sql = "insert into Status(desStatus,desIcono,numOrden,"
		sql = sql & "estRegistro,codUsuario,fchCreacion) "
		sql = sql & " values('"&desStatus&"','"&desIcono&"','"&numOrden&"','"&estRegistro&"','"&idUsuarioCreador&"',now())"
		cn.open strUrl
		cn.execute(sql)
		cn.close	
	
	end if
	
	
	'Limpiar variables
	idStatus 	= ""
	desStatus 	= ""
	desIcono 	= ""
	numOrden 	= ""
	estRegistro = ""
	
else

	if idStatus<>"" then
		sql = "select desStatus,desIcono,numOrden,estRegistro"
		sql = sql & " from Status  where idStatus="& idStatus

		cn.open strUrl
		rs.open sql, cn
		if not rs.eof then
			desStatus = rs.fields(0)
			desIcono = rs.fields(1)	
			numOrden = rs.fields(2)
			estRegistro = rs.fields(3)
		end if
		rs.close()
		cn.close
		
	end if
	
end if

sql = "select idStatus ,desStatus, desIcono, numOrden, estRegistro  "
sql = sql & " from Status "
sql = sql & " order by numOrden"

	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
	arrStatus = rs.GetRows()
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
</head>
<body bgcolor="#e9e9e9">
<center>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="titulo" align="left">
    <br />
    Registro de situaciones
    <br />
	<br />
</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" align="center">
    <fieldset>
            <legend class="titulo_grupo" >Situaciones:</legend> 
            <div style="width:100%; height:260px; overflow-y:auto;">
            <table border="0" align="center" cellpadding="1" cellspacing="1" class="contenido_tabla1" width="98%">
              <tr bgcolor="#DF7401">
                <td align="center" class="titulo_tabla" width="15%" style="border-right:#000 1px solid;border-left:#000 1px solid;">Cod.Situacion</td>
                <td align="center" class="titulo_tabla" width="30%" style="border-right:#000 1px solid">Descripci&oacute;n</td>
                <td align="center" class="titulo_tabla" width="25%" style="border-right:#000 1px solid">Icono</td>
                <td align="center" class="titulo_tabla" width="15%" style="border-right:#000 1px solid">Orden</td>
                <td align="center" class="titulo_tabla" width="15%" style="border-right:#000 1px solid">Activo?</td>
              </tr>
           <%if isArray(arrStatus) then
		   colorfila = 0
		   		for i=0 to ubound(arrStatus,2)
			if colorfila = 0 then
					color= "#FFFFFF"
			'colorfila=1
'			else
'					color="#COCOCO"
'					colorfila=0
			end if
			%>   
              <tr class="contenido_tabla1"  bgcolor="<%=color%>">
                <td align="center">
                <a href="lista_situacion.asp?idStatus=<%=arrStatus(0,i)%>&idUsuarioCreador=<%=idUsuarioCreador%>&tip_usuario=<%=tip_usuario%>"><%=arrStatus(0,i)%>               	 </a>
                </td>
                <td align="center">&nbsp;<%=ucase(arrStatus(1,i))%></td>
                <td align="center">&nbsp;<img src="images/banderas/<%=arrStatus(2,i)%>" border="0" /> </td>
                <td align="center">&nbsp;<%=arrStatus(3,i)%></td>
                <td align="center">&nbsp;<%if arrStatus(4,i)="S" then%>SI<%else%>NO<%end if%></td>
              </tr>
              	<%
			 	next
			 end if	
			 %> 
            </table>
</div>
            <br />
           
     </fieldset>       
    	
    </td>
  </tr>
  <tr>
    
    <td align="center" colspan="2">
    <fieldset style="width:97%">
            <legend class="titulo_grupo" >Crear / Actualizar Situacion:</legend> 
        <form name="frm" method="post" action="lista_situacion.asp"  >
            <table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="contenido_tabla1">
    		<tr>
            	<td>Descripcion:</td>
                <td colspan="2" align="left"><input name="desStatus" type="text" value="<%=desStatus%>" size="40" maxlength="30" /></td>
                <td>Activo?</td>
                <td align="left"><input type="checkbox" name="estRegistro" value="S" <%if estRegistro="S" then%> checked="checked" <%end if%> /> </td>
            </tr>
            <tr>
                <td>Icono:</td>
                <td align="left"><input name="desIcono" type="text" value="<%=desIcono%>" size="20" maxlength="30" />
                </td>
                <td>Orden:</td>
                <td>
                	<input type="text" name="numOrden" size="5" value="<%=numOrden%>" />
                </td>                
                <td><input type="submit" name="btnEnviar" value="Grabar situaci&oacute;n" />    </td>
            </tr>        
          </table>
        
        
         <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
         <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
         <input type="hidden" name="idStatus" value="<%=idStatus%>" />
                    
        </form>
        </fieldset>
    </td>
  </tr>
</table>
</center>
</body>
</html>
