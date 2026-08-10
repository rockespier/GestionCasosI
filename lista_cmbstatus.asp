<%@LCID = 1034%>
<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

idUsuarioCreador = request("idUsuarioCreador")
tip_usuario = request("tip_usuario")
idCaso  = request("idCaso")

	if idCaso<>"" then

		sql = "select a.fchCreacion,a.codUsuario,b.desstatus,b.desicono "
		sql = sql & " from StatusxCaso a, status b"
		sql = sql & " where idcaso=" & idCaso
		sql = sql & " and a.idstatus = b.idstatus"
		sql = sql & " order by  a.fchCreacion,b.numorden"

		cn.open strUrl
		rs.open sql, cn
		if not rs.eof then
		arrCambios = rs.GetRows()
		end if
		rs.close()
		cn.close
	
		idCaso = "0000" & idCaso
	
	end if


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>::: Network Care Assist ::: </title>
</head>
<body bgcolor="#ffffff" style="font-size:11px;">
<center>
<table width="98%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td class="titulo">
    <br />
    Registro de cambios Caso # NCA- <%=right(idCaso,5) %>
    <br />
	<br />
	</td>    
  </tr>
  <tr>
    <td align="center">   
            <div style="width:100%; height:260px; overflow:scroll-y">
            <table border="0" align="center" cellpadding="1" cellspacing="1" class="contenido_tabla1" width="98%">
              <tr bgcolor="#DF7401">
                <td align="center" class="titulo_tabla" width="30%" style="border-right:#000 1px solid;border-left:#000 1px solid;">Fecha</td>
                <td align="center" class="titulo_tabla" width="30%" style="border-right:#000 1px solid">Usuario</td>          
                <td align="center" class="titulo_tabla" width="40%" style="border-right:#000 1px solid">Status</td>				
              </tr>
           <%if isArray(arrCambios) then
		   colorfila = 0
		   		for i=0 to ubound(arrCambios,2)
			if colorfila = 0 then
					color= "#FFFFFF"
				colorfila=1
			else
					color="#COCOCO"
					colorfila=0
			end if
			%>   
              <tr class="contenido_tabla1"  bgcolor="<%=color%>">
                <td align="center"><%=arrCambios(0,i)%></td>
                <td align="center">&nbsp;<%=ucase(arrCambios(1,i))%></td>               
                <td align="center">&nbsp;<%=arrCambios(2,i)%></td>
              </tr>
              	<%
			 	next
			 end if	
			 %> 
            </table>
			</div>   	
    </td>
  </tr>  
</table>
</center>
</body>
</html>
