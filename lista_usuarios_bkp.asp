<%@LCID = 1034%>
<!--#include virtual="/main.asp"-->
<%
Set rn = Server.CreateObject("ADODB.RECORDSET")

Mensaje=""
cod_usuario = request("cod_usuario")
des_usuario   = Request("des_usuario")

CantidadMostrada = request("paginarpor")
if CantidadMostrada = "" then
CantidadMostrada = 20
end if

pagina = request("pagina")
if pagina = "" then
pagina = 1
else 
pagina = cInt(request("pagina"))
end if


sql=""
	sql = " select cod_usuario,est_mostrar,des_usuario,des_email,des_cargo,des_pais"
	sql = sql & " from usuario "
	
	if des_usuario<>"" then
	sql = sql & " where des_usuario like '%"& des_usuario &"%'"
	end if	
	sql = sql & " order by cod_usuario "

sqladd = "limit "&(pagina-1) * CantidadMostrada &", " & CantidadMostrada &" "
'response.Write(sql)
'response.End()
sql1 = sql & sqladd
	
	cn.open strUrl
	rn.open sql1, cn
	
		If rn.EOF = false AND rn.BOF = false then

			Set RS2 = Server.CreateObject("ADODB.Recordset")
			rs2.open sql, cn
			total = 0
			while not rs2.eof
			total = total + 1
			rs2.movenext
			wend
			'paginas = int( Total / CantidadMostrada ) + 1
			if (Total mod CantidadMostrada) = 0 then Total = int(Total / CantidadMostrada) else cant_paginas = int(Total / CantidadMostrada) + 1
		End if


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>::: Network Care Assist ::: </title>
<LINK href="estilo.css" type=text/css rel=stylesheet>
</head>

<body>
<center>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50"><img src="images/47.jpg" width="50" height="50" /></td>
        <td class="style3" width="150">Usuarios</td>
        <td align="right">
        <form name="frm" id="frm" action="lista_usuarios.asp" method="post">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="24%" class="style3" >[<a href="registro_usuario.asp?cod_usuario=<%=cod_usuario%>">Nuevo Registro</a>]</td>
            <td width="5%">&nbsp;</td>
            <td width="23%" class="style3" >[Filtros]</td>
            <td colspan="3" class="style4">Apellidos y Nombres:&nbsp;&nbsp;<input type="text" name="des_usuario" size="25" value="<%=des_usuario%>" /></td>
            <td width="11%"><input type="submit" name="submit" class="button" value="Buscar" />&nbsp;</td>
          </tr>
         
        </table>
        <input name="cod_usuario" type="hidden" value="<%=cod_usuario%>" />
        </form></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" align="center" cellpadding="1" cellspacing="2">
      <tr bgcolor="#DDE5D6">
        <td width="11%" class="style7">Actualizar</td>
        <td width="8%" class="style7">Estado</td>
        <td width="41%" class="style7">Apellidos y Nombres</td>
        <td width="15%" class="style7">Email</td>
        <td width="15%" class="style7">Cargo</td>
        <td width="10%" class="style7">Usuario</td>        
        </tr>
<%
while not rn.eof
	  %>
      <tr>
        <td align="center">
        <a href="actualizar_usuario.asp?id_usuario=<%=rn.Fields(0)%>&amp;cod_usuario=<%=cod_usuario%>">
        <img src="images/38.jpg" width="25" height="25" border="0" /></a></td>        
        <td align="center">
        <%if rn.Fields(1)="S" then%>
        <img src="images/ver-ok.jpg" width="25" height="25" />
        <%else%>
        <img src="images/ver-ko.jpg" width="25" height="25" />
        <%end if%>
        </td>
        <td class="style6">&nbsp;<%=Ucase(rn.Fields(2))%>&nbsp;</td>
        <td align="center" class="style6"><%=rn.Fields(3)%></td>
        <td align="center" class="style6"><%=rn.Fields(4)%></td>
        <td align="left" class="style6"><%=rn.Fields(0)%></td>
       
      </tr>
<%
		rn.movenext
wend

' Cerramos y limpiamos...
rn.Close

Set rn = Nothing
cn.close()
Set cn = Nothing
%>
    <tr><td colspan="6">&nbsp;</td></tr>
    <tr>
    <td colspan="6" align="center" background="../images/mbbg.png" bgcolor="#DDE5D6">
        <p align="left">
		<%For p = 1 To paginas%>
		<a href="lista_noticia.asp?pagina=<%=p%>"><%=p%></a> | 
		<%Next%>
		</p>   
    </td>
    </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</center>
</body>
</html>
