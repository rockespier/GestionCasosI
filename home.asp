<!--#include file="main.asp"-->

<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

Mensaje=""
cod_usuario = Request("cod_usuario")
fchini   = Request.Form("fchini")
fchfin   = Request.Form("fchfin")
tip_usuario = Request.Form("tip_usuario")
 	
opciondefault = "blank.htm"
	
	strSql = "select o.id_opcion,o.des_opcion,o.img_opcion from accesos a, opciones o"
	strSql = strSql &" where a.id_opcion = o.id_opcion"
	strSql = strSql &" and o.tip_opcion='P' and a.id_usuario ='"& cod_usuario &"'"
	strSql = strSql &" order by num_orden"
	
	cn.open strUrl
		rs.open strSql, cn
		if not rs.eof then
		arrOpciones = rs.GetRows()
		end if
		rs.close()
	cn.close
	
'	sql = "insert into suceso(dessuceso,codusuario) values('Login al sistema','"&cod_usuario&"')"
'	cn.open strUrl
'	cn.execute(sql)
'	cn.close()

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Med Global Assist</title>
<LINK href="admin_login.css" type=text/css rel=stylesheet>
<LINK href="estilo.css" type=text/css rel=stylesheet>

</head>

<body>
<center>
<table width="1048" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left"><span class="style1">Sistema de</span> <span class="style3">Gesti&oacute;n de Casos</span></td>
  </tr>
  <tr>
    <td bgcolor="#ff4935" height="4px"></td>
  </tr>
  <tr>
    <td background="../images/mbbg.png" bgcolor="#DDE5D6"><div align="right">
      <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#fff">
        <tr>
            <td background="../images/mbbg.png" align="right" style="padding-right:5px;">
                <a href="index.asp"><span class="titulo_menu">Usuario:</span></a> <a href="index.asp">
                <span class="titulo_menu"><%=cod_usuario%></span></a>
            </td>
             <td width="90">
                 <div align="right"><a href="index.asp">
                 <img src="logo.png" /></a>
                 </div>
             </td>
        </tr>
      </table>
    
    </td>
  </tr>
  <tr>
    <td style="border-bottom: 1px #ff4935 solid;" align="right" bgcolor="#ff4935">
 
      <table border="0" cellspacing="0" cellpadding="0" style="border-color:#ff4935; border-style:solid; border-width:medium; height:23px">       
        <tr>
        <%if isArray(arrOpciones) then
			for i=0 to ubound(arrOpciones,2)
				if i=0 then
				opciondefault=arrOpciones(0,i)
				end if
		%>
          <td class="style4" width="140" bgcolor="#ff4935" style="border:#000 1px solid;">
          <div align="center">
          <a href="<%=trim(arrOpciones(0,i))%>?idUsuarioCreador=<%=cod_usuario%>&tip_usuario=<%=tip_usuario%>" target="detalle" style="color:#FFFFFF">
		  <%=trim(arrOpciones(1,i))%></a>
          </div>
          </td>
   		<%next
		end if%>        

        </tr>
      </table>
   
    </td>
  </tr>
	<tr>
    	<td>
        <iframe name="detalle" id="detalle" src="<%=trim(opciondefault)%>?idUsuarioCreador=<%=cod_usuario%>&tip_usuario=<%=tip_usuario%>" 
             scrolling="no" frameborder="0" width="100%" height="1200px"></iframe>
        </td>
    </tr>
</table>
</center>


</body>

</html>

