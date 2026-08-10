<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

idUsuarioCreador = request("idUsuarioCreador")
tip_usuario = request("tip_usuario")
est_activo = SiVacio(request("est_activo"),"N")


sql = "select cod_usuario ,des_usuario,des_email,des_cargo,des_pais,est_activo"
sql = sql & " from usuario "

if est_activo="N" then
	sql = sql & " where est_activo='S'"
end if

sql = sql & " order by des_usuario"

	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
	arrUsuario = rs.GetRows()
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
<script>
	function Enviar(){
		frm1.submit();
	}
</script>
</head>
<body bgcolor="#e9e9e9">
<center>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="titulo">
    <br />
    Consulta de usuarios
    <br />
	<br />
</td>
    <td>&nbsp;</td>
  </tr>
    <tr>
    <td>&nbsp;</td>
    <td align="center"><br />
    	<form name="frm" method="post" action="registro_usuario.asp">
         <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
          <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
            <input type="submit" name="btnEnviar" value="Nuevo Usuario" />            
        </form>
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center">
    <fieldset>
            <legend class="titulo_grupo" >Usuarios:</legend> 
			<form id="frm1" name="frm1" method="post" action="lista_usuarios.asp">
			Mostrar Inactivos: <input type="checkbox" value="S" name="est_activo" <%if est_activo="S" then%> checked <%end if%> onclick="javascript:Enviar();" />
			
            <div style="width:100%; height:260px; overflow:scroll-y">
            <table border="0" align="center" cellpadding="1" cellspacing="1" class="contenido_tabla1" width="98%">
              <tr bgcolor="#FF4935">
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid;border-left:#000 1px solid;">Cod.Usuario</td>
                <td align="center" class="titulo_tabla" width="20%" style="border-right:#000 1px solid">Descripción</td>
                <td align="center" class="titulo_tabla" width="20%" style="border-right:#000 1px solid">Correo</td>
                <td align="center" class="titulo_tabla" width="20%" style="border-right:#000 1px solid">Cargo</td>               
                <td align="center" class="titulo_tabla" width="20%" style="border-right:#000 1px solid">País</td>                                          
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">Activo?</td>
              </tr>
           <%if isArray(arrUsuario) then
		   colorfila = 0
		   		for i=0 to ubound(arrUsuario,2)
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
                <a href="actualizar_usuario.asp?cod_usuario=<%=arrUsuario(0,i)%>&idUsuarioCreador=<%=idUsuarioCreador%>&tip_usuario=<%=tip_usuario%>"><%=arrUsuario(0,i)%>               	 </a>
                </td>
                <td align="center">&nbsp;<%=ucase(arrUsuario(1,i))%></td>    
                <td align="center">&nbsp;<%=ucase(arrUsuario(2,i))%></td>    
                <td align="center">&nbsp;<%=ucase(arrUsuario(3,i))%></td>    
                <td align="center">&nbsp;<%=ucase(arrUsuario(4,i))%></td>               
                <td align="center">&nbsp;<%if arrUsuario(5,i)="S" then%>SI<%else%>NO<%end if%></td>
              </tr>
              	<%
			 	next
			 end if	
			 %> 
            </table>
</div>
</form>
            <br />
           
     </fieldset>       
    	
    </td>
  </tr>
 
 </table>
</center>
</body>
</html>
