<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

idUsuarioCreador = request("idUsuarioCreador")
tip_usuario = request("tip_usuario")

idEmpresa 				= request("idEmpresa")
desEmpresa 				= request("desEmpresa")
estRegistro 			= SiVacio(request("estRegistro"),"N")
desDatosfiscales 		= request("desDatosfiscales")
empresacasodiasmax		= request("empresacasodiasmax")
empresacasoserviciomax	= request("empresacasoserviciomax")
empresarazonsocial		= request("empresarazonsocial")
empresadireccion		= request("empresadireccion")
empresacostosimple		= request("empresacostosimple")
empresacostomiddle		= request("empresacostomiddle")
empresacostocomplex		= request("empresacostocomplex")
empresatransferenciabanco	= request("empresatransferenciabanco")

if Request.ServerVariables("REQUEST_METHOD") = "POST" then

	if idEmpresa<>"" then
		'Editar
		sql = "update empresa set desEmpresa='"& desEmpresa &"',"
		sql = sql & " desDatosfiscales='"& desDatosfiscales &"',"
		sql = sql & " estRegistro='"& estRegistro &"',"
		sql = sql & " empresacasodiasmax="&empresacasodiasmax&","
		sql = sql & " empresacasoserviciomax="&empresacasoserviciomax&","
		sql = sql & " empresarazonsocial='"&empresarazonsocial&"',"
		sql = sql & " empresadireccion='"&empresadireccion&"',"
		sql = sql & " empresacostosimple="&empresacostosimple&","
		sql = sql & " empresacostomiddle="&empresacostomiddle&","
		sql = sql & " empresacostocomplex="&empresacostocomplex&","
		sql = sql & " empresatransferenciabanco="&empresatransferenciabanco&""
		sql = sql & " where idEmpresa="& idEmpresa 
		cn.open strUrl
		cn.execute(sql)
		cn.close	
	
	else
		'Nuevo
		sql = "insert into empresa(desEmpresa,"
		sql = sql & "estRegistro,desDatosfiscales,"
		sql = sql & " empresacasodiasmax,"
		sql = sql & " empresacasoserviciomax,"
		sql = sql & " empresarazonsocial,"
		sql = sql & " empresadireccion,"
		sql = sql & " empresacostosimple,"
		sql = sql & " empresacostomiddle,"
		sql = sql & " empresacostocomplex,"
		sql = sql & " empresatransferenciabanco"
		sql = sql & " values('"&desEmpresa&"','"&estRegistro&"','"&desDatosfiscales&"',"
		sql = sql & empresacasodiasmax &","
		sql = sql & empresacasoserviciomax &","
		sql = sql & "'"& empresarazonsocial &"',"
		sql = sql & "'"& empresadireccion &"',"
		sql = sql & empresacostosimple &","
		sql = sql & empresacostomiddle &","
		sql = sql & empresacostocomplex &","
		sql = sql & empresatransferenciabanco &")"
		cn.open strUrl
		cn.execute(sql)
		cn.close	
	
	end if
	
	'Limpiar variables
	idEmpresa 	= ""
	desEmpresa 	= ""
	estRegistro = ""
	desDatosfiscales = ""
	empresacasodiasmax=""
	empresacasoserviciomax=""
	empresarazonsocial=""
	empresadireccion=""
	empresacostosimple=""
	empresacostomiddle=""
	empresacostocomplex=""
	empresatransferenciabanco=""
	
else

	if idEmpresa<>"" then
		sql = "select desEmpresa,estRegistro,desDatosfiscales,"
		sql = sql & " empresacasodiasmax,"
		sql = sql & " empresacasoserviciomax,"
		sql = sql & " empresarazonsocial,"
		sql = sql & " empresadireccion,"
		sql = sql & " empresacostosimple,"
		sql = sql & " empresacostomiddle,"
		sql = sql & " empresacostocomplex,"
		sql = sql & " empresatransferenciabanco"
		sql = sql & " from empresa  where idEmpresa="& idEmpresa

		cn.open strUrl
		rs.open sql, cn
		if not rs.eof then
			desEmpresa = rs.fields(0)		
			estRegistro = rs.fields(1)
			desDatosfiscales = rs.fields(2)
			empresacasodiasmax	= rs.fields(3)
			empresacasoserviciomax = rs.fields(4)
			empresarazonsocial = rs.fields(5)
			empresadireccion = rs.fields(6)
			empresacostosimple = rs.fields(7)
			empresacostomiddle = rs.fields(8)
			empresacostocomplex = rs.fields(9)
			empresatransferenciabanco = rs.fields(10)
		end if
		rs.close()
		cn.close
		
	end if
	
end if

sql = "select idEmpresa ,desEmpresa, "
sql = sql & " empresacasodiasmax,"
sql = sql & " empresacasoserviciomax,"
sql = sql & " empresacostosimple,"
sql = sql & " empresacostomiddle,"
sql = sql & " empresacostocomplex,"
sql = sql & " estRegistro"
sql = sql & " from empresa "
sql = sql & " order by desEmpresa"

	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
	arrEmpresa = rs.GetRows()
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
</head>
<body bgcolor="#e9e9e9">
<center>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="titulo" align="left">
    <br />
    Registro de empresa
    <br />
	<br />
</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" align="center">
    <fieldset>
            <legend class="titulo_grupo" >Empresas:</legend> 
            <div style="width:100%; height:260px; overflow:scroll-y">
            <table border="0" align="center" cellpadding="1" cellspacing="1" class="contenido_tabla1" width="98%">
              <tr bgcolor="#FF4935">
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid;border-left:#000 1px solid;">Id.</td>
                <td align="center" class="titulo_tabla" width="20%" style="border-right:#000 1px solid">Descripción</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">Días Max</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">Servicio Max</td>                
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">Imp.Simple</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">Imp.Middle</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">Imp.Complex</td>          
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">Activo?</td>
              </tr>
           <%if isArray(arrEmpresa) then
		   colorfila = 0
		   		for i=0 to ubound(arrEmpresa,2)
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
                <a href="lista_empresa.asp?idEmpresa=<%=arrEmpresa(0,i)%>&idUsuarioCreador=<%=idUsuarioCreador%>&tip_usuario=<%=tip_usuario%>"><%=arrEmpresa(0,i)%>               	 </a>
                </td>
                <td align="center">&nbsp;<%=ucase(arrEmpresa(1,i))%></td>
                <td align="center">&nbsp;<%=arrEmpresa(2,i)%></td>
                <td align="center">&nbsp;<%=arrEmpresa(3,i)%></td>
                <td align="center">$&nbsp;<%=arrEmpresa(4,i)%></td>               
                <td align="center">$&nbsp;<%=arrEmpresa(5,i)%></td>               
                <td align="center">$&nbsp;<%=arrEmpresa(6,i)%></td>                                               
                <td align="center">&nbsp;<%if arrEmpresa(7,i)="S" then%>SI<%else%>NO<%end if%></td>
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
            <legend class="titulo_grupo" >Crear / Actualizar Empresa:</legend> 
        <form name="frm" method="post" action="lista_empresa.asp"  >
            <table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="contenido_tabla1">
            <tr>
            	<td>Nombre:</td>
                <td colspan="2">
                <input name="desEmpresa" type="text" value="<%=desEmpresa%>" size="60" maxlength="20" required /></td>                
                <td>
                Cantidad Máxima de Servicios                 
                <input name="empresacasoserviciomax" type="text" required
                value="<%=empresacasoserviciomax%>" size="10" maxlength="4" />
                </td>
            </tr>
            <tr>
            	<td>Razón Social:</td>
                <td colspan="2">
                <input name="empresaRazonsocial" type="text" 
                value="<%=empresarazonsocial%>" size="60" maxlength="30" /></td>
                <td>
                Días Máximos para Cierre                 
                <input name="empresacasodiasmax" type="text" required
                value="<%=empresacasodiasmax%>" size="10" maxlength="4" />
                </td>
            </tr>            
            <tr>
	            <td>Dirección:</td>
                <td>
                <input name="empresadireccion" type="text" 
                value="<%=empresadireccion%>" size="30" maxlength="40" /></td>
                <td>Identificacion Fiscal:</td>
                <td><input name="empresaidentificacionfiscal" type="text" 
                value="<%=empresaidentificacionfiscal%>" size="30" maxlength="15" /></td>
            </tr>
            <tr>
	            <td>Giro del Negocio:</td>
                <td><input name="empresagirodenegocio" type="text" value="<%=empresagirodenegocio%>" size="30" maxlength="30" /></td>
                <td>Costo Transferencia Banco:</td>
                <td>$ <input name="empresatransferenciabanco" type="text" value="<%=empresatransferenciabanco%>" size="20" maxlength="6" /></td>
            </tr>
            <tr>
	            <td colspan="4">
                <fieldset style="width:97%">
            	<legend class="titulo_grupo" >Costo por Tipo de Servicio:</legend> 
                <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                	<tr>
                    	<td>Simple</td>
                        <td>$ <input name="empresacostosimple" type="text" value="<%=empresacostosimple%>" size="20" maxlength="6" /></td>
                        <td>Middle</td>
                        <td>$ <input name="empresacostomiddle" type="text" value="<%=empresacostomiddle%>" size="20" maxlength="6" /></td>
                        <td>Complex</td>
                        <td>$ <input name="empresacostocomplex" type="text" value="<%=empresacostocomplex%>" size="20" maxlength="6" /></td>
                    </tr>
                </table>
            	</fieldset>    
                </td>
            </tr>
            <tr>
            	<td>Activo?</td>
                <td><input type="checkbox" name="estRegistro" value="S" 
				<%if estRegistro="S" then%> checked="checked" <%end if%> /></td>
                <td>&nbsp;</td>
                <td><input type="submit" name="btnEnviar" value="Grabar Empresa" /></td>
                                                
            </tr> 
          </table>        
        
         <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
         <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
         <input type="hidden" name="idEmpresa" value="<%=idEmpresa%>" />
                    
        </form>
        </fieldset>
    </td>
  </tr>
</table>
</center>
</body>
</html>
