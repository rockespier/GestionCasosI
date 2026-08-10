<%@LCID = 1034%>
<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")
titulo_grupo = request("tip_filtro_busqueda")
des_filtro	= request("des_filtro")
idUsuarioCreador = request("idUsuarioCreador")
tip_usuario = request("tip_usuario")
idStatus = request("idStatus")
limit = SiVacio(request("txtLimit"),35)

fchInicio  	= SiVacio(Request("fchInicio"),DateAdd("d",-30,date()))
fchFin   	= SiVacio(Request("fchFin"),date())

'sql = "select c.desIcono,LPAD(a.idCaso,5,'0'),a.desNombre,a.idRefExterna,a.fchCreacion,a.idUsuarioCreador,c.desStatus,c.numorden"
sql = "select c.desIcono,a.idCaso,a.desNombre,a.idRefExterna,a.fchCreacion,a.idUsuarioCreador,c.desStatus,c.numorden"
sql = sql & " from Caso a LEFT JOIN empresa d ON a.idEmpresa = d.idEmpresa, StatusxCaso b, Status c "
sql = sql & " where a.idCaso = b.idCaso"
sql = sql & " and b.estRegistro='S'"
sql = sql & " and b.idStatus = c.idStatus" 
'"and a.idEmpresa=d.idEmpresa"


select case titulo_grupo
case "Empresa"
	sql = sql & " and ucase(d.desEmpresa) like '" & ucase(des_filtro) & "%'"
case "Referencia"
	sql = sql & " and a.idRefExterna='"& des_filtro & "'"
case "Caso"

if instr(des_filtro,"-")>0 then
	idCaso_bd = cint(mid(des_filtro,5))
else
	idCaso_bd = cint(des_filtro)
end if

	sql = sql & " and a.idCaso="& idCaso_bd
case "Servicio"	

    idServicio = left(trim(des_filtro),len(trim(des_filtro)) - 3)
	sql = sql & " and a.idCaso in (select Caso_idCaso from servicio where idServicio="& idServicio &")"
	
case "Cliente"
	sql = sql & " and ucase(a.desNombre) like '%" & ucase(des_filtro) & "%'"
case "Proveedor"
    sql = sql & " and a.idCaso in (select Caso_idCaso from servicio where desproveedor like '%"& ucase(des_filtro) &"%')"
end select

if idStatus <> ""  then
	sql = sql & " and b.idStatus=" & idStatus
end if

'if fchInicio <> "" then
'
'fchInicio1 = fchInicio
'fchFin1 = fchFin
'	sql = sql & " and a.fchCreacion between '"& FormateaFecha(fchInicio1,"yyyymmdd") &"' and '" & FormateaFecha(fchFin1,"yyyymmdd") & "'"
'end if

	sql = sql & " order by a.idCaso desc,a.fchCreacion desc limit "& limit
	
	
	
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
    Consulta de casos
    <br />
	<br />
</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="45%"  align="left">
    <form name="frm" action="lista_caso.asp" method="post">
    	<fieldset style="width:500px">
            <legend class="titulo_grupo" >Buscar por:</legend>            
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
                  <input name="tip_filtro_busqueda" type="radio" value="" <%if titulo_grupo="" then%> checked <%end if%> />Ninguno
                  <input name="tip_filtro_busqueda" type="radio" value="Empresa" <%if titulo_grupo="Empresa" then%> checked <%end if%>  />Empresa
                  <input name="tip_filtro_busqueda" type="radio" value="Referencia" <%if titulo_grupo="Referencia" then%> checked <%end if%>  />Referencia
                  <input name="tip_filtro_busqueda" type="radio" value="Caso" <%if titulo_grupo="Caso" then%> checked <%end if%>  />Caso
                  <input name="tip_filtro_busqueda" type="radio" value="Servicio" <%if titulo_grupo="Servicio" then%> checked <%end if%>  />Servicio
                  <input name="tip_filtro_busqueda" type="radio" value="Cliente" <%if titulo_grupo="Cliente" then%> checked <%end if%>  />Cliente
                  <input name="tip_filtro_busqueda" type="radio" value="Proveedor" <%if titulo_grupo="Proveedor" then%> checked <%end if%>  />Proveedor
                <!--
                <select name="tip_filtro_busqueda">
                				<option value="">[Seleccione]</option>
                                <option value="Empresa">Empresa</option>
                                <option value="Referencia">Referencia</option>
                                <option value="Caso">idCaso</option>
								<option value="Servicio">idServicio</option>
								<option value="Cliente">Cliente</option>
                                <option value="Proveedor">Proveedor</option>
                            </select> --></td>
                <td><!--<input type="text" name="fchInicio" id="fchInicio" value="<%=FormateaFecha(fchInicio,"ddmmyyyy")%>" size="10" readonly />--></td>
                <td><!--<input type="text" name="fchFin" id="fchFin" value="<%=FormateaFecha(fchFin,"ddmmyyyy")%>" size="10" readonly />-->
                <input type="text" name="des_filtro" size="15" value="<%=des_filtro%>" /></td>
                <td width="50"><input type="submit" value="Buscar" /></td>
              </tr>
			  <tr>
				<td>Mostrar M&aacute;ximo:<input type="text" size="3" name="txtLimit" value=<%=SiVacio(limit,35)%> /> </td>
			  </tr>
            </table>          
        </fieldset>
        <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
        <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
     </form>    
    </td>
    <td align="right" style="padding-right:5px">
    <%
	sql = "select idStatus,desStatus,desIcono"
				sql = sql & " from Status" 
				sql = sql & " where estRegistro='S'"
				sql = sql & " order by numorden"

				cn.open strUrl
				rs.open sql, cn
				if not rs.eof then
				arrStatus = rs.GetRows()
				end if
				rs.close()
				cn.close
	%>

					<table border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                      <tr>
						  <%if isArray(arrStatus) then
						  		for i=0 to ubound(arrStatus,2)	%>
                            	<td align="center">                                 
                <a href="lista_caso.asp?idStatus=<%=arrStatus(0,i)%>&idUsuarioCreador=<%=idUsuarioCreador%>&tip_usuario=<%=tip_usuario%>&txtLimit=500" >
                <img src="images/banderas/<%=arrStatus(2,i)%>" border="0" title="<%=arrStatus(1,i)%>" />   
                </a>                             
                                </td>
                          <%	next %>
                      	</tr>  
                        <tr>
                        	<%for j=0 to ubound(arrStatus,2)%>
                            	<td class="contenido_tabla2" width="75" align="center">
                        				<%=arrStatus(1,j)%>
                          </td>
                            <%next%>
                        </tr>  
						  <%end if%>  
                      
                    </table>

    </td>
  </tr>
  <tr>
    <td colspan="2">
    <fieldset>
            <legend class="titulo_grupo" >Casos:</legend> 
            <div style="width:100%; height:260px; overflow-y:auto;">
            <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr bgcolor="#DF7401">
                <td align="center" class="titulo_tabla" width="16%" style="border-right:#000 1px solid;border-left:#000 1px solid;">&nbsp;Estado</td>
                <td align="center" class="titulo_tabla" width="16%" style="border-right:#000 1px solid">&nbsp;Id Caso</td>
                <td align="center" class="titulo_tabla" width="26%" style="border-right:#000 1px solid">&nbsp;Nombres</td>
                <td align="center" class="titulo_tabla" width="16%" style="border-right:#000 1px solid">&nbsp;Ref. Externa</td>
                <td align="center" class="titulo_tabla" width="16%" style="border-right:#000 1px solid">&nbsp;Fecha</td>
                <td align="center" class="titulo_tabla" width="10%" style="border-right:#000 1px solid">&nbsp;Usuario</td>
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
                <td align="center">
                <a href="actualizar_caso.asp?idCaso=<%=arrCasos(1,i)%>&idUsuarioCreador=<%=idUsuarioCreador%>&tip_usuario=<%=tip_usuario%>">
                	<img src="images/banderas/<%=arrCasos(0,i)%>" border="0" title="<%=arrCasos(6,i)%>" />
                </a>
                </td>
                <td>&nbsp;NCA-<%=arrCasos(1,i)%></td>
                <td>&nbsp;<%=ucase(arrCasos(2,i))%></td>
                <td>&nbsp;<%=arrCasos(3,i)%></td>
                <td>&nbsp;<%=FormateaFecha(arrCasos(4,i),"ddmmyyyy")%></td>
                <td>&nbsp;<%=arrCasos(5,i)%></td>
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
    	<form name="frm" method="post" action="registro_caso.asp">
         <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
          <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
            <input type="submit" name="btnEnviar" value="Nuevo Caso" />            
        </form>
    </td>
  </tr>
</table>
</center>
</body>
</html>
