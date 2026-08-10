<%@LCID = 1034%>
<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

idUsuarioCreador = request("idUsuarioCreador")
idCaso = request("idCaso")
tip_usuario = request("tip_usuario")
idStatusAntes = request("idStatusAntes")
empresacasoserviciomax = request("empresacasoserviciomax")

if instr(idCaso,"-")>0 then
	idCaso_bd = clng(mid(idCaso,5))
else
	idCaso_bd = clng(idCaso)
end if

if Request.ServerVariables("REQUEST_METHOD") = "POST" and request("txtModo")="OK" then
	
	idEmpresa = SiVacio(request("idEmpresa"),0)
	anioCaso = year(date())
	mesCaso  = month(date())
		
	tipoCaso = SiVacio(request("tipoCaso"),request("tipoCasoAntes"))
	idRefExterna = request("idRefExterna")
	desNombre = request("desNombre")
	fchNacimiento = request("fchNacimiento")
	dia_nac = request("dia")
	mes_nac = request("mes")
	ano_nac = request("ano")	
	edadTitular = anioCaso - year(fchNacimiento)
	
	if mesCaso < mes_nac then
		edadTitular = edadTitular - 1
	end if	
	
	desTelefono = request("desTelefono")
	desDireccion = trim(replace(request("desDireccion"),chr(34),"''"))	
	idPais = request("idPais")
	idCiudad = 0
	des_ciudad = trim(replace(request("des_ciudad"),chr(34),"''"))	
	Nacimiento = FormateaFecha(fchNacimiento,"yyyymmdd")
	idStatus = request("idStatus")
	

	sql = "update caso set idEmpresa='"& idEmpresa &"',"
	if tipoCaso <> "" then
		sql = sql & " tipoCaso='"& tipoCaso & "',"
	end if
	sql = sql & " idRefExterna='"& idRefExterna &"',desNombre='"& desNombre &"',"
	sql = sql & " fchNacimiento='"& Nacimiento &"',edadTitular="&edadTitular&","
	sql = sql & " desTelefono='"& desTelefono &"',desDireccion='"& desDireccion &"',"
	sql = sql & " idPais='"& idPais &"',des_ciudad='"& des_ciudad &"',"
	sql = sql & " dia_nac='"&dia_nac&"', mes_nac='"&mes_nac&"', ano_nac='"& ano_nac &"'"
	sql = sql & " where idCaso="& idCaso_bd
	
	'response.Write(sql)
	'response.End()

	cn.open strUrl
	cn.execute(sql)
	cn.close	
	
		if trim(idStatus) <> trim(idStatusAntes) then
			sql = "update StatusxCaso set estRegistro='N' where idCaso=" & idCaso_bd
			cn.open strUrl
			cn.execute(sql)
			cn.close		
			
			sql = "insert into StatusxCaso values("&idStatus&","&idCaso_bd&",getdate(),'"&idUsuarioCreador&"','S');"
			cn.open strUrl
			cn.execute(sql)
			cn.close			
		end if		
			idCaso = "00000" & idCaso
	idCaso = Right(idCaso,5)
		'sql = "insert into suceso(dessuceso,codusuario,idCaso) values('Caso modificado:NCA-"& idCaso &"','"&idUsuarioCreador&"','NCA-"&idCaso&"')"
'		cn.open strUrl
'		cn.execute(sql)
'		cn.close()

	end if

	'sql = " select LPAD(x.idCaso,5,'0'),x.tipoCaso ,"
	sql = " select x.idCaso,x.tipoCaso ,"
	sql = sql & "  x.idEmpresa, x.idRefExterna , "
	sql = sql & "  x.fchCaso,  x.desNombre ,  x.fchNacimiento , x.edadTitular , x.desTelefono ,"
	sql = sql & "  x.desDireccion ,  x.idPais , x.idCiudad ,  x.des_ciudad ,"
	sql = sql & "  x.idUsuarioCreador , x.fchCreacion, "
	sql = sql & "  (select a.idStatus "
	sql = sql & "  from Status a, StatusxCaso b"
	sql = sql & "  where a.idStatus= b.idstatus"
	sql = sql & "  and b.idcaso=x.idcaso"
	sql = sql & "  and b.estregistro='S') idstatus, dia_nac, mes_nac, ano_nac,"
	sql = sql & "  empresacasodiasmax,empresacasoserviciomax"
	sql = sql & "  from caso x , empresa e"
	sql = sql & "  where x.idCaso=" & idCaso_bd
	sql = sql & "  and  x.idEmpresa = e.idEmpresa"
	
	'response.Write(sql)
	'response.End()
	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
		idCaso 	  	 = rs.fields(0)
		tipoCaso  	 = rs.fields(1)
		idEmpresa 	 = rs.fields(2)
		idRefExterna = rs.fields(3)
		fchCaso		= FormateaFecha(rs.fields(4),"ddmmyyyy")
		desNombre	= trim(rs.fields(5))
		fchNacimiento = FormateaFecha(rs.fields(6),"ddmmyyyy")
		dia_nac = rs.fields(16)
		mes_nac = trim(rs.fields(17))
		ano_nac = rs.fields(18)
		mesCaso  = month(date())		

	anioCaso = year(date())
	edad = anioCaso - year(fchNacimiento)
	if mes_nac <> "" then
		if mesCaso - mes_nac < 0 then
			edad = edad - 1
		end if	
	else
		edad = 0
	end if
		'edad  = rs.fields(7)
		desTelefono	= trim(rs.fields(8))
		desDireccion = trim(rs.fields(9))
		
		idPais = rs.fields(10)
		idCiudad = rs.fields(11)
		des_ciudad = trim(rs.fields(12))
		idStatus = rs.fields(15)
		idStatusAntes = idStatus
		tipoCasoAntes = tipoCaso
		empresacasodiasmax = rs.fields(19)
		empresacasoserviciomax = rs.fields(20)

		
	end if
	rs.close()
	cn.close	

	sql = ""
	sql = sql & " select idEmpresa,desEmpresa"
	sql = sql & " from Empresa"
	sql = sql & " where estregistro='S'"
	sql = sql & " order by 2"
	
	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
	arrEmpresa = rs.GetRows()
	end if
	rs.close()
	cn.close
	
	sql = ""
	sql = sql & " select codigo,nombre"
	sql = sql & " from pais"
	sql = sql & " order by nombre"
	
	cn.open strUrl
	rs.open sql, cn
	if not rs.eof then
	arrPais = rs.GetRows()
	end if
	rs.close()
	cn.close
	
	sql = ""
	sql = sql & " select id_opcion from accesos"
	sql = sql & " where id_opcion='tipo_caso'"
	sql = sql & " and id_usuario='"& idUsuarioCreador &"'"
	cn.open strUrl
	rs.open sql, cn

	if not rs.eof then
		arrAcceso = rs.GetRows()
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
          $('#fchNacimiento').datepicker(
            {   dateFormat: 'dd/mm/yy',
                //minDate: '+0D',
				minDate: '-100Y',
                //maxDate: '+1Y',
				maxDate: '+0D',
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

function fecnac(){
  if(frm.ano.value!="" && frm.mes.value!="" && frm.dia.value!=""){	
   var anio = "<%=year(date())%>";
   //var mes  = "<%=right(("0" & month(date())),2)%>";
   var mes  = "<%=month(date())%>";
   var dia  = "<%=day(date())%>";
   //alert(mes);
   var dif_anio = parseInt(anio) - parseInt(frm.ano.value);
   var dif_mes  = parseInt(mes) - parseInt(frm.mes.value);
   var dif_dia  = parseInt(dia) - parseInt(frm.dia.value);
   //alert(parseInt(mes));
   //alert(parseInt(frm.mes.value));
   //alert("dif_mes:" + dif_mes);
   if(dif_mes < 0)
  	 {
		//alert(dif_anio);	 
	 dif_anio= parseInt(dif_anio)-1;
	 
	 }
   else{if(parseInt(dif_mes)==0)
	         {
	  			if(parseInt(dif_dia)<0)
				  {dif_anio= parseInt(dif_anio)-1;}
			 }
	  	  }

   frm.edad.value = dif_anio;
   frm.fchNacimiento.value = frm.dia.value + "/" + frm.mes.value + "/" + frm.ano.value;
   //frm.edad.value = parseInt(anio) - parseInt(frm.ano.value);
   //alert(frm.edad.value);
   }
}

function solonumeros(e)
{
 var key;
 if(window.event) // IE
 {  key = e.keyCode; }
  else if(e.which) // Netscape/Firefox/Opera
 {  key = e.which; }

 if (key < 48 || key > 57) 
    {      return false;    }
 return true;
}

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

function reemplazar(input)
{s = input.value;
var i;
var returnString = "";
for (i = 0; i < s.length; i++) {  // Search through string and append to unfiltered values to returnString.
var c = s.charAt(i);
 if(c=="'"){ c="''";}
 returnString +=c;
}
 //s.replace("'", "''");
 input.value = returnString;
	}

  </script> 
</head>
<body bgcolor="#e9e9e9" style="font-size:11px;">

<form name="frm" method="post" action="actualizar_caso.asp" onsubmit="reemplazar(desDireccion);reemplazar(des_ciudad);">  
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="titulo">
    <br />
    Registro de casos
    <br />
	<br />
	</td>
   
  </tr>
   
   <tr>
    <td >
 
    	<fieldset style="width:1000px">
          <legend class="titulo_grupo" >Datos del caso:</legend>            
            <table width="100%" border="0" cellspacing="0" cellpadding="0" class="contenido_tabla1">
              <tr>
                <td width="16%">#Caso</td>
                <td><input type="text" name="idCaso" size="10" value="INV-<%=idCaso%>" readonly="readonly" /></td>
                <td>&nbsp;</td>
                <td>Tipo de caso</td>
                <td>
                <%if isArray(arrAcceso) then%>
                <select name="tipoCaso" class="contenido_tabla1"  >
                <option value="">[Seleccione]</option>
                    <option value="1" <%if tipoCaso="1" then %> selected="selected" <%end if%>>SIMPLE</option>
                    <option value="2" <%if tipoCaso="2" then %> selected="selected" <%end if%>>MEDIUM</option>
                    <option value="3" <%if tipoCaso="3" then %> selected="selected" <%end if%>>COMPLEX</option>
                </select>
                <%else
					select case tipoCaso
					case "1"
						response.Write("SIMPLE")
					case "2"
					    response.Write("MEDIUM")
					case "3"
					    response.Write("COMPLEX")
					end select	
				%>                	
                <%end if%>
                </td>
                <td>&nbsp;</td>
                <td>Empresa
                </td>
                <td>
                <select name="idEmpresa" class="contenido_tabla1">
                    <option value="">[Seleccione]</option>
                    <%if isArray(arrEmpresa) then
						for i=0 to ubound(arrEmpresa,2)%>
                    <option value="<%=arrEmpresa(0,i)%>" <%if trim(idEmpresa)=trim(arrEmpresa(0,i)) then %> selected="selected" <%end if%>><%=arrEmpresa(1,i)%></option>
					<%	next
					end if%>
                </select>
                </td>
               <td colspan="2">
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
               		<table border="0" cellspacing="0" cellpadding="0">
                      <tr>
						  <%if isArray(arrStatus) then
						  		for i=0 to ubound(arrStatus,2)
									if idstatus = arrStatus(0,i) then
										ancho = "28" 
										alto="27"
									else
										ancho = "20" 
										alto="19"	
									end if
								%>
                            	<td>
                                
                                <img src="images/banderas/<%=arrStatus(2,i)%>" width="<%=ancho%>" height="<%=alto%>" border="0" title="<%=arrStatus(1,i)%>" />
                                
                                </td>
                          <%	next %>
                      	</tr>  
                        <tr>
                        	<%for j=0 to ubound(arrStatus,2)%>
                            	<td>
                        <input type="radio" name="idStatus" value="<%=arrStatus(0,j)%>" <%if arrStatus(0,j)= idstatus then%> checked="checked" <%end if%> />
                                </td>
                            <%next%>
                        </tr>  
						  <%end if%>  
                      
                    </table>
					<%if trim(tip_usuario) = "A" then%>
						<a href="lista_cmbstatus.asp?idCaso=<%=idCaso%>" target="_blank">ver cambios</a>
					<%end if%>
               </td>
              </tr>
            </table>          
        </fieldset>
      
    </td>
   
  </tr>
  <tr>
  	<td >
    <fieldset style="width:1000px">
          <legend class="titulo_grupo" >Datos del caso:</legend>   
            <table width="100%" border="0" cellspacing="1" cellpadding="1" class="contenido_tabla1">
  <tr>
    <td width="16%">Nombre</td>
    <td align="left"><input name="desNombre" type="text" value="<%=desNombre%>" size="30" maxlength="100" onblur="limpiar(this)"/></td>
    <td>Ref. Externa</td>
    <td><input name="idRefExterna" type="text" value="<%=idRefExterna%>" size="10" maxlength="15"/></td>
    <td>&nbsp;</td>
    <td>Fecha</td>
    <td><input type="text" name="fchCaso" size="10" value="<%=FormateaFecha(fchCaso,"ddmmyyyy")%>" readonly /></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Nacimiento</td>
    <td align="left"><select name="dia" alerta="Seleccione Dia de Nacimiento" onchange="fecnac()" class="contenido_tabla1">
      <option value="">[Dia]</option>
      <option value="1" <%if trim(dia_nac)="01" or trim(dia_nac)="1" then%> selected <%end if%>>01</option>
      <option value="2" <%if trim(dia_nac)="02" or trim(dia_nac)="2" then%> selected <%end if%>>02</option>
      <option value="3" <%if trim(dia_nac)="03" or trim(dia_nac)="3" then%> selected <%end if%>>03</option>
      <option value="4" <%if trim(dia_nac)="04" or trim(dia_nac)="4" then%> selected <%end if%>>04</option>
      <option value="5" <%if trim(dia_nac)="05" or trim(dia_nac)="5" then%> selected <%end if%>>05</option>
      <option value="6" <%if trim(dia_nac)="06" or trim(dia_nac)="6" then%> selected <%end if%>>06</option>
      <option value="7" <%if trim(dia_nac)="07" or trim(dia_nac)="7" then%> selected <%end if%>>07</option>
      <option value="8" <%if trim(dia_nac)="08" or trim(dia_nac)="8" then%> selected <%end if%>>08</option>
      <option value="9" <%if trim(dia_nac)="09" or trim(dia_nac)="9" then%> selected <%end if%>>09</option>
      <%for i=10 to 31%>
      <option value="<%=i%>"  <%if trim(dia_nac)=trim(i) then%> selected <%end if%>><%=i%></option>
      <%next%>
      </select>
      <select name="mes" alerta="Seleccione Mes de Nacimiento" onchange="fecnac()" class="contenido_tabla1">
        <option value="" >[Mes] </option>
        <option value="1" <%if mes_nac="01" or mes_nac="1" then%> selected <%end if%>>Ene</option>
        <option value="2" <%if mes_nac="02" or mes_nac="2" then%> selected <%end if%>>Feb</option>
        <option value="3" <%if mes_nac="03" or mes_nac="3" then%> selected <%end if%>>Mar</option>
        <option value="4" <%if mes_nac="04" or mes_nac="4" then%> selected <%end if%>>Abr</option>
        <option value="5" <%if mes_nac="05" or mes_nac="5" then%> selected <%end if%>>May</option>
        <option value="6" <%if mes_nac="06" or mes_nac="6" then%> selected <%end if%>>Jun</option>
        <option value="7" <%if mes_nac="07" or mes_nac="7" then%> selected <%end if%>>Jul</option>
        <option value="8" <%if mes_nac="08" or mes_nac="8" then%> selected <%end if%>>Ago</option>
        <option value="9" <%if mes_nac="09" or mes_nac="9" then%> selected <%end if%>>Sep</option>
        <option value="10" <%if mes_nac="10" then%> selected <%end if%>>Oct</option>
        <option value="11" <%if mes_nac="11" then%> selected <%end if%>>Nov</option>
        <option value="12" <%if mes_nac="12" then%> selected <%end if%>>Dic</option>
      </select>
      <select name="ano" alerta="Seleccione Ano de nacimiento" onchange="fecnac()" class="contenido_tabla1">
        <option value="">[A&ntilde;o]</option>
        <%for j=year(date())-90 to year(date())%>
        <option value="<%=j%>" <%if trim(ano_nac)=trim(j) then%> selected <%end if%>><%=j%></option>
        <%next%>
      </select>
      <input type="hidden" name="fchNacimiento" id="fchNacimiento" size="10" value="<%=fchNacimiento%>" 
    maxlength="10" readonly="readonly" alerta=""/></td>
   
    <td>Edad</td>
    <td width="110"><input name="edad" type="text" value="<%=edad%>" size="3" readonly="readonly" /></td>
    <td width="120">&nbsp;</td>
    <td>Telefono</td>
    <td colspan="2">&nbsp;<input name="desTelefono" type="text" value="<%=desTelefono%>" size="25" maxlength="50"/></td>
    
  </tr>
  <tr>
    <td>Direcci&oacute;n</td>
    <td align="left">&nbsp;<input name="desDireccion" type="text" value="<%=desDireccion%>" size="30" maxlength="200" /></td>
    
    <td>Pa&iacute;s</td>
    <td colspan="2">&nbsp;<select name="idPais" class="contenido_tabla1" style="width:160px">
                    <option value="">[Seleccione]</option>
                    <%if isArray(arrPais) then
						for i=0 to ubound(arrPais,2)%>
                    <option value="<%=arrPais(0,i)%>" <%if trim(idPais)=trim(arrPais(0,i)) then %> selected="selected" <%end if%>><%=ucase(arrPais(1,i))%></option>
					<%	next
					end if%>
                </select>
    </td>
    <td>Ciudad</td>
    <td>&nbsp;<input name="des_ciudad" type="text" value="<%=des_ciudad%>" size="15" maxlength="25"  /></td>
    <td><input type="submit" name="btnGrabar" value="Grabar Caso" /></td>
    <td>&nbsp;</td>
  </tr>
</table>
 
    </fieldset>        
    </td>
  </tr>
  <%if idCaso<>"" then%>
 <tr>
 	<td>
    	<iframe src="registro_servicio.asp?idCaso=<%=idCaso%>&idUsuarioCreador=<%=idUsuarioCreador%>&empresacasoserviciomax=<%=empresacasoserviciomax%>" frameborder="0" width="100%" height="380px"
         scrolling="no"></iframe>
    </td>
 </tr>
 <tr>
 	<td>
<iframe src="registro_incidencia.asp?idCaso=<%=idCaso%>&idUsuarioCreador=<%=idUsuarioCreador%>" frameborder="0" width="100%" height="350px" scrolling="auto">
</iframe>
    </td>
 </tr>
 <%end if%> 
 </table> 
 <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
 <input type="hidden" name="idStatusAntes" value="<%=idStatusAntes%>" />
 <input type="hidden" name="tipoCasoAntes" value="<%=tipoCasoAntes%>" /> 
 <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
 <input type="hidden" name="empresacasoserviciomax" value="<%=empresacasoserviciomax%>" />
 
 <input type="hidden" name="txtModo" value="OK" />
  </form>
</body>
</html>
