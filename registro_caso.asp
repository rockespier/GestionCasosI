<%@LCID = 1034%>
<!--#include file="main.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RECORDSET")

idUsuarioCreador = request("idUsuarioCreador")
tip_usuario = request("tip_usuario")

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
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" and request("txtModo")<>"" then
	
	idEmpresa = SiVacio(request("idEmpresa"),0)
	
	anioCaso = year(date())
	mesCaso  = month(date())
	
	tipoCaso = SiVacio(request("tipoCaso"),0)
	idRefExterna = request("idRefExterna")
	desNombre = request("desNombre")
	fchNacimiento = request("fchNacimiento")

	edadTitular = anioCaso - year(fchNacimiento)
	dia_nac = request("dia")
	mes_nac = request("mes")
	ano_nac = request("ano")
		
	if mesCaso < mes_nac then
		edadTitular = edadTitular - 1
	end if
	
	desTelefono = request("desTelefono")
	desDireccion = trim(replace(request("desDireccion"),chr(34),"''"))
	idPais = request("idPais")
	idCiudad = 0
	des_ciudad = trim(replace(request("des_ciudad"),chr(34),"''"))	
	Nacimiento = FormateaFecha(fchNacimiento,"yyyymmdd")

	
		
		sql = " insert into caso("
		sql = sql & "  idEmpresa, anioCaso , tipoCaso ,  idRefExterna ,"
		sql = sql & "  fchCaso,  desNombre ,  fchNacimiento , edadTitular , desTelefono ,"
  		sql = sql & "  desDireccion ,  idPais , idCiudad ,  des_ciudad ,"
  		sql = sql & "  idUsuarioCreador , fchCreacion, dia_nac, mes_nac, ano_nac )"
		sql = sql & "  values ("
		sql = sql & idEmpresa & "," & anioCaso & "," & tipoCaso & ",'" &idRefExterna &"',getdate(),'"
		sql = sql & desNombre &"','" & Nacimiento & "'," & edadTitular & ",'" & desTelefono &"','"& desDireccion & "',"
		sql = sql & "'"& idPais & "'," & idCiudad & ",'" & des_ciudad &"','"& idUsuarioCreador &"',getdate(),"
		sql = sql & "'" & dia_nac & "','" & mes_nac &"','"& ano_nac &"')"  
		'response.Write(sql)
		'response.End()
		cn.open strUrl
		cn.execute(sql)
		cn.close
		
		'sql = "select LPAD(max(idCaso),5,'0'),max(idcaso) from caso"
		sql = "select max(idCaso),max(idcaso) from caso"
		cn.open strUrl
		rs.open sql, cn
		if not rs.eof then
		idCaso = rs.fields(0)
		id = rs.fields(1)
		end if
		rs.close()
		cn.close
		
		'idCaso = "00000" & idCaso
		'idCaso = Right(idCaso,5)
	'	sql = "insert into suceso(dessuceso,codusuario,idCaso) values('Caso registrado:NCA-"& idCaso &"','"&idUsuarioCreador&"','NCA-"&idCaso&"')"
'		cn.open strUrl
'		cn.execute(sql)
'		cn.close()
		
		sql = "insert into StatusxCaso values(1,"&id&",getdate(),'"& idUsuarioCreador &"','S');"
		cn.open strUrl
		cn.execute(sql)
		cn.close
		
		response.Redirect("actualizar_caso.asp?idCaso="&id&"&idUsuarioCreador="&idUsuarioCreador&"&tip_usuario="&tip_usuario&"")
		
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
     src="http://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.min.js"></script>
  <script type="text/javascript"
     src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.7.2/jquery-ui.min.js"></script>
  <script type="text/javascript" language="javascript">
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

function Validar(){   
var correcto = false; 
	for (i=0;i<frm.length;i++) 	
		{var tempobj=frm.elements[i];
		
			if (tempobj.alerta != "" && tempobj.type != 'hidden') {//alert(tempobj.type);						
				if (tempobj.value==''){
				alert(tempobj.alerta);				
				tempobj.focus();
				correcto = false;			
				break;	 
				}else{correcto = true;}			
			}
		}
	if (correcto)
		{
		 frm.submit();
		}
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

  </script> 
</head>
<body bgcolor="#e9e9e9" style="font-size:11px;">

<form name="frm" method="post" action="registro_caso.asp" >  
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
            <table width="100%" border="0" cellspacing="1" cellpadding="1" class="contenido_tabla1">
              <tr>
                <td width="16%">#Caso</td>
                <td><input type="text" name="idCaso" size="10" value="INV-<%=idCaso%>" readonly="readonly" alerta="" /></td>
                <td>&nbsp;</td>
                <td>Tipo de caso</td>
                <td>
                <select name="tipoCaso" class="contenido_tabla1" alerta="Seleccione tipo de caso">
                <option value="">[Seleccione]</option>
                    <option value="1" <%if tipoCaso="1" then %> selected="selected" <%end if%>>SIMPLE</option>
                    <option value="2" <%if tipoCaso="2" then %> selected="selected" <%end if%>>MEDIUM</option>
                    <option value="3" <%if tipoCaso="3" then %> selected="selected" <%end if%>>COMPLEX</option>
                </select>
                </td>
                <td>&nbsp;</td>
                <td>Empresa
                </td>
                <td>
                <select name="idEmpresa" class="contenido_tabla1" alerta="Seleccione empresa">
                    <option value="">[Seleccione]</option>
                    <%if isArray(arrEmpresa) then
						for i=0 to ubound(arrEmpresa,2)%>
                    <option value="<%=arrEmpresa(0,i)%>" <%if idEmpresa=arrEmpresa(0,i) then %> selected="selected" <%end if%>><%=arrEmpresa(1,i)%></option>
					<%	next
					end if%>
                </select>
                </td>
                
                
                <td>&nbsp;</td>
                <td>&nbsp;</td>
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
    <td width="180" align="left"><input name="desNombre" type="text" value="<%=desNombre%>" size="25" maxlength="100" alerta=""  onblur="limpiar(this)" /></td>
    <td>Ref. Externa</td>
    <td><input name="idRefExterna" type="text" value="<%=idRefExterna%>" size="10" maxlength="15" alerta=""/></td>
    <td>&nbsp;</td>
    <td>Fecha</td>
    <td><input type="text" name="fchCaso" size="8" value="<%=FormateaFecha(date(),"ddmmyyyy")%>" readonly="readonly" alerta=""/></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Nacimiento</td>
    <td  align="left"><select name="dia" alerta="Seleccione Dia de Nacimiento" onchange="fecnac()" class="contenido_tabla1">
      <option value="">[Dia]</option>
      <option value="1">01</option>
      <option value="2">02</option>
      <option value="3">03</option>
      <option value="4">04</option>
      <option value="5">05</option>
      <option value="6">06</option>
      <option value="7">07</option>
      <option value="8">08</option>
      <option value="9">09</option>
      <%for i=10 to 31%>
      <option value="<%=i%>"><%=i%></option>
      <%next%>
      </select>
      <select name="mes" alerta="Seleccione Mes de Nacimiento" onchange="fecnac()" class="contenido_tabla1">
        <option value="" >[Mes] </option>
        <option value="1">Ene</option>
        <option value="2">Feb</option>
        <option value="3">Mar</option>
        <option value="4">Abr</option>
        <option value="5">May</option>
        <option value="6">Jun</option>
        <option value="7">Jul</option>
        <option value="8">Ago</option>
        <option value="9">Sep</option>
        <option value="10">Oct</option>
        <option value="11">Nov</option>
        <option value="12">Dic</option>
      </select>
      <select name="ano" alerta="Seleccione Ano de nacimiento" onchange="fecnac()" class="contenido_tabla1">
        <option value="">[A&ntilde;o]</option>
        <%for i=year(date())-90 to year(date())%>
        <option value="<%=i%>"><%=i%></option>
        <%next%>
      </select>
      <input type="hidden" name="fchNacimiento" id="fchNacimiento" size="10" value="<%=fchNacimiento%>" 
    maxlength="10" readonly="readonly" alerta=""/></td>
   
    <td>Edad      </td>
    <td width="110"><input name="edad" type="text" value="<%=edad%>" size="3" readonly="readonly" alerta="" /></td>
    <td>&nbsp;</td>
    <td>Telefono</td>
 <td colspan="3"><input type="text" name="desTelefono" size="25" value="<%=desTelefono%>" maxlength="50" alerta=""/></td>

  </tr>
  <tr>
    <td>Direcci&oacute;n</td>
    <td align="left">&nbsp;<input name="desDireccion" type="text" value="<%=desDireccion%>" size="30" maxlength="200" alerta=""  onblur="limpiar(this);" /></td>
    
    <td>Pa&iacute;s</td>
    <td colspan="2"><select name="idPais" class="contenido_tabla1" alerta=""  style="width:160px">
                    <option value="">[Seleccione]</option>
                    <%if isArray(arrPais) then
						for i=0 to ubound(arrPais,2)%>
                    <option value="<%=arrPais(0,i)%>" <%if idPais=arrPais(0,i) then %> selected="selected" <%end if%>><%=ucase(arrPais(1,i))%></option>
					<%	next
					end if%>
                </select>
    </td>
    <td>Ciudad</td>
    <td width="110">&nbsp;<input name="des_ciudad" type="text" value="<%=des_ciudad%>" size="15" maxlength="25" alerta=""  onblur="limpiar(this);" /></td>
    <td>&nbsp;</td>
    <td><input type="submit" name="btnGrabar" value="Grabar Caso"  /></td>
  </tr>
</table>
 
    </fieldset>        
    </td>
  </tr>
  <%if idCaso<>"" then%>
 <tr>
 	<td>
    	<iframe src="registro_servicio.asp?idCaso=<%=idCaso%>&idUsuarioCreador=<%=idUsuarioCreador%>" frameborder="0" width="100%" height="380px"></iframe>
    </td>
 </tr>
 <%end if%> 
 </table> 
 <input type="hidden" name="idUsuarioCreador" value="<%=idUsuarioCreador%>" />
 <input type="hidden" name="tip_usuario" value="<%=tip_usuario%>" />
 <input type="hidden" name="txtModo" value="OK" />
  </form>
</body>
</html>
