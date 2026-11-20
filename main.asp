<%
Dim cn, ruta, correo_notifica

'dbhost = "204.93.172.30"
'dbname_and_username = "euroamer_network"
'dbpwd = "network!1"
'
'strUrl = "Driver={MySQL ODBC 3.51 Driver};Server="& dbhost &";Port=3306;uid="& dbname_and_username &";pwd="& dbpwd &";database="& dbname_and_username &";"

strUrl = "Driver={SQL Server};Server=p1427.use1.mysecurecloudhost.com;Database=euroamer_innovadoctors;uid=euroamer_admin_2008;pwd=2008SqlAdmin!"

'strUrl = "Driver={SQL Server};Server=184.168.194.62;Database=medglobalassist_bd;uid=medglobal_admin;pwd=2008SqlAdmin!"

Set cn = Server.CreateObject("ADODB.Connection")


correo_online = "info@innovadoctors.com"
correo_emisiones = "info@innovadoctors.com"
cc=99

Function SiVacio(variable,valor)		
		if variable="" or isnull(variable) then
		variable=valor
		end if
		SiVacio=variable
End function

' Escapa texto libre para mostrarlo en HTML conservando saltos de linea (RT-113).
' No usa Server.HTMLEncode porque romperia los acentos (se guardan como bytes UTF-8).
' Los <br> de registros antiguos se respetan.
Function TextoHTML(texto)
	Dim t
	t = texto & ""
	t = Replace(t, "&", "&amp;")
	t = Replace(t, "<", "&lt;")
	t = Replace(t, ">", "&gt;")
	t = Replace(t, "&lt;br&gt;", "<br>", 1, -1, vbTextCompare)
	t = Replace(t, "&lt;br/&gt;", "<br>", 1, -1, vbTextCompare)
	t = Replace(t, "&lt;br /&gt;", "<br>", 1, -1, vbTextCompare)
	t = Replace(t, vbCrLf, "<br>")
	t = Replace(t, vbCr, "<br>")
	t = Replace(t, vbLf, "<br>")
	TextoHTML = t
End Function

' Escapa texto para meterlo dentro de un <textarea> (conserva saltos de linea).
Function TextoTextarea(texto)
	Dim t
	t = texto & ""
	t = Replace(t, "&", "&amp;")
	t = Replace(t, "<", "&lt;")
	t = Replace(t, ">", "&gt;")
	TextoTextarea = t
End Function

Function FormateaFecha(fecha,formato)

if fecha <>"" then
	'Descomponerla
	dia = "00" & day(fecha)
	mes = "00" & month(fecha)
	ano = year(fecha)

  		if formato = "ddmmyyyy" then
			fecha = right(dia,2) &"/"& right(mes,2) & "/"& ano
		end if
		if formato = "mmddyyyy" then
			fecha = right(mes,2) &"/"& right(dia,2) & "/"& ano
		end if	
		if formato = "yyyymmdd" then
			fecha = ano & "/" & right(mes,2) &"/"& right(dia,2)  
		end if	
end if
	
	FormateaFecha = fecha

End function

Function Encriptar(palabra)
  NuevaPalabra = ""
  longitud = len(palabra)  
  cont1 = 0
  cont2 = 0
  cont3 = 0
  cont4 = 0
  cont5 = 0
  for i=1 to longitud
      car = mid(palabra,i,1)
	  select case car
	  case "a"
		  if cont1=0 then
		      car = "}"
		  else
			  car = ":"
		  end if
		  cont1 = cont1 +1	      
	  case "e"
	      car = "-"	  
		  if cont2=0 then
		      car = "-"
		  else
			  car = "~"
		  end if
		  cont2 = cont2 +1		  
	  case "i"
		  if cont3=0 then
		      car = "^"
		  else
			  car = "*"
		  end if
		  cont3 = cont3 +1		  
	  case "o"
		  if cont4=0 then
		      car = ")"
		  else
			  car = "_"
		  end if
		  cont4 = cont4 +1
	  case "u"
		  if cont5=0 then
		      car = "("
		  else
			  car = "|"
		  end if
		  cont5 = cont5 +1
	  case "r"
	      car = "$"
	  case "c"
	      car = "@"
	  case "k" 
	      car = "{"
	  end select
	  NuevaPalabra = NuevaPalabra & car
  next  
  Encriptar = NuevaPalabra 
End Function

Function DesEncriptar(palabra)
  NuevaPalabra = ""
  longitud = len(palabra)  
  for i=1 to longitud
      car = mid(palabra,i,1)
	  select case car
	  case "}",":"
	      car = "a"
	  case "-","~"
	      car = "e"	  
	  case "^","*"
	      car = "i"
	  case ")","_"
	      car = "o"
	  case "(","|"
	      car = "u"
	  case "$"
	      car = "r"
	  case "@"
	      car = "c"
	  case "{" 
	      car = "k"
	  end select
	  NuevaPalabra = NuevaPalabra & car
  next  
  DesEncriptar = NuevaPalabra 

End Function

%>