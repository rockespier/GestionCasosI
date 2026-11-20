<%
' Desplegables reutilizables (RT-116, RT-117).
' Archivo solo ASCII a proposito: los acentos van como entidades HTML.

' Monedas ISO 4217. Las mas usadas primero, luego alfabetico.
Const LISTA_MONEDAS = "USD,EUR,PEN,AED,AFN,ALL,AMD,ANG,AOA,ARS,AUD,AWG,AZN,BAM,BBD,BDT,BGN,BHD,BIF,BMD,BND,BOB,BRL,BSD,BTN,BWP,BYN,BZD,CAD,CDF,CHF,CLP,CNY,COP,CRC,CUP,CVE,CZK,DJF,DKK,DOP,DZD,EGP,ERN,ETB,FJD,FKP,GBP,GEL,GHS,GIP,GMD,GNF,GTQ,GYD,HKD,HNL,HTG,HUF,IDR,ILS,INR,IQD,IRR,ISK,JMD,JOD,JPY,KES,KGS,KHR,KMF,KPW,KRW,KWD,KYD,KZT,LAK,LBP,LKR,LRD,LSL,LYD,MAD,MDL,MGA,MKD,MMK,MNT,MOP,MRU,MUR,MVR,MWK,MXN,MYR,MZN,NAD,NGN,NIO,NOK,NPR,NZD,OMR,PAB,PGK,PHP,PKR,PLN,PYG,QAR,RON,RSD,RUB,RWF,SAR,SBD,SCR,SDG,SEK,SGD,SHP,SLE,SOS,SRD,SSP,STN,SVC,SYP,SZL,THB,TJS,TMT,TND,TOP,TRY,TTD,TWD,TZS,UAH,UGX,UYU,UZS,VES,VND,VUV,WST,XAF,XCD,XCG,XOF,XPF,YER,ZAR,ZMW,ZWG"

' Imprime un <select> de monedas. Si el valor guardado no esta en la lista
' (texto libre historico), se agrega como opcion para no perderlo.
Sub ComboMoneda(nombre, valor)
	Dim codigos, cod, actual, encontrado
	actual = UCase(Trim(valor & ""))
	codigos = Split(LISTA_MONEDAS, ",")
	encontrado = (actual = "")
	Response.Write "<select name=""" & nombre & """ class=""contenido_tabla1"">"
	Response.Write "<option value=""""></option>"
	For Each cod In codigos
		Response.Write "<option value=""" & cod & """"
		If cod = actual Then
			Response.Write " selected=""selected"""
			encontrado = True
		End If
		Response.Write ">" & cod & "</option>"
	Next
	If Not encontrado Then
		Response.Write "<option value=""" & Replace(Trim(valor), """", "&quot;") & """ selected=""selected"">" & Trim(valor) & "</option>"
	End If
	Response.Write "</select>"
End Sub

' Imprime el <select> de Servicio Solicitado. Se graba el texto de la opcion.
' La preseleccion se hace en el navegador (ver script) porque los textos con
' acento se guardan/leen con otra codificacion y no se pueden comparar en VBScript.
Sub ComboServicioSolicitado(nombre, valor)
	Response.Write "<select name=""" & nombre & """ id=""" & nombre & """ class=""contenido_tabla1"" data-valor=""" & Replace(Trim(valor & ""), """", "&quot;") & """>"
	Response.Write "<option value="""">[Seleccione]</option>"
	Response.Write "<option value=""Domicilio"">Domicilio</option>"
	Response.Write "<option value=""Centro M&eacute;dico"">Centro M&eacute;dico</option>"
	Response.Write "<option value=""Telemedicina"">Telemedicina</option>"
	Response.Write "<option value=""Repatriaci&oacute;n funeraria"">Repatriaci&oacute;n funeraria</option>"
	Response.Write "</select>"
	Response.Write "<script type=""text/javascript"">(function(){" & _
		"var s=document.getElementById('" & nombre & "'),v=s.getAttribute('data-valor');if(!v)return;" & _
		"for(var i=0;i<s.options.length;i++){if(s.options[i].value.toLowerCase()==v.toLowerCase()){s.selectedIndex=i;return;}}" & _
		"var o=document.createElement('option');o.value=v;o.text=v;s.appendChild(o);s.selectedIndex=s.options.length-1;" & _
		"})();</script>"
End Sub
%>
