<% 
Response.Expires = -1
Server.ScriptTimeout = 600
%>
<!--#include file="main.asp"-->
<!-- #include file="freeaspupload.asp" -->
<%
' ****************************************************
' Cambiar el valor de la siguiente variable
' para indicar el directorio de destino.
' El directorio indicado debe tener permisos de escritura
' de caso contrario el script fallará mostrando un error.
  Dim uploadsDirVar
  uploadsDirVar =  Server.MapPath("\eaa\network\files") 
'"\\shared.hosting.local\nfs\cust\1\32\70\507231\web\files\"
' ****************************************************

idProveedor = request("idProveedor")

contador=1

function SaveFiles
    Dim Upload, fileName, fileSize, ks, i, fileKey, resumen
    Set Upload = New FreeASPUpload
    Upload.Save(uploadsDirVar)
	' If something fails inside the script, but the exception is handled
	If Err.Number <> 0 then Exit function
		SaveFiles = ""
		ks = Upload.UploadedFiles.keys
		if (UBound(ks) <> -1) then
			resumen = "<B>Archivos subidos:</B> "
			for each fileKey in Upload.UploadedFiles.keys
				
					strSql=""
					strSql= strSql & "update proveedor set desArchivo='" & Upload.UploadedFiles(fileKey).FileName & "'"	
					strSql = strSql & " where idProveedor=" & idProveedor
					
					cn.open strUrl
					cn.Execute strSql
					cn.close
					resumen = resumen & Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "
					contador=contador+1				
			next
		else
			resumen = "El nombre del archivo especificado en el formulario no es valido en el sistema."
		end if
	'comentar la siguiente linea si no se desea mostrar el resumen
	SaveFiles = resumen
end function


if Request.ServerVariables("REQUEST_METHOD") = "POST" then

	response.write SaveFiles()
	'cn.close
	Set cn = nothing

end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>::: Network Care Assist ::: </title>
<LINK href="estilo.css" type=text/css rel=stylesheet>
</head>
<BODY>
<FORM METHOD="POST" ENCTYPE="multipart/form-data" 
 ACTION="adjuntararchivo.asp?idProveedor=<%=idProveedor%>">
	  Archivo: &nbsp;<INPUT NAME="FILE1" TYPE="FILE" class="texto" SIZE="40"><br>
       <INPUT TYPE=SUBMIT VALUE="Guardar" class="boton">
</FORM>
<br>

</body>
</html>

