# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Qué es

"Sistema de Gestión de Casos" de Innova Doctors (asistencia médica / viajero): aplicación **Classic ASP (VBScript)** sobre IIS con SQL Server vía ADODB/ODBC. Sin build, sin package manager, sin tests ni linter. Se despliega copiando los archivos al hosting (Plesk, `C:\Inetpub\vhosts\euroamer-32727.package\`, dominio `gestiondecasos.innovadoctors.com`).

Para probar localmente hace falta IIS con ASP Clásico habilitado y acceso a la BD remota; no hay entorno de desarrollo propio.

## Arquitectura

- **`main.asp`** se incluye (`<!--#include file="main.asp"-->`) en casi todas las páginas. Define `strUrl` (cadena de conexión SQL Server, credenciales en claro), el objeto `cn` (ADODB.Connection, sin abrir) y helpers: `SiVacio`, `FormateaFecha(fecha, "ddmmyyyy"|"mmddyyyy"|"yyyymmdd")`, `Encriptar`/`DesEncriptar` (sustitución de caracteres, no criptografía real).
- **Patrón de página**: cada `.asp` es autocontenido — lee `Request(...)`, arma SQL por concatenación de strings, hace `cn.open strUrl` → `rs.open sql, cn` → `GetRows()`/loop → `cn.close`, y renderiza HTML en la misma página. No hay capa de datos ni plantillas compartidas.
- **Convención de nombres**: `lista_*.asp` (listados/búsqueda), `registro_*.asp` (alta), `actualizar_*.asp` (edición), `*_xls.asp` (exportación a Excel vía content-type). Entidades principales: caso, servicio/serviciocaso, incidencia, status/statusxcaso, proveedor/proveedorservicio, empresa, usuario, facturación. `suceso` es la tabla de bitácora.
- **Flujo de login**: `index.asp` (form) → `login.asp` (valida contra `usuario`, contraseña en texto plano) → `home.asp`, que construye el menú desde `accesos` + `opciones` (por usuario, `tip_opcion='P'`) y carga las opciones en un iframe (`blank.htm` por defecto).
- **Estado/identidad**: se propaga por querystring/form (`cod_usuario`, `idUsuarioCreador`, `tip_usuario`, `idCaso`). `Session("codUsr")` se setea en `login.asp` pero ninguna página lo verifica — no hay control de acceso real por página.
- **Correo**: `EnviarCorreo.asp` (form) → `correo.asp` (CDO.Message, SMTP con auth básica) → redirige de vuelta con `msj`.
- **Uploads**: `adjuntararchivo.asp` usa `freeASPUpload.asp` (librería de terceros) y guarda el nombre del archivo en `proveedor.desArchivo`.
- Frontend: jQuery plugins sueltos (`jquery.datePicker.js`, `jquery.ddslick.min.js`, `js/apprise`), CSS en `estilo.css` / `admin_login.css`.

## Advertencias

- Todo el SQL es concatenado (inyección SQL generalizada). Tenerlo presente al modificar páginas; no ampliar estos patrones. (`admsql.asp`, consola SQL arbitraria, fue eliminado; no reintroducir algo similar.)
- `main.asp` y `correo.asp` contienen credenciales hardcodeadas (BD y SMTP).
- Páginas usan `<%@LCID = 1034%>` (español) y `charset=utf-8`; respetar la codificación del archivo al editar (algunos tienen BOM).
- Archivos `*_bkp.asp`, `*2.asp`, `index_original.htm`, `prueba.html`, `default.aspx` son copias/restos; confirmar cuál está en uso antes de editar.
- `.user.ini` es config PHP del hosting, no la usa la app.
- Fechas: el servidor está en EE. UU. Toda fecha "ahora" en SQL debe ser hora de Lima: usar `DATEADD(hour,-5,GETUTCDATE())`, nunca `getdate()`/`now()` (Perú no tiene horario de verano).
