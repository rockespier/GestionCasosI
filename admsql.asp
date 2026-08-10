<%@LCID = 1034%>
<!--#include file="main.asp"-->
<%sql = Request("sql")%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<div class="header"><h3>ADO Primer: <span class="hilite">Viewing the Recordset of any Data Source</span></h3></div>

<hr size=1 width=90%>

<div align="center">

<form name="frm" method="post" action="admsql.asp">
<textarea name="sql" cols="170" rows="20"><%=sql%></textarea>
<input name="en" type="submit" value="enviar" />
</form>

<%rowCount=0
 if not isEmpty(Request("sql")) then
   strTable=Request("sql") %>
  
  <% 'create selected table recordset
   set rsTable=Server.CreateObject("ADODB.Recordset")
   cn.open strUrl
   rsTable.Open strTable, cn 
if not rsTable.eof then     
   %>

   <table cellspacing=1 cellpadding=5 border=0>
   <!-- begin column headers for selected table -->
   <tr><% for each field in rsTable.Fields %>
      <th><%= field.Name %></th>   
      <% next %></tr>
   <% 'cycle thru the record set and display each row results
 
   do until rsTable.EOF %>
      <tr><% for each field in rsTable.Fields %>
         <td><%= field.Value %></td>
      <% next %></tr>
   <!--  next row = next record -->
   <% rowCount = rowCount+1
   'increment record position with MoveNext method
   rsTable.MoveNext
   loop
  
   'close connection and flush objects from memory
   cn.Close
   set cn = Nothing
   set rsTable = Nothing %>
   </table>
   <% if rowCount>8 then %>
      <p><a href="#">Go Top</a></p>
   <% end if
 end if
end if %>
</div>

<% if not isEmpty(Request.Form) then %>
   <br>
   <hr size=1 width=90%>


<% end if %>

</body>
</html>

