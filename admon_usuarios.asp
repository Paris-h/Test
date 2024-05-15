<%@ Language=VBScript%>

<%
	Refresh=Day(Date()) & Hour(Time()) & Minute(Time()) & Second(Time())
	if session("userid") =  "" then 
		Response.Redirect "index.asp?Refresh=" & Refresh & "&Mensaje=" & Server.UrlEncode("Su sesión ya no es valida.")
	end if
	Response.Buffer = True
	Response.CacheControl = "Private"
	Response.Expires = -1000
%>

<html>
<head>
<title>Sony (IQScout)</title>
</head>
<link href="ismx.css" rel=stylesheet>
<body bgcolor="#ffffff" topmargin=0 leftmargin=0 bottommargin=0 rightMargin=0>
<table width=800 border=0 cellpadding=0 cellspacing=0>
	<tr>
		<td colspan=3>
			<img src="imagenes/banner_top2.gif">
		</td>
	</tr>
	<tr>
		<!--- menú --->
		<td valign=top width=150>
			<!--- #include file="menu_lateral.asp" --->
		</td>
		<!--- fin menu--->
		<!--- inicio espacio en blaco--->
		<td valign=top width=10>&nbsp;</td>
		<!--- fin espacio --->
		<!---Inicio Contenido --->
		<td valign=top width=650>
			<b><br><br>
			
			<dl><dl>En Construcción
		</td>
		<!--- Fin de Contenido --->
	</tr>
</table>
</body>
</html>
