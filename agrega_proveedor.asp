<%@ Language=VBScript %>

<%
	Refresh = Day(Date()) & Hour(Time()) & Minute(Time()) & Second(Time())
	Response.Buffer = True
	Response.CacheControl = "Private"
	Response.Expires = 0
	on error resume next
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.RecordSet")
	Set rs1 = Server.CreateObject("ADODB.RecordSet")
	Set rs2 = Server.CreateObject("ADODB.RecordSet")
	Conn.Open "rayo"

	
	' Acciones
	If Clng(Request("continuar.y")) > 0  Then 
	
		if Request.Form("nombre") <> "" then
			val = val + 1
		else
			error = true
			texto= texto & "&nbsp;" & " * Olvido escribir el nombre del proveedor" & "<br>"
		end if
		
		op2="Select * From Proveedor where Proveedor_Nombre = '" & ucase(Request.Form("nombre")) & "'"
		rs2.open,conn,2,2
		if rs.EOF then
			error=false
		else
			error=true
			texto= texto & "&nbsp;" & " * El proveedor ya existe en el sistema" & "<br>"
		end if
		
		if Request.Form("direccion") <> "" then
			val = val + 1
		else
			error = true
			texto= texto & "&nbsp;" & " * Olvido escribir la dirección" & "<br>"
		end if
		if Request.Form("tel") <> "" then
			val = val + 1
		else
			error = true
			texto= texto & "&nbsp;" & " * Olvido escribir el teléfono" & "<br>"
		end if
		

		if error = false then
			op1="Select * From Proveedor"
			RS1.Open op1,conn,3,2
				RS1.AddNew
				rs1("Proveedor_Nombre")= ucase(Request.Form("nombre"))
				rs1("Proveedor_Direccion")= Request.Form("direccion")
				rs1("Proveedor_Telefono")= Request.Form("tel")
				rs1("Proveedor_Fax")= Request.Form("fax")
				rs1("Proveedor_Contacto")= Request.Form("contacto")
						
				Rs1.Update 
				RS1.Close
				Mensaje = true
				texto_mensaje = " El proveedor se dio de alta con éxito en el sistema" & "<br>"	
				
		end if	
	end if 
%>
<HTML>
<HEAD>
</HEAD>
<BODY>
<basefont size=1 face=Verdana>
<body bgcolor="#ffffff" topmargin=0 leftmargin=0>
<title>Rayo Volks</title>
<table width=800 border=0 cellpadding=0 cellspacing=0>
	<tr valign=top>
		<td colspan=2>
			<img src="imagenes/banner_top.gif">
		</td>
	</tr>
	<tr>
		<td colspan=2 align=right><font face=verdad size=1><%=fecha%></td>
	</tr>
	<tr valign=top>
		<td valign=top width=135>
			<!--#include file="librerias/menu.asp"-->
		</td>
		<td valign=top width=665 bgcolor=white height=430><br><br>
			<form method="post" action="<%=Request.ServerVariables("URL")%>?Refresh=<%=refresh%>" id=form1 name=form1>
			<table width=665 height=337 border=0 cellpadding=0 cellspacing=0 bgcolor=#B1CDDF>
					<tr height=25>
						<td colspan=2 bgcolor=#7BA5C3 valign=middle>
							<font face=verdana size=2 color=white><b>&nbsp;&nbsp;Alta de Proveedores
						</td>
					</tr>	
					<tr valign=top height=20>
						<td width=100 valign=top><br><br>
							<font face=arial size=2>&nbsp;&nbsp;Nombre*:
						</td>
						<td><br><br>
							<input type=text name=nombre size=30 maxlength=150 value="<%=request.form("nombre")%>">
						</td>	
					</tr>
					<tr valign=top height=20>
						<td valign=top>
							<font face=arial size=2>&nbsp;&nbsp;Dirección*:
						</td>
						<td>
							<input type=text name=direccion size=30 maxlength=200 value="<%=request.form("direccion")%>">
						</td>	
					</tr>
					<tr valign=top height=20>
						<td valign=top>
							<font face=arial size=2>&nbsp;&nbsp;Teléfono*:
						</td>
						<td>
							<input type=text name=tel size=15 maxlength=20 value="<%=request.form("tel")%>">
						</td>	
					</tr>
					<tr valign=top height=20>
						<td valign=top>
							<font face=arial size=2>&nbsp;&nbsp;Fax:
						</td>
						<td>
							<input type=text name=fax size=15 maxlength=20 value="<%=request.form("fax")%>">
						</td>	
					</tr>
					<tr valign=top height=20>
						<td width=100 valign=top>
							<font face=arial size=2>&nbsp;&nbsp;Contacto:
						</td>
						<td>
							<input type=text name=contacto size=30 maxlength=150 value="<%=request.form("contacto")%>">
						</td>	
					</tr>
					<tr valign=top align=right>
						<td colspan=2>
							<br><input type="image" name="continuar" src="imagenes/agregar.gif" border=0 hspace=50></a>
						</td>
					</tr>
				<%if error=true then %>
					<tr valign=top>
						<td valign=top colspan=2>
							<font size=2 color=red face=arial><b><%=texto%>
						</td>
					</tr>
				<%end if%>
				<%if mensaje=true then %>
					<tr valign=top>
						<td colspan=2>
							<font size=2 color=navy face=arial><b>&nbsp;<%=texto_mensaje%>
						</td>
					</tr>
				<%end if%>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan=2><img src="imagenes/buttom.gif"></td>
	</tr>
</table>
</html>