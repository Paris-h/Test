
<%@ Language=VBScript %>
<%
	Refresh = Day(Date()) & Hour(Time()) & Minute(Time()) & Second(Time())
	Response.Buffer = True
	Response.CacheControl = "Private"
	Response.Expires = 0
	on error resume next
	
	
%>

<!--- #include file="conexion.asp" --->

<%
	
	dia=day(now())
	mes =Month(now())
	ano = Year(now())
	ndia = Weekday(now())
	hora= hour(time)
	min= minute(time)
	
	if mes = 1 then mes1="Enero" end if
	if mes = 2 then mes1="Febrero" end if
	if mes = 3 then mes1="Marzo" end if
	if mes = 4 then mes1="Abril" end if
	if mes = 5 then mes1="Mayo" end if
	if mes = 6 then mes1="Junio" end if
	if mes = 7 then mes1="Julio" end if
	if mes = 8 then mes1="Agosto" end if
	if mes = 9 then mes1="Septiembre" end if
	if mes = 10 then mes1="Octubre" end if
	if mes = 11 then mes1="Noviembre" end if
	if mes = 12 then mes1="Diciembre" end if
	if ndia = 1 then ndia1="Domingo" end if
	if ndia = 2 then ndia1="Lunes" end if
	if ndia = 3 then ndia1="Martes" end if
	if ndia = 4 then ndia1="Miércoles" end if
	if ndia = 5 then ndia1="Jueves" end if
	if ndia = 6 then ndia1="Viernes" end if
	if ndia = 7 then ndia1="Sábado" end if
	
	fecha = ndia1 & " " & dia & " de " &mes1 & " de " & ano 
	
	' Acciones
	
	If Clng(Request("agregar.y")) > 0  Then 
		if Request.Form("nombre") <> "" then
			val = val + 1
		else
			error = true
			texto= texto & "&nbsp;" & " * Olvido escribir el nombre del cliente" & "<br>"
		end if
		
		op2="Select cliente_Nombre From cliente Where cliente_Nombre = '"&ucase(trim(Request.Form("nombre")))&"'"
		rs2.open op2,conn,2,2
		if rs2.EOF then
			error=false
		else
			error=true
			texto= texto & "&nbsp;" & " * El cliente ya existe en el sistema" & "<br>"
		end if
		rs2.close
		
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
		if Request.Form("rfc") <> "" then
			val = val + 1
		else
			error = true
			texto= texto & "&nbsp;" & " * Olvido escribir el RFC" & "<br>"
		end if
		

		if error = false then
			op1="Select * From Cliente"
			RS1.Open op1,conn,3,2
				RS1.AddNew
				rs1("Cliente_Nombre")= ucase(Request.Form("nombre"))
				rs1("Cliente_Direccion")= Request.Form("direccion")
				rs1("Cliente_Telefono")= Request.Form("tel")
				rs1("Cliente_Fax")= Request.Form("fax")
				rs1("Cliente_rfc")= Request.Form("rfc")
				Rs1.Update 
				RS1.Close
				Mensaje = true
				texto_mensaje = " El cliente se dio de alta con éxito en el sistema" & "<br>"	
		end if	
	end if 
	
	If Clng(Request("modificar.y")) > 0  Then 
		if Request.Form("nombre") <> "" then
			val = val + 1
		else
			error = true
			texto= texto & "&nbsp;" & " * Olvido escribir el nombre del proveedor" & "<br>"
		end if
		if Request.Form("rfc") <> "" then
			val = val + 1
		else
			error = true
			texto= texto & "&nbsp;" & " * Olvido escribir el RFC" & "<br>"
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
			op1="Select * From cliente Where cliente_iD= " & Request("prov_id")
			RS1.Open op1,conn,3,2
				rs1("Cliente_Nombre")= ucase(Request.Form("nombre"))
				rs1("Cliente_Direccion")= Request.Form("direccion")
				rs1("Cliente_Telefono")= Request.Form("tel")
				rs1("Cliente_Fax")= Request.Form("fax")
				rs1("Cliente_rfc")= Request.Form("rfc")
				Rs1.Update 
				RS1.Close
				Mensaje = true
				texto_mensaje = " El proveedor se actualizo con éxito" & "<br>"	
		end if	
	end if 
	
	If Clng(Request("eliminar.y")) > 0  Then 
'		op1="Select * From tblCompany Where IId = " & Request.Form("fabricante2")
'				RS1.Open op1,conn,3,2
'					rs1("bActivo")= 1 ' Baja lógica
'					Rs1.Update 
'					RS1.Close
'					Mensaje = true
'					texto_mensaje = "La marca se dio de baja del sistema con éxito" & "<br>"	
	
	end if
%>


<html>
<head>
<title>Clientes</title>
<Script language=JavaScript>
		function CRT()

		{
		document.forms[0].submit();	
		}
</Script>
</head>
<body bgcolor="#ffffff" topmargin=0 leftmargin=0>
<basefont size=1 face=Verdana>
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
						<td colspan=4 bgcolor=#7BA5C3 valign=middle>
							<font face=verdana size=2 color=white><b>&nbsp;&nbsp;Clientes
						</td>
					</tr>

				<tr>
					<td width=120 valign=top>
						<font face=arial size=2>&nbsp;Disponibles:
					</td>
					<td valign=top>
						<select name="proveedor" size=6 onchange="CRT();">
							<%
								op="Select * From Cliente order by Cliente_Nombre"
								rs.Open op,conn,2,2
								if rs.EOF then
							%>
								<option value="0">No existe ningún cliente dado de alta en el sistema
							<%
								else
								While not RS.EOF 
							%>
							
								<option value="<%=RS("Cliente_iD")%>"><%=RS("Cliente_Nombre")%>
							<%
								RS.MoveNext 
								Wend
								end if
								RS.Close
								
							%>
						</select>
					</td>
				</tr>
				
			<%if Request.Form("proveedor") > 0  then%>		
				<%
					op="Select * From Cliente Where Cliente_iD = " & Request.Form("proveedor")
					RS.Open op, conn,2,2
					nombre_cliente = RS("Cliente_Nombre")
					direccion= rs("Cliente_Direccion")
					tel = rs("Cliente_Telefono")
					fax = rs("Cliente_Fax")
					rfc = rs("Cliente_RFC")
					RS.Close
				%>
			<%end if%>
			<input type=hidden name=prov_id value=<%=Request.Form("proveedor")%>>	
				<tr>
					<td width=120>
						<font face=arial size=2><br>&nbsp;&nbsp;Nombre*:
					</td>
					<td>
						<br><input type=text name=nombre size=40 maxlength=150 value="<%=trim(nombre_cliente)%>">
					</td>
				</tr>
				<tr valign=top height=20>
						<td valign=top>
							<font face=arial size=2>&nbsp;&nbsp;Dirección*:
						</td>
						<td>
							<input type=text name=direccion size=50 maxlength=200 value="<%=trim(direccion)%>">
						</td>	
					</tr>
					<tr valign=top height=20>
						<td valign=top>
							<font face=arial size=2>&nbsp;&nbsp;Teléfono*:
						</td>
						<td>
							<input type=text name=tel size=15 maxlength=20 value="<%=trim(tel)%>">
						</td>	
					</tr>
					<tr valign=top height=20>
						<td valign=top>
							<font face=arial size=2>&nbsp;&nbsp;Fax:
						</td>
						<td>
							<input type=text name=fax size=15 maxlength=20 value="<%=trim(fax)%>">
						</td>	
					</tr>
					<tr valign=top height=20>
						<td width=100 valign=top>
							<font face=arial size=2>&nbsp;&nbsp;RFC*:
						</td>
						<td>
							<input type=text name=rfc size=30 maxlength=50 value="<%=trim(rfc)%>">
						</td>	
					</tr>
				<tr>
					<td width=120>
						<font face=arial size=2>&nbsp;&nbsp;Acción:
					</td>
					<td>
						<br><input name="agregar" type=image src="imagenes/agregar.gif"><input name="modificar" type=image src="imagenes/modificar.gif"><!--<input name="eliminar" type=image src="imagenes/eliminar1.gif">-->
					</td>
				</tr>
				<input type=hidden name="fabricante1" value=<%=Request.Form("fabricante")%>>
				<input type=hidden name="fabricante2" value=<%=Request.Form("fabricante")%>>
			
			<%if error=true then %>
				<tr>
					<td><br><font face=arial size=2>&nbsp;</td>
					<td>
						<br>
						<font size=2 color=red face=arial><b><%=texto%>
					</td>
				</tr>
			<%end if%>
			<%if mensaje=true then %>
				<tr>
					<td><br><font face=arial size=2>&nbsp;</td>
					<td>
						<br>
						<font size=2 color=navy face=arial><b><%=texto_mensaje%>
					</td>
				</tr>
			<%end if%>
			</table>		
		</td>
		<!--- Fin de Contenido --->
	</tr>
	<tr>
		<td colspan=2><img src="imagenes/buttom.gif"></td>
	</tr>
</table>
</html>
