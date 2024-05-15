<%@ Language=VBScript %>

<%
	Refresh = Day(Date()) & Hour(Time()) & Minute(Time()) & Second(Time())
	Response.Buffer = True
	Response.CacheControl = "Private"
	Response.Expires = 0
	'on error resume next

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
	session ("paris") = 0
	
	' Consulta
	If Clng(Request("buscar.y")) > 0  then
	
	
		if Request.Form("claveproducto") = "" then
			error=true
			txt_error = txt_error & "Olvido escribir la clave del producto" & "<br>"
			session ("buscar")=false
			session ("paris") = 1
			
			if Request.Form("claveproducto_cb") = "" then
				error=true
				txt_error = txt_error & "Olvido escribir el código de barras" & "<br>"
				session ("buscar")=false
				session ("paris") = 1	
			else
				session ("buscar")=true
				error=false
				session ("paris") = 2
			end if
			
		else
			session ("buscar")=true
			error=false
			session ("paris") = 2
		end if
	
		if session ("paris") = 2 then
			tipo=0
			
			 if Request.Form("claveproducto") = "" then
			  tipo=2
			 end if
			 
			 if Request.Form("claveproducto_cb") = "" then
			  tipo=1
			 end if
			
			if tipo=1 then
				op="Select * from ProductoInventario where Producto_Clave = " & "'" & ucase(Request.Form("claveproducto"))& "'"
				rs.Open op, conn,2,2
				if rs.EOF then
					ver_info=false 
				else
					nuevo = false
					session ("prod_id")=rs("producto_id")
					ProductoId=rs("producto_id")
					inventarioId = rs("Inventario_Id")
					ProdPrecioVta = rs("producto_precioVenta")
					ProdCantidad = rs("inventario_cantidad")
					ProDescrip = rs("Producto_Descripcion")
					ver_info=true 
				end if
				rs.close
			end if	
			
			if tipo=2 then
				op="Select * from ProductoInventario where Producto_ClaveCB = " & "'" & ucase(Request.Form("claveproducto_cb"))& "'"
				rs.Open op, conn,2,2
				if rs.EOF then
					ver_info=false 
				else
					nuevo = false
					session ("prod_id")=rs("producto_id")
					ProductoId=rs("producto_id")
					inventarioId = rs("Inventario_Id")
					ProdPrecioVta = rs("producto_precioVenta")
					ProdCantidad = rs("inventario_cantidad")
					ProDescrip = rs("Producto_Descripcion")
					ver_info=true 
				end if
				rs.close
			end if
		end if
	End if
	
	
	If Clng(Request("actualizar.y")) > 0  then
	
		op="Select * from ProductoInventario Where Producto_Id =" & session("prod_id")
		rs.Open op, conn,3,2
			rs("Inventario_Cantidad") = request.form("cantidad")
			rs("Producto_PrecioVenta") = request.form("precio_compra")
			rs("Producto_Descripcion") = request.form("descripcion")
		rs.update
		Rs.close
		
		mensaje=true
		txt_mensaje = txt_mensaje & "El Stock se actualizo con éxito" & "<br>"
		
		
	End If
%>


<html>
<head>
<title>Agregar / Editar Productos</title>
<basefont size=1 face=Verdana>
<Script language=JavaScript>
		function CRT()

		{
		document.forms[0].submit();	
		}
</Script>
</head>
<body bgcolor="#ffffff" topmargin=0 leftmargin=0>
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
						<td colspan=4 bgcolor=#7BA5C3 valign=middle height=25>
							<font face=verdana size=2 color=white><b>&nbsp;&nbsp;Agregar / Editar Productos
						</td>
					</tr>
					<tr>
						<td width=120 valign=top>
							<font face=arial size=2><b>&nbsp;Buscar:
						</td>
						<td valign=top>
								<font face=verdana size=2>
								&nbsp;&nbsp;Clave Producto:&nbsp;&nbsp;&nbsp;&nbsp;<input type=text name=claveproducto size=30 maxlength=50 value="<%=Request.Form("claveproducto")%>">
								<br><br>
								&nbsp;&nbsp;Código de barras:&nbsp;&nbsp;<input type=text name=claveproducto_cb size=30 maxlength=50 value="<%=Request.Form("claveproducto_cb")%>">
								<br><br>
									&nbsp;&nbsp;<input type=image src="imagenes/buscar.gif" border=0 name=buscar>
						</td>
					</tr>
					
		<%if ver_info=true then%>
				<tr>
					<td colspan=4 height=200 valign=top>
						<table>
							<tr>
								<td width=130  valign=middle>
									<font size=2 face=verdana>Precio
								</td>
								<td valign=top>
									<input type="text" name=precio_compra size=10 maxlength=50 value="<%=ProdPrecioVta%>" >
								</td>
							</tr>
							<tr>
								<td width=130  valign=middle>
									<font size=2 face=verdana>Cantidad
								</td>
								<td valign=top>
									<input type="text" name=cantidad size=10 maxlength=50 value="<%=Request.Form("cantidad")%>">&nbsp; <font size=1 face=verdana><b>Stock Actual: <%=(ProdCantidad)%>
								</td>
							</tr>
							<tr>
								<td width=130  valign=middle>
									<font size=2 face=verdana>Descripción
								</td>
								<td valign=top>
									<input type="text" name=descripcion size=70 maxlength=150 value="<%=ProDescrip%>">
								</td>
							</tr>
								
							
							<tr>
								<td width=130  valign=middle>
									
								</td>
								<td valign=top>
									<input type=image src="imagenes/actualizar.gif" border=0 name=actualizar>
								</td>
							</tr>
							
						<%end if%>
							<%if error=true then%>
							<tr>
								<td colspan=2>
								<dl><dl>
								<font color=red face=verdana size=1><b><%=txt_error%></b></font>
								</td>
							</tr>
							<%end if%>
							<%if mensaje=true then%>
							<tr>
								<td colspan=2>
								<dl><dl>
								<font color=navy face=verdana size=1><b><%=txt_mensaje%></b></font>
								</td>
							</tr>
							<%end if%>
							
						</table>
					</td>
				</tr>
			</table>	
		</td>
	</tr>
</table>
</html>
