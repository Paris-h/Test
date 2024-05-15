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
	
	' Alta 
	If Clng(Request("continuar.y")) > 0  then
	    if Request.Form("clave") = "" then
			error=true
			txt_error = txt_error & "Olvido escribir la clave del producto" & "<br>"
		end if
		if Request.Form("nombre") = "" then
			error=true
			txt_error = txt_error & "Olvido escribir el nombre" & "<br>"
		end if
		if Request.Form("descripcion")	= "" then
			error=true
			txt_error = txt_error & "Olvido escribir la especificación del producto" & "<br>"
		end if
		if not IsNumeric(Request.Form("precio")) then
			error=true
			txt_error = txt_error & "Olvido escribir el precio de compra o el dato no es númerico" & "<br>"
		end if
		
	
	
		if error=false then
		
			op="Select * from Catalogo where Catalogo_Clave = " & "'" & ucase(Request.Form("clave"))& "'"
			rs.Open op, conn,2,2
			if rs.EOF then
				nuevo = true
			else
				nuevo = false
				clave_existe= rs("Catalogo_ID")
			end if
			rs.close
		
			if nuevo = true then
				op1="Select * from Catalogo"
				rs1.Open op1, conn,3,2
				rs1.AddNew 
					rs1("Catalogo_Clave") = Request.Form("clave")
					rs1("Catalogo_Nombre") = Request.Form("nombre")
					rs1("Catalogo_MarcaID") = Request.Form("marca")
					rs1("Catalogo_Especificaciones") = Request.Form("descripcion")
					rs1("Catalogo_Precio") = Request.Form("precio")
				rs1.Update
				rs1.Close
				mensaje = true
				txt_mensaje ="El producto se dio de alta en el catalogo"
			else
				op1="Select * from Catalogo Where Catalogo_ID="& clave_existe
				rs1.Open op1, conn,3,2
					rs1("Catalogo_Clave") = Request.Form("clave")
					rs1("Catalogo_Nombre") = Request.Form("nombre")
					rs1("Catalogo_MarcaID") = Request.Form("marca")
					rs1("Catalogo_Especificaciones") = Request.Form("descripcion")
					rs1("Catalogo_Precio") = Request.Form("precio")
				rs1.Update
				rs1.Close
				mensaje = true
				txt_mensaje ="El producto se modifico en el cátalogo"
			end if
			
		end if
	End if
%>


<html>
<head>
<title>Rayo Volks - Cátalo de Productos</title>
<basefont size=1 face=Verdana>
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
				<tr valign=top>
					<td colspan=4 height=200 valign=top>
						<table width=665>
							<tr height=25>
								<td colspan=2 bgcolor=#7BA5C3 valign=middle height=25>
									<font face=verdana size=2 color=white><b>&nbsp;&nbsp;Cátalogo de Productos
								</td>
							</tr>
							<tr>
								<td width=130  valign=middle>
									<font size=1><br>Clave del Producto
								</td>
								<td valign=top><br>
									<input type="text" name=clave size=30 maxlength=50  value="<%=Request.Form("clave")%>">
								</td>
							</tr>
							<tr>
								<td width=130  valign=middle>
									<font size=1>Marca
								</td>
								<td valign=top>
									<select name="marca" size=1 >
										<%
											op="Select * From Marca order by Marca_nombre"
											rs.Open op,conn,2,2
											if rs.EOF then

											else%>
										<%	
											While not RS.EOF 
											
											If CInt(Request("Marca")) = RS("Marca_Id") Then
											Seleccionado = " Selected"
											Else 
												Seleccionado = ""
											End If
										%>
										
											<option value="<%=RS("Marca_Id")%>" <%=Seleccionado%>><%=RS("Marca_nombre")%>
										<%
											RS.MoveNext 
											Wend
											end if
											RS.Close
										%>
									</select>
								</td>
							</tr>
							<tr>
								<td width=130  valign=middle>
									<font size=1>Nombre
								</td>
								<td valign=top>
									<input type="text" name=nombre size=70 maxlength=150 value="<%=Request.Form("nombre")%>">
								</td>
							</tr>
							<tr>
								<td width=130  valign=top>
									<font size=1>Especificaciónes
								</td>
								<td valign=top>									<TEXTAREA rows=3 cols=53  name=descripcion wrap=virtual><%=Request.Form("descripcion")%></TEXTAREA>
								</td>
							</tr>
							<tr>
								<td width=130  valign=middle>
									<font size=1>Precio
								</td>
								<td valign=top>
									<input type="text" name=precio size=10 maxlength=50 value="<%=Request.Form("precio")%>">
								</td>
							</tr>
							
							<tr>
								<td colspan=2  valign=middle>
									<input type=image src="imagenes/continuar.gif" border=0 name=continuar alt="Calcular Precio">
								</td>
							</tr>
							<tr>
							
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
		<!--- Fin de Contenido --->
	</tr>
</table>
</html>
