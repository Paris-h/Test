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
	
	
	If Clng(Request("continuar.y")) > 0  Then 
				
		if Request.Form("opcion") > 0 	then
		
			if Request.Form("opcion") =1 then
				if trim(Request.Form("clave")) = "" then
					error=true
					texto="Debe de escribir la clave del producto"
					resultado = false
				else
					resultado = true
					tipo=1
				end if
			end if
			if Request.Form("opcion") =2 then
				if trim(Request.Form("producto")) = "" then
					error=true
					texto="Debe de escribir la descripción del producto"
					resultado = false
				else
					resultado = true
					tipo=2
				end if
			end if
			if Request.Form("opcion") =3 then
				if trim(Request.Form("marca")) = "" then
					error=true
					texto="Debe de seleccionar la marca"
					resultado = false
				else
					resultado = true
					tipo=3
				end if
			end if
			
			tipo =  Request.Form("opcion")
			criterio1 = trim(Request.Form("clave"))
			criterio2 = trim(Request.Form("producto"))
			criterio3 = trim(Request.Form("marca"))
		else
			error=true
			texto="Debe de seleccionar un críterio de busqueda"
		end if
	end if 
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
			<input type=hidden name=venta value=1>
			<table width=665 border=0 cellpadding=0 cellspacing=0 bgcolor=#B1CDDF>
					<tr>
						<td colspan=3 align=right bgcolor=white>
							<a href=alta_catalogo.asp><img src="imagenes/alta_prod.gif" border=0 alt="Devoluciones"></a>
						</td>
					</tr>
					<tr height=25>
						<td colspan=3 bgcolor=#7BA5C3 valign=middle width=550>
							<font face=verdana size=2 color=white><b>&nbsp;&nbsp;Consulta de Cátalogo
						</td>
					</tr>
					<tr valign=top>
						<td valign=top align=center width=5>
							<input type="radio" id=opcion name=opcion value=3 <%if Request.Form("opcion")=3 then%> Checked<%end if%>>
						</td>
						<td valign=middle>
							<font face=verdana size=1>&nbsp;Marca
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
					<tr valign=top>
						<td width=5 valign=top align=center>
							<input type="radio" id=opcion name=opcion value=1 <%if Request.Form("opcion")=1 then%> Checked<%end if%>>
						</td>
						<td width=140 valign=middle>
							<font face=verdana size=1>&nbsp;Clave de producto
						</td>
						<td width=500>
							<input type=text name=clave size=30 maxlength=150 value="<%=request.form("clave")%>">
						</td>	
					</tr>
					<tr valign=top>
						<td valign=top align=center width=5>
							<input type="radio" id=opcion name=opcion value=2 <%if Request.Form("opcion")=2 then%> Checked<%end if%>>
						</td>
						<td valign=middle>
							<font face=verdana size=1>&nbsp;Descripción Producto
						</td>
						<td>
							<input type=text name="producto" size=30 maxlength=150 value="<%=request.form("producto")%>">
						</td>	
					</tr>
					
					<tr valign=top align=right>
						<td colspan=3>
							<br><input type="image" name="continuar" src="imagenes/continuar.gif" border=0 hspace=50></a>
						</td>
					</tr>
				<%if error=true then %>
					<tr valign=top>
						<td valign=top colspan=3>
							<font size=2 color=red face=verdana><b><%=texto%>
						</td>
					</tr>
				<%end if%>
				<%if resultado = true then%>
				<tr valign=top>
					<td colspan=3>
						<IFRAME NAME="grid_catalogo" width="665" height="300" SRC="grid_catalogo.asp?pagina=1&tipo=<%=tipo%>&criterio1=<%=trim(Request.Form("clave"))%>&criterio2=<%=trim(Request.Form("producto"))%>&criterio3=<%=trim(Request.Form("marca"))%>" frameborder=0 scrolling=auto>
						</IFRAME>
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