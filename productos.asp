<%
	if Session("VerPrecioNeto") = "" then 
		VerPrecioNeto = Session("Var").GetValor("APP_ECOMMERCE_VER_PRECIO_NETO", 0, "B", "1")
		Session("VerPrecioNeto") = IIF(VerPrecioNeto,"SI","NO")
	end if 
	
	TextoBusqueda = Request.QueryString("q")
	
	Ct_Cve_Categoria = Request.QueryString("ct")
	Dp_Cve_Departamento = Request.QueryString("dp")
	Mr_Cve_Marca = Request.QueryString("mr")
	Ln_Cve_Linea = Request.QueryString("ln")
	Fm_Cve_Familia = Request.QueryString("fm")
	Sf_Cve_SubFamilia = Request.QueryString("sf")
	
	'QueryFilter = ""
	if Ct_Cve_Categoria <> "" then QueryFilter = QueryFilter & "&ct=" & Ct_Cve_Categoria
	if Dp_Cve_Departamento <> "" then QueryFilter = QueryFilter & "&dp=" & Dp_Cve_Departamento
	if Mr_Cve_Marca <> "" then QueryFilter = QueryFilter & "&mr=" & Mr_Cve_Marca
	if Ln_Cve_Linea <> "" then QueryFilter = QueryFilter & "&ln=" & Ln_Cve_Linea
	if Fm_Cve_Familia <> "" then QueryFilter = QueryFilter & "&fm=" & Fm_Cve_Familia
	if Sf_Cve_SubFamilia <> "" then QueryFilter = QueryFilter & "&sf=" & Sf_Cve_Subfamilia

	QueryFilterSearch = QueryFilter
	if TextoBusqueda <> "" then QueryFilterSearch = QueryFilterSearch & "&q=" & TextoBusqueda

	sFiltros = ""
	sQuery = ""

	if Ct_Cve_Categoria <> "" Then 
		sFiltros = sFiltros & "Producto.Ct_Cve_Categoria = '" & Ct_Cve_Categoria & "' AND "
		sQuery = AddQueryString(sQuery,"ct=" & Ct_Cve_Categoria)
	end if
	
	if Dp_Cve_Departamento <> "" Then 
		sFiltros = sFiltros & "Producto.Dp_Cve_Departamento = '" & Dp_Cve_Departamento & "' AND "
		sQuery = AddQueryString(sQuery,"dp=" & Dp_Cve_Departamento)
	End if
	
	if Mr_Cve_Marca <> "" Then 
		sFiltros = sFiltros & "Producto.Mr_Cve_Marca = '" & Mr_Cve_Marca & "' AND "
		sQuery = AddQueryString(sQuery,"mr=" & Mr_Cve_Marca)
	end if
	
	if Ln_Cve_Linea <> "" Then 
		sFiltros = sFiltros & "Producto.Ln_Cve_Linea = '" & Ln_Cve_Linea & "' AND "
		sQuery = AddQueryString(sQuery,"ln=" & Ln_Cve_Linea)
	end if
	
	if Fm_Cve_Familia <> "" Then 
		sFiltros = sFiltros & "Producto.Fm_Cve_Familia = '" & Fm_Cve_Familia & "' AND "
		sQuery = AddQueryString(sQuery,"fm=" & Fm_Cve_Familia)
	end if
	
	if Sf_Cve_SubFamilia <> "" Then 
		sFiltros = sFiltros & "Producto.Sf_Cve_SubFamilia = '" & Sf_Cve_SubFamilia & "' AND "
		sQuery = AddQueryString(sQuery,"sf=" & Sf_Cve_SubFamilia)
	end if

	ConExistencia = ChangeYN(Request.QueryString("cext"))
	ConPrecio = ChangeYN(Request.QueryString("cpre"))
	ConPromocion = ChangeYN(SoloPromociones)
	
	sQuerySearchForm = sQuery

	if ConExistencia = "SI" then sQuery = AddQueryString(sQuery,"cext=" & ConExistencia)
	if ConPrecio = "SI" then sQuery = AddQueryString(sQuery,"cpre=" & ConPrecio)

	sQueryPage = sQuery
	if TextoBusqueda <> "" Then 
		sFiltros = sFiltros & "("
		sFiltros = sFiltros & "	Producto.Pr_Cve_Producto LIKE '%" & TextoBusqueda & "%' OR "
		sFiltros = sFiltros & "	Producto.Pr_Descripcion LIKE '%" & TextoBusqueda & "%' OR "
		sFiltros = sFiltros & "	Producto.Pr_Descripcion_Corta LIKE '%" & TextoBusqueda & "%' OR "
		sFiltros = sFiltros & "	Producto.Pr_Barras LIKE '%" & TextoBusqueda & "%' OR "
		sFiltros = sFiltros & "	Producto.Pr_Numero_Parte LIKE '%" & TextoBusqueda & "%' "
		sFiltros = sFiltros & ") AND "
		sQueryPage = AddQueryString(sQueryPage,"q=" & TextoBusqueda)
	end if
	
	OrdenarPor = Request.QueryString("order")
	if OrdenarPor = "" then OrdenarPor = "1"
	OrdenarCampo = "Producto.Pr_Cve_Producto"

	select case OrdenarPor
		case "1"
			OrdenarCampo = "Producto.Pr_Cve_Producto"
		case "2"
			OrdenarCampo = "Producto.Pr_Descripcion"
	end select 

	Sucursales = Session("Sucursales")
	Almacenes = Session("Almacenes")

	'Si valida con promociones'
	if ConPromocion = "SI" then 
		sFiltros = sFiltros & " ISNULL(Ds.Ds_Descuento,0) > 0 AND "
	end if 

	if ConExistencia = "SI" then 
		sFiltros = sFiltros & " Producto.Pr_Cve_Producto IN ("
		sFiltros = sFiltros & "		SELECT Existencia.Pr_Cve_Producto FROM Existencia "
		sFiltros = sFiltros & "		WHERE Existencia.Sc_Cve_Sucursal IN ('" & replace(Sucursales,",","','") & "') AND "
		sFiltros = sFiltros & "			Existencia.Al_Cve_Almacen IN ('" & replace(Almacenes,",","','") & "') AND "
		sFiltros = sFiltros & "			Existencia.Ex_Cantidad_Control_1 > 0 "
		sFiltros = sFiltros & "	) AND "
	end if 

	if ConPrecio = "SI" then 
		sFiltros = sFiltros & "	Producto.Pr_Cve_Producto IN ("
		sFiltros = sFiltros & "		SELECT Pr_Cve_Producto FROM Producto "
		sFiltros = sFiltros & "		WHERE dbo.Get_Precio(CAST('" & FormatF(date) & "' as DATE),'" & Session("Moneda") & "','" & Session("Ec").Sc_Cve_Sucursal & "', '" & Session("Ec").Cl_Cve_Cliente & "', Pr_Cve_Producto, '00', '00', Pr_Unidad_Control_1) > 0 AND Es_Cve_Estado = 'AC' AND Pr_Maneja_Talla = 'NO' AND Pr_Maneja_Color = 'NO' "
		sFiltros = sFiltros & "		UNION ALL "
		sFiltros = sFiltros & "		SELECT Pr.Pr_Cve_Producto FROM Producto Pr "
		sFiltros = sFiltros & "			INNER JOIN Producto_Talla_Color Ptc ON Ptc.Pr_Cve_Producto = Pr.Pr_Cve_Producto "
		sFiltros = sFiltros & "		WHERE dbo.Get_Precio(CAST('" & FormatF(date) & "' as DATE),'" & Session("Moneda") & "','" & Session("Ec").Sc_Cve_Sucursal & "', '" & Session("Ec").Cl_Cve_Cliente & "', Pr.Pr_Cve_Producto, Ptc.Tl_Cve_Talla, Ptc.Cl_Cve_Color, Pr.Pr_Unidad_Control_1) > 0 AND Pr.Es_Cve_Estado = 'AC' AND (Pr.Pr_Maneja_Talla = 'SI' OR Pr_Maneja_Color = 'SI') "
		sFiltros = sFiltros & "	) AND "
	end if 

	SQL = ""
	SQL = SQL & "SELECT "
	SQL = SQL & "	ROW_NUMBER() OVER (ORDER BY " & OrdenarCampo & " ASC) as IdTable, "
	SQL = SQL & "	Producto.Pr_Cve_Producto,"
	SQL = SQL & "	Producto.Pr_Descripcion, "
	SQL = SQL & "	Producto.Pr_Clave_Corta, "
	SQL = SQL & "	Producto.Pr_Descripcion_Corta, "

	'Talla color'
	SQL = SQL & "	Producto.Pr_Maneja_Talla, "
	SQL = SQL & "	Producto.Pr_Maneja_Color, "

	SQL = SQL & "	Producto.Pr_Unidad_Venta, "
	SQL = SQL & "	Producto.Pr_Unidad_Control_1, "

	'Marca y departamento'
	SQL = SQL & "	Producto.Mr_Cve_Marca, "
	SQL = SQL & "	Marca.Mr_Descripcion, "
	SQL = SQL & "	Producto.Dp_Cve_Departamento, "
	SQL = SQL & "	Departamento.Dp_Descripcion, "

	'Config ecommerce'
	SQL = SQL & "	Producto_Ecommerce.Pr_Muestra_Precio, "
	SQL = SQL & "	Producto_Ecommerce.Pr_Muestra_Existencia, "
	SQL = SQL & "	Producto_Ecommerce.Pr_Muestra_Serie, "
	SQL = SQL & "	Producto_Ecommerce.Pr_Permite_Captura, "

	'Orden de compra pendiente'
	SQL = SQL & "	IsNull(SUM(Orden_Compra.Oc_Cantidad_Control_1), 0) as Oc_Cantidad_Control_1, "

	'Descuentos'
	SQL = SQL & "	ISNULL(Ds.Ds_Descuento, 0) as Ds_Descuento, "
	SQL = SQL & "	ISNULL(Ds.Ds_Cantidad, 0) as Ds_Cantidad "

	SQL = SQL & "FROM Producto "
	SQL = SQL & "	INNER JOIN Producto_Ecommerce ON Producto.Pr_Cve_Producto = Producto_Ecommerce.Pr_Cve_Producto "
	SQL = SQL & "	INNER JOIN Producto_Talla_Color Ptc ON Ptc.Pr_Cve_Producto = Producto.Pr_Cve_Producto "
	SQL = SQL & "	INNER JOIN Marca ON Producto.Mr_Cve_Marca = Marca.Mr_Cve_Marca "
	SQL = SQL & "	INNER JOIN Departamento ON Departamento.Dp_Cve_Departamento = Producto.Dp_Cve_Departamento "

	'Orden de compra'
	SQL = SQL & "	LEFT JOIN Orden_Compra ON Producto.Pr_Cve_Producto = Orden_Compra.Pr_Cve_Producto AND "
	SQL = SQL & "		Orden_Compra.Oc_Fecha_Entrega >= '" & FormatF(date) & "' "

	sHoraActual = "18991230 " & Hour(now()) & ":" & Minute(now()) & ":" & Second(now()) & ".000"

	SQL = SQL & "	LEFT JOIN ("
	SQL = SQL & "		SELECT ROW_NUMBER() OVER(PARTITION BY Pr_Cve_Producto ORDER BY Pr_Cve_Producto, Ds_Cantidad ASC) as IdPartida, "
	SQL = SQL & "			Pr_Cve_Producto, Ds_Tipo, Ds_Descuento, Ds_Cantidad, Un_Cve_Unidad "
	SQL = SQL & "		FROM Descuento "
	SQL = SQL & "		WHERE Ds_Tipo = 'Pr' AND "
	SQL = SQL & "			(Ds_Sucursal IN ('" & replace(Sucursales,",","','") & "') OR Ds_Sucursal = '%') AND "
	SQL = SQL & "			('" & FormatF(date) & "' BETWEEN Descuento.Ds_Fecha_Inicial AND Descuento.Ds_Fecha_Final) AND "
	SQL = SQL & "			('" & sHoraActual & "' BETWEEN Ds_Hora_Inicial AND Ds_Hora_Final) AND "
	SQL = SQL & "			(PATINDEX ('%' + rtrim(ltrim(str(datepart(weekday,'" & FormatF(date) & "')))) + '%', Ds_Dias_Aplicacion ) > 0) AND "
	SQL = SQL & "			Es_Cve_Estado = 'LI' "
	SQL = SQL & "	) Ds ON Ds.Pr_Cve_Producto = Producto.Pr_Cve_Producto AND Ds.Un_Cve_Unidad = Producto.Pr_Unidad_Venta AND Ds.IdPartida = 1 "

	SQL = SQL & "WHERE "
	SQL = SQL & "	" & sFiltros
	SQL = SQL & "	Producto.Es_Cve_Estado = 'AC' "
	SQL = SQL & "GROUP BY "
	SQL = SQL & "	Producto.Pr_Cve_Producto,"
	SQL = SQL & "	Producto.Pr_Descripcion, "
	SQL = SQL & "	Producto.Pr_Clave_Corta, "
	SQL = SQL & "	Producto.Pr_Descripcion_Corta, "
	SQL = SQL & "	Producto.Pr_Maneja_Talla, "
	SQL = SQL & "	Producto.Pr_Maneja_Color, "
	SQL = SQL & "	Producto.Pr_Unidad_Venta, "
	SQL = SQL & "	Producto.Pr_Unidad_Control_1,  "
	SQL = SQL & "	Producto.Mr_Cve_Marca, "
	SQL = SQL & "	Marca.Mr_Descripcion, "
	SQL = SQL & "	Producto.Dp_Cve_Departamento, "
	SQL = SQL & "	Departamento.Dp_Descripcion, "
	SQL = SQL & "	Producto_Ecommerce.Pr_Muestra_Precio, "
	SQL = SQL & "	Producto_Ecommerce.Pr_Muestra_Existencia, "
	SQL = SQL & "	Producto_Ecommerce.Pr_Muestra_Serie, "
	SQL = SQL & "	Producto_Ecommerce.Pr_Permite_Captura, "

	SQL = SQL & "	ISNULL(Ds.Ds_Descuento, 0), "
	SQL = SQL & "	ISNULL(Ds.Ds_Cantidad, 0) "
	'Response.write SQL
	PageRow = Request.QueryString("pg")
	if PageRow = "" then PageRow = "1"
    if not IsNumeric(PageRow) then PageRow = 1
    
    RowForPage = 10
    FinRow = cdbl(RowForPage) * cdbl(PageRow)
    IniRow = cdbl(FinRow) - (cdbl(RowForPage) - 1)
    
    tSQL = ""
    tSQL = tSQL & "SELECT COUNT(*) as Total "
    tSQL = tSQL & "FROM (" & SQL & ") BusquedaTable "
    
    Set rsTotal = Server.CreateObject("ADODB.Recordset")
    set rsTotal = Session("Var").Conexion.Execute(tSQL)
    
    if err.number <> 0 then 
        Response.Write "Err1: " & err.description
        Response.end 
    end if 

    TotalRegistros = 0
    
    if not rsTotal.EOF then
        TotalRegistros = cdbl(rsTotal("Total"))
    end if 
    
    set rsTotal = nothing 
    
    if cdbl(FinRow) > cdbl(TotalRegistros) then FinRow = TotalRegistros

    NumberPages = 1
	if cdbl(RowForPage) > 0 then 
		NumberPages = MaxInteger(cdbl(TotalRegistros) / cdbl(RowForPage))
		if cdbl(NumberPages) <= 0 then NumberPages = 1
	end if 
    
    tSQL = ""
    tSQL = tSQL & "SELECT * "
    tSQL = tSQL & "FROM (" & SQL & ") ContactoTable " 
    tSQL = tSQL & "WHERE IdTable BETWEEN " & IniRow & " AND " & FinRow & " "
    
	%>
	<script type="text/javascript">
		var oTC = [];
	</script>
	<div class="row">
		<%
		Dim rsListado, Ec, conDescuento, rsPtc
		Set Ec = Session("Ec")
		Set rsListado = Session("Var").Conexion.Execute(tSQL)
		Set oMov = Server.CreateObject("movPRO.Movimiento")
		Set rsPtc = Server.CreateObject("ADODB.Recordset")
		
		Contador = 0
		Do While Not rsListado.EOF

			'Primero marcamos las variables que se usaran'
			Contador = Contador + 1
			
			mostrarPrecio = true
			mostrarExistencia = true
			mostrarSerie = true
			mostrarDescuento = false
			permiteCaptura = true
			
			'MOSTRAR PRECIO
			If Not isnull(rsListado("Pr_Muestra_Precio"))  Then
				If rsListado("Pr_Muestra_Precio") <> "SI" Then
					mostrarPrecio = false
				End If
			Else
				mostrarPrecio = true	'valor por defecto
			End If
			
			'MOSTRAR EXISTENCIA
			If Not isnull(rsListado("Pr_Muestra_Existencia"))  Then
				If rsListado("Pr_Muestra_Existencia") <> "SI" Then
					mostrarExistencia = false
				End If
			Else
				mostrarExistencia = true	'valor por defecto
			End If
		
			'MOSTRAR SERIE
			If Not isnull(rsListado("Pr_Muestra_Serie"))  Then
				If rsListado("Pr_Muestra_Serie") <> "SI" Then
					mostrarSerie = false
				End If
			Else
				mostrarSerie = true	'valor por defecto
			End If
		
			'PERMITE CAPTURA
			If Not isnull(rsListado("Pr_Permite_Captura"))  Then
				If rsListado("Pr_Permite_Captura") <> "SI" Then
					permiteCaptura = false
				End If
			Else
				permiteCaptura = true	'valor por defecto
			End If
			
			If Session("esInvitado") = true AND Application("INVITADO_PERMITIR_COMPRA") = false Then
				mostrarPrecio = Session("Var").GetValor("APP_ECOMMERCE_INVITADO_MOSTRAR_PRECIO",0)
				permiteCaptura = Session("Var").GetValor("APP_ECOMMERCE_INVITADO_PERMITIR_CAPTURA",0)
				mostrarExistencia = Session("Var").GetValor("APP_ECOMMERCE_INVITADO_MOSTRAR_EXISTENCIA",0)
			End If

			If Not mostrarPrecio Then permiteCaptura = false

			'Ahora se procesan los valores del producto, precio, etc
			Pr_Cve_Producto = rsListado("Pr_Cve_Producto")
			Tl_Cve_Talla = "00"
			Cl_Cve_Color = "00"
			Pd_Unidad = rsListado("Pr_Unidad_Venta")

			ManejaTallaColor = False
			ManejaTallaColor = (rsListado("Pr_Maneja_Talla") = "SI" OR rsListado("Pr_Maneja_Color") = "SI")

			PrecioNeto = 0
			PrecioAntiguo = 0
			PrecioAntiguoNeto = 0
			Descuento = 0
			DescuentoNeto = 0
			Precio = 0
			Impuesto = 0
			DescuentoFactor = 0
			DescuentoCantidad = 0
			Existencia = 0
			DefVal = "Talla/Color"
			DefValValue = "00/00"

			'Si se maneja talla color, buscamos las variables de cada caso
			if ManejaTallaColor then 
				sHoraActual = "18991230 " & Hour(now()) & ":" & Minute(now()) & ":" & Second(now()) & ".000"
				
				SQL = ""
				SQL = SQL & "SELECT Ptc.Pr_Cve_Producto, "
				SQL = SQL & "	Ptc.Tl_Cve_Talla, "
				SQL = SQL & "	Talla.Tl_Descripcion, "
				SQL = SQL & "	Ptc.Cl_Cve_Color, "
				SQL = SQL & "	Color.Cl_Descripcion, "
				SQL = SQL & "	Color.Cl_RGB, "
				SQL = SQL & "	ISNULL(Ds.Ds_Cantidad, 0) as Ds_Cantidad, "
				SQL = SQL & "	ISNULL(Ds.Ds_Descuento, 0) as Ds_Descuento "
				SQL = SQL & "FROM Producto Pr "
				SQL = SQL & "	INNER JOIN Producto_Talla_Color Ptc ON Ptc.Pr_Cve_Producto = Pr.Pr_Cve_Producto "
				SQL = SQL & "	INNER JOIN Talla ON Talla.Tl_Cve_Talla = Ptc.Tl_Cve_Talla "
				SQL = SQL & "	INNER JOIN Color ON Color.Cl_Cve_Color = Ptc.Cl_Cve_Color "
				SQL = SQL & "	LEFT JOIN ("
				SQL = SQL & "		SELECT "
				SQL = SQL & "			ROW_NUMBER() OVER(PARTITION BY Pr_Cve_Producto ORDER BY Pr_Cve_Producto, Ds_Talla, Ds_Color, Ds_Cantidad ASC) as IdPartida, "
				SQL = SQL & "			Pr_Cve_Producto, Ds_Talla, Ds_Color, Ds_Tipo, Ds_Descuento, Ds_Cantidad, Un_Cve_Unidad "
				SQL = SQL & "		FROM Descuento "
				SQL = SQL & "		WHERE Ds_Tipo = 'Pr' AND "
				SQL = SQL & "			(Ds_Sucursal IN ('" & replace(Sucursales,",","','") & "') OR Ds_Sucursal = '%') AND "
				SQL = SQL & "			('" & FormatF(date) & "' BETWEEN Descuento.Ds_Fecha_Inicial AND Descuento.Ds_Fecha_Final) AND "
				SQL = SQL & "			('" & sHoraActual & "' BETWEEN Ds_Hora_Inicial AND Ds_Hora_Final) AND "
				SQL = SQL & "			(PATINDEX ('%' + rtrim(ltrim(str(datepart(weekday,'" & FormatF(date) & "')))) + '%', Ds_Dias_Aplicacion ) > 0) AND "
				SQL = SQL & "			Es_Cve_Estado = 'LI' "
				SQL = SQL & "	) Ds ON Ds.Pr_Cve_Producto = Pr.Pr_Cve_Producto AND "
				SQL = SQL & "		Ds.Un_Cve_Unidad = Pr.Pr_Unidad_Venta AND "
				SQL = SQL & "		((Ds.Ds_Talla = Ptc.Tl_Cve_Talla AND Ds.Ds_Color IN ('%',Ptc.Cl_Cve_Color)) OR "
				SQL = SQL & "		(Ds.Ds_Color = Ptc.Cl_Cve_Color AND Ds.Ds_Talla IN ('%',Ptc.Tl_Cve_Talla)) OR "
				SQL = SQL & "		(Ds.Ds_Talla = '%' AND Ds.Ds_Color = '%')) AND "
				SQL = SQL & "		Ds.IdPartida = 1 "
				SQL = SQL & "WHERE Pr.Pr_Cve_Producto = '" & rsListado("Pr_Cve_Producto") & "' AND Pr.Es_Cve_Estado = 'AC' "
				SQL = SQL & "ORDER BY Tl_Cve_Talla ASC"
				'Response.Write SQL
				set rsPtc = Server.CreateObject("ADODB.Recordset")
				set rsPtc = Session("Var").Conexion.Execute(SQL)

				if not rsPtc.EOF then 
					PrecioNeto = Ec.GetPrecioNeto(Pr_Cve_Producto, rsPtc("Tl_Cve_Talla"),rsPtc("Cl_Cve_Color"))
					Precio = Ec.GetPrecio(Pr_Cve_Producto, rsPtc("Tl_Cve_Talla"),rsPtc("Cl_Cve_Color"))
					Impuesto = Ec.GetImpuesto(Pr_Cve_Producto, rsPtc("Tl_Cve_Talla"),rsPtc("Cl_Cve_Color"))
					DescuentoCantidad = cdbl(rsPtc("Ds_Cantidad"))
					DescuentoFactor = cdbl(rsPtc("Ds_Descuento")) / 100
					Existencia = oMov.Get_Existencia(Session("Sucursales"), Session("Almacenes"), cStr(Pr_Cve_Producto), cstr(rsPtc("Tl_Cve_Talla")), cstr(rsPtc("Cl_Cve_Color")), , cStr(rsListado("Pr_Unidad_Control_1")))
					DefVal = ""
					if rsListado("Pr_Maneja_Talla") = "SI" then DefVal = DefVal & rsPtc("Tl_Descripcion")
					if rsListado("Pr_Maneja_Talla") = "SI" and rsListado("Pr_Maneja_Color") = "SI" then DefVal = DefVal & " / "
					if rsListado("Pr_Maneja_Color") = "SI" then DefVal = DefVal & "<span class=""badge rounded-pill bg-primary"" style=""background-color:" & rsPtc("Cl_RGB") & "!important;border: 1px solid #d5d5d5;"">&nbsp;</span> " & rsPtc("Cl_Descripcion")
					DefValValue = rsPtc("Tl_Cve_Talla") & "/" & rsPtc("Cl_Cve_Color")
				end if 
			else
				PrecioNeto = Ec.GetPrecioNeto (Pr_Cve_Producto, Tl_Cve_Talla, Cl_Cve_Color)
				Precio = Ec.GetPrecio(Pr_Cve_Producto, Tl_Cve_Talla, Cl_Cve_Color)	
				Impuesto = Ec.GetImpuesto (Pr_Cve_Producto, Tl_Cve_Talla, Cl_Cve_Color)
				DescuentoCantidad = cdbl(rsListado("Ds_Cantidad"))
				DescuentoFactor = cdbl(rsListado("Ds_Descuento")) / 100
				Existencia = oMov.Get_Existencia(Session("Sucursales"), Session("Almacenes"), cStr(Pr_Cve_Producto), cstr(Tl_Cve_Talla), cstr(Cl_Cve_Color), , cStr(rsListado("Pr_Unidad_Control_1")))
			end if 

			mostrarDescuento = false 

			If DescuentoFactor > 0 Then
				Descuento = Precio * CDbl(DescuentoFactor)
				DescuentoNeto = cdbl(PrecioNeto) * cdbl(DescuentoFactor)
			Else
				Descuento = 0
			End If
			
			If Descuento > 0 Then
				mostrarDescuento = true
				PrecioAntiguo = Precio
				PrecioAntiguo = FormatCurrency(PrecioAntiguo,2)
				PrecioAntiguoNeto = FormatCurrency(PrecioNeto,2)
				Precio = Precio - Descuento
				'PrecioNeto = FormatCurrency((Precio + Impuesto) - Descuento,2)
				PrecioNeto = FormatCurrency(FormatNumber(cdbl(PrecioNeto) - cdbl(DescuentoNeto),0),2)
			End IF

			If Precio = -1 Then
				Precio = "N/A"
				PrecioNeto = "N/A"
			Else
				Precio = FormatCurrency(Precio,2)
			End IF
			
			If PrecioNeto <> "N/A" Then
				Precio = FormatCurrency(Precio,2)
			End If

			If oMov.Error <> "" Then
				Response.Write oMov.Error
				Response.End 
			End If	
			
	
		
			IdProd = "Item" & rsListado("Pr_Cve_Producto")
			
			%>
			<div class="col-12 item-ecommerce mb-3 pb-3" id="<%=IdProd%>_Row">
				<div class="row">
					<div class="col-12 col-md-4 col-lg-3 text-center">
						<input type="hidden" name="<%=IdProd%>_Cve" value="<%=rsListado("Pr_Cve_Producto")%>">
						<a href="/ecommerce/producto/detalle/?id=<%=rsListado("Pr_Cve_Producto") %>">
							<img src="/ecommerce/img/producto.asp?Pr_Cve_Producto=<%=rsListado("Pr_Cve_Producto")%>" class="img-fluid" border="0" width="100px">
						</a>
					</div>
					<div class="col-12 col-md-8 col-lg-9">
						<div class="row">
							<div class="col-12">
								<p class="fs-12 text-truncate m-0" style="max-width: 100%;">
									<a class="cl-2" href="/ecommerce/producto/detalle/?id=<%=rsListado("Pr_Cve_Producto") %>"><%=rsListado("Pr_Descripcion") %></a>
								</p>
								<p class="m-0">
									<span class="fs-8 cl-8">ID: <%=rsListado("Pr_Cve_Producto") %></span> 
									/ <a class="fs-8 cl-2" href="/ecommerce/buscar/?tp=MARCA&mr=<%=rsListado("Mr_Cve_Marca") %>"><%=rsListado("Mr_Descripcion") %></a>
								</p>
							</div>
							<%
							if err.number <> 0 then 
									Response.Write "Err0: " & err.description 
									Response.end 
								end if 
							%>
							<div class="col-6" id="<%=IdProd%>_PCont">
								<% If mostrarPrecio Then %><p></p>
									<p class="m-0">
										<span class="<%=IIF(mostrarDescuento,"cl-11","cl-8")%>" id="<%=IdProd%>_Price"><%=IIF(Session("VerPrecioNeto")="SI",PrecioNeto,Precio) %></span>
										<span class="fs-7"><%=Session("Moneda")%></span>
										<% If mostrarDescuento Then %>
											<span class="fs-8" style="text-decoration:line-through;"><span id="<%=IdProd%>_OldPrice"><%=IIF(Session("VerPrecioNeto")="SI",PrecioAntiguoNeto,PrecioAntiguo) %></span> <span class="fs-7"><%=Session("Moneda")%></span></span>
										<% end if %>
									</p>
									<% 
									if mostrarDescuento and cdbl(DescuentoCantidad) > 0 then %>
										<p class="m-0 fs-7 cl-11">* Descuento apartir de <span id="<%=IdProd%>_DescQuantity"><%=cdbl(DescuentoCantidad) %></span><%=" " & rsListado("Pr_Unidad_Venta") %> </p>
										<%
									End If
								End If
								%>
								<%if Application("E_Mostrar_Existencia") = True AND mostrarExistencia then %>
									<% If Application("Disponible") Then %>
									<p class="m-0 fs-8">
										<span class="fs-8">Existencia: </span><span class="fs-10 <%=IIF(cdbl(Existencia)>0,"cl-12","cl-2")%>"><span id="<%=IdProd%>_Existence"><%=Existencia %></span><%=" " & rsListado("Pr_Unidad_Venta") & " "%></span>
									</p>
									<% end if %>
								<% end if %>
							</div>
							<%
								if err.number <> 0 then 
									Response.Write "Err1: " & err.description 
									Response.end 
								end if 
						%>
							<div class="col-6">
								<div class="row">
									<div class="<%=mpCol("12")%>">
									<% if ManejaTallaColor then 
										%>
										<div class="btn-group d-grid dropdown mb-1">
											<input type="hidden" name="<%=IdProd %>_TCVal" value="" />
											<button class="btn btn-sm btn-block dropdown-toggle border select-dropdown" type="button" data-bs-toggle="dropdown" aria-expanded="false" id="<%=IdProd %>_Text" style="text-align: justify;"><%="Talla/Color" %></button>
											<ul class="dropdown-menu" id="<%=IdProd %>_List">
												<%
												RowNum = 0
												if not rsPtc.EOF then 
													do while not rsPtc.EOF
														RowNum = RowNum + 1

														DefVal = ""
														if rsListado("Pr_Maneja_Talla") = "SI" then DefVal = DefVal & rsPtc("Tl_Descripcion")
														if rsListado("Pr_Maneja_Talla") = "SI" and rsListado("Pr_Maneja_Color") = "SI" then DefVal = DefVal & " / "
														if rsListado("Pr_Maneja_Color") = "SI" then DefVal = DefVal & "<span class=""badge rounded-pill bg-primary border-1"" style=""background-color:" & rsPtc("Cl_RGB") & "!important;border: 1px solid #d5d5d5;"">&nbsp;</span> " & rsPtc("Cl_Descripcion")
														%>
														<li data-name="<%=IdProd %>" data-val="<%=rsPtc("Tl_Cve_Talla") & "/" & rsPtc("Cl_Cve_Color") %>" class="dropdown-item"><%=DefVal %></li>
														<%
														if err.number <> 0 then 
															Response.Write err.description 
															Response.end 
														end if 

														rsPtc.Movenext 
													loop
												end if 
												%>
											</ul>
										</div>
									<% else %>
									<input type="hidden" name="<%=IdProd %>_TCVal" value="00/00" />
									<% end if %>
									</div>
									<div class="<%=mpCol("12")%>">
										<div class="input-group input-group-sm mb-3">
										  <input type="number" class="form-control form-control-sm" id="<%=IdProd%>_Quantity" name="<%=IdProd%>_Quantity" placeholder="0" <% If not permiteCaptura then %>readonly<% end if %>>
										  <div class="input-group-append">
										    <button class="btn bg-2 cl-7 fs-10 btn-sm" type="button" onclick="AddProducto('<%=IdProd%>')" <% If not permiteCaptura then %>disabled<% end if %>>Agregar</button>
										  </div>
										</div>
									</div>
								</div>
							</div>
						</div>
					</div>
				</div>
			</div>
			<%
			
			set rsPtc = nothing 

			rsListado.MoveNext 
		Loop
		%>
	</div>
	<div class="row">
		<nav class="col-12" aria-label="Page navigation">
			<ul class="pagination pagination-sm justify-content-center">
				<%

				UrlBasePagination = "./?" & IIF(sQueryPage<> "",sQueryPage & "&","")

				NumberAct = cdbl(PageRow)
				if cdbl(PageRow) > 1 then 
					if cdbl(NumberAct) > cdbl(NumberPages) then 
						NumberAct = cdbl(NumberPages)
					end if 
				%>
				<li class="page-item"><a class="page-link cl-2" href="<%=UrlBasePagination %>pg=<%=cdbl(NumberAct - 1) %>" tabindex="-1">Anterior</a></li>
				<% 
				else
					%>
					<li class="page-item disabled"><a class="page-link" href="javascript:void(0)" tabindex="-1">Anterior</a></li>
					<%
				end if 

				'Para escritorio
				String_Page = GetPagination(NumberAct,NumberPages,5)
				aNumberPages = Split(String_Page,",")
				
				for each numPage in aNumberPages
					if cdbl(numPage) = cdbl(NumberAct) then 
						%>
						<li class="page-item active"><span class="page-link bg-2 cl-7"><%=cdbl(NumberAct) %></span></li>
						<%
					else
						%>
						<li class="page-item"><a class="page-link cl-2" href="<%=UrlBasePagination %>pg=<%=cdbl(numPage) %>"><%=cdbl(numPage) %></a></li>
						<%
					end if 
				next
				
				if cdbl(NumberAct) < cdbl(NumberPages) then 
				%>
				<li class="page-item"><a class="page-link cl-2" href="<%=UrlBasePagination %>pg=<%=cdbl(NumberAct) + 1 %>">Siguiente</a></li>
				<% 
				else 
				%>
				<li class="page-item disabled"><a class="page-link" href="javascript:void(0)">Siguiente</a></li>
				<%
				end if 
				%>
			</ul>
		</nav>
	</div>
	<script type="text/javascript">
		function AgregarAlCarrito(idProd){
			var cantidad = $('#Cantidad'+idProd).val();

		}

		$.fn.selectpicker.Constructor.BootstrapVersion = '5';
		$('.selectpicker').selectpicker();

		$('.dropdown-menu li').on('click',function(){
			var item_value = $(this).data('val');
			var item_text = $(this).html();
			var item_name = $(this).data('name');

			$('[name="' + item_name + '_TCVal"]').val(item_value);
			$('#' + item_name + '_Text').html(item_text);
			
			LoadPriceData(item_name);
		});

		function LoadPriceData(nameid){
			if($('[name="'+nameid+'_TCVal"]').val() != ''){
				//Si tiene talla y color, se actualizan los precios por ajax
				$.ajax({
					type:'POST',
					url:'/ecommerce/inc/load/tcprice.asp',
					data:{
						Pr_Cve_Producto: $('[name="'+nameid+'_Cve"]').val(),
						TCValue: $('[name="'+nameid+'_TCVal"]').val()
					},success:function(data){
						//Aqui se actualiza
						if(data.substring(0,2)=='OK'){
							$('#'+nameid+'_PCont').html(data.substring(3));
							
							/*var datos = JSON.parse(data.substring(3));
							$('#'+nameid+'_Price').html(datos.price);
							$('#'+nameid+'_OldPrice').html(datos.oldprice);
							$('#'+nameid+'_DescQuantity').html(datos.descquantity);
							$('#'+nameid+'_Existence').html(datos.existence);*/
						}else{
							console.log(data);
						}
					},error:function(a,b,c){
						console.log(c);
					}
				});
			}
		}

		function AddProducto(nameid){
			//Se debe obtener los datos primero para procesar el pedido
			var producto = $('[name="'+nameid+'_Cve"]').val();
			var tallacolor = $('[name="'+nameid+'_TCVal"]').val();
			var cantidad = $('[name="'+nameid+'_Quantity"]').val();
			
			if(tallacolor==''){
				AlertMessage('Advertencia','Seleciona la talla y/o color del producto','warning');
				return;
			}
			
			waitingDialog.show();

			$.ajax({
				type:'POST',
				url:'/ecommerce/pedido/action/addproducto.asp',
				data: { Pr_Cve_Producto: producto, TallaColor: tallacolor, Cantidad: cantidad },
				success:function(data){
					waitingDialog.hide();
					if(data.substring(0,2)=='OK'){
						AlertMessage('Informaci�n','Producto agregado al carrito, haz clic en el menu <b><a href="/ecommerce/pedido/resumen/">Pedido</a></b> para ver los productos agregados.','info');
					}else{
						ShowAlert(data);
					}
				},error:function(a,b,c){
					waitingDialog.hide();
					AlertMessage('Error',c,'danger');
					console.log(c);
				}
			});
		}
	</script>
