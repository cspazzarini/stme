<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<!-- #include FILE="./Includes/cnx.inc" -->
<%if instr(session("accesos"), "*21*") =0 then response.redirect("sinacceso.asp")


afiliado=request.form("afiliado")
monto=request.form("monto")
cuotas=request.form("cuotas")
comercio=trim(request.form("comercio"))
concepto=trim(request.form("concepto"))
tipo=trim(request.form("tipo"))

if concepto="" then concepto=request("tipo") & " por $ " & monto
if comercio="" then comercio=0

 
if monto <> "" and cuotas <> "" then

	
	ahoraes=year(date) & right("00" & month(date),2) & right("00" & day(date),2)
	mimes=year(date) & right("00" & month(date),2) 
	
	xtotal=monto*cuotas

	cadena="SELECT     mes FROM         dbo.INSTRUMENTOS_CIERREMES WHERE     (mes = " & mimes & ")"
	set rs=conexion.execute(cadena)
	LoteCerrado=false
	if rs.eof=false then
		LoteCerrado=true
	end if
	rs.close
	set rs=nothing


				MES=MONTH(DATE)
				ANO=YEAR(DATE)
				DIA=DAY(DATE)
				
				TMES=MES
				TANO=ANO
				
				if LoteCerrado=true then
					IF TMES=12 THEN
						TANO=TANO+1
						TMES=1
					ELSE
						TMES=TMES+1
					END IF	
				end if

	
	cadena="insert into instrumentos (numeroafiliado, tipoinstrumento, fechasolicitud, monto, restapagar, cantidadpagos, impreso, userid, comercio, concepto, tipoadelanto) values ("
	
	cadena=cadena & "'" & afiliado & "', "
	cadena=cadena & "'" & request("tipo") & "', "
	cadena=cadena & ahoraes & ", "
	cadena=cadena & xtotal & ", "
	cadena=cadena & xtotal & ", "
	cadena=cadena & cuotas & ", "
	cadena=cadena & "0, "
	cadena=cadena & session("userid") & ", "
	cadena=cadena & comercio & ", "
	cadena=cadena & "'" & concepto & "', "
	cadena=cadena & "'" &  tipo & "') "

 
	conexion.execute(cadena)
	
	cadena="select max(instid) from instrumentos where numeroafiliado=" & afiliado & " and monto=" & xtotal & " and fechasolicitud=" & ahoraes
	set rs=conexion.execute(cadena)
	ID=RS.FIELDS(0)
	RS.CLOSE
	SET RS=NOTHING
	
	
	MONTOCUOTA=round(MONTO/CUOTAS,3)
	total=0


	
	FOR T=1 TO CUOTAS

		
		if int(t) < int(cuotas) then
			total=total+montocuota
			cadena="insert into instrumentos_cuotas (instid, numerodecuota, valorcuota, restapagar, fechacobro, tipoinstrumento) values ("
			cadena=cadena & id & ", "
			cadena=cadena & t & ", "
			cadena=cadena & replace(monto,",",".") & ","
			cadena=cadena & replace(monto,",",".") & ","
			CADENA=CADENA & TANO & RIGHT("00" & TMES,2) & ", '" & request("tipo") & "')"
		else
			montocuota=monto - total
			cadena="insert into instrumentos_cuotas (instid, numerodecuota, valorcuota, restapagar, fechacobro, tipoinstrumento) values ("
			cadena=cadena & id & ", "
			cadena=cadena & t & ", "
			cadena=cadena & replace(monto,",",".") & ","
			cadena=cadena & replace(monto,",",".") & ","
			CADENA=CADENA & TANO & RIGHT("00" & TMES,2) & ", '" & request("tipo") & "')"
		end if
		response.write(cadena)
		conexion.execute(cadena)
		IF TMES=12 THEN
			TANO=TANO+1
			TMES=1
		ELSE
			TMES=TMES+1
		END IF		
	NEXT
	
	
	conexion.close
	set conexion=nothing

    if tipo="Cheque" then
		response.redirect("done.asp?destino=financiera")
	else
		response.redirect("financiera_11bono.asp?id=" & id & "&legajo=" & afiliado & "&tipo=" & REQUEST("tipo"))
	end if

end if
%>
<script>javascript:history.go(-1)</script>