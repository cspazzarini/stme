<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include FILE="./Includes/cnx.inc" -->
<%

	
	
	mimes=201407
	mimes2=20140728

	secerro=false

IF SECERRO=FALSE THEN

			
			'PROCESA TODO MENOS CUOTAS


			cadena="SELECT        TOP (100) PERCENT dbo.instrumentos_cuotas.instid, dbo.instrumentos_cuotas.cuotaid, dbo.instrumentos_cuotas.restapagar, "
			CADENA=cADENA & "                          dbo.instrumentos_cuotas.fechacobro, dbo.instrumentos_cuotas.tipoinstrumento"
			CADENA=cADENA & " FROM            dbo.afiliados INNER JOIN"
			CADENA=cADENA & "                          dbo.instrumentos ON dbo.afiliados.numeroafiliado = dbo.instrumentos.numeroafiliado INNER JOIN"
			CADENA=cADENA & "                          dbo.instrumentos_cuotas ON dbo.instrumentos.instid = dbo.instrumentos_cuotas.instid LEFT OUTER JOIN"
			CADENA=cADENA & "                          dbo.instrumentos_pagos ON dbo.instrumentos_cuotas.cuotaid = dbo.instrumentos_pagos.cuotaid"
			CADENA=cADENA & " WHERE        (dbo.instrumentos_cuotas.fechacobro <= " & mimes & ") AND (dbo.instrumentos_cuotas.restapagar > 0) AND (dbo.afiliados.tipoafiliado = 1) AND "
			CADENA=cADENA & "                          (dbo.instrumentos_pagos.cuotaid IS NULL)"
			CADENA=cADENA & " ORDER BY dbo.instrumentos.instid"
			
			
			set rs=conexion.execute(cadena)
			if rs.eof=false then
				do until rs.eof

					cadena="SELECT     cuotaid, MESDESCUENTO FROM         dbo.INSTRUMENTOS_PAGOS WHERE     "
					cadena=cadena & " (MESDESCUENTO = " & mimes & ") AND (cuotaid = " & rs.fields(1) & ") and descargada is null"
					
					set rs2=conexion.execute(cadena)
					if rs2.eof=true then
							x=x+1
				'			if trim(rs.fields(3)) <> trim(rs.fields(4))  then
									cadena="insert into instrumentos_pagos (instid, cuotaid, monto, userid, fecha,mesdescuento, tipoinstrumento) values ("
									cadena=cadena & rs.fields(0) & ", "
									cadena=cadena & rs.fields(1) & ", "
									cadena=cadena & rs.fields(2) & ", "
									cadena=cadena & "0, "
									cadena=cadena & mimes2 & ","
									cadena=cadena & mimes & ","
									cadena=cadena & "'" & rs.fields(4) & "')"
									conexion.execute(cadena)
									
									cadena="update instrumentos_cuotas set restapagar=restapagar-" & rs.fields(2)  & " where cuotaid=" & rs.fields(1)
									conexion.execute(cadena)
									
									cadena="update instrumentos set restapagar=restapagar-" & rs.fields(2)  & " where instid=" & rs.fields(0)
									conexion.execute(cadena)
					end if
					rs.movenext
				loop
			end if
			rs.close
			set rs=nothing

		
			'NUEVO CIERRE DE LOTE SOLO AFILIADOS STME
			CADENA=" SELECT     TOP (100) PERCENT dbo.instrumentos.numeroafiliado, SUM(dbo.instrumentos_pagos.monto) AS Expr1, dbo.instrumentos.tipoinstrumento"
			CADENA=CADENA & " FROM         dbo.instrumentos_pagos INNER JOIN"
			CADENA=CADENA & "                       dbo.instrumentos ON dbo.instrumentos_pagos.instid = dbo.instrumentos.instid INNER JOIN"
			CADENA=CADENA & "                       dbo.afiliados ON dbo.instrumentos.numeroafiliado = dbo.afiliados.numeroafiliado"
			CADENA=CADENA & " WHERE     (dbo.instrumentos_pagos.descargada = 0 OR"
			CADENA=CADENA & "                       dbo.instrumentos_pagos.descargada IS NULL) AND (dbo.instrumentos_pagos.MESDESCUENTO = " & mimes & ") AND (dbo.instrumentos.tipoinstrumento <> 'CREDITO') AND "
			CADENA=CADENA & "                       (dbo.afiliados.tipoafiliado = 1)"
			CADENA=CADENA & " GROUP BY dbo.instrumentos.numeroafiliado, dbo.instrumentos.tipoinstrumento			"

			
			SET RS=CONEXION.EXECUTE(CADENA)
			IF RS.EOF=FALSE THEN
				DO UNTIL RS.EOF
					cadena="INSERT INTO CIERRES_LOTE (legajo, importe, vencimiento, mes_procesado, userid, tipoinstrumento) values ("
					cadena=cadena & "'" & rs.fields(0) & "', "
					cadena=cadena & rs.fields(1) & ", "
					cadena=cadena & mimes & ", "
					cadena=cadena & mimes & ", "
					cadena=cadena & session("userid") & ","
					cadena=cadena & "'" & rs.fields(2) & "')"
					conexion.execute(cadena)
					RS.MOVENEXT
				LOOP
					
			END IF
			RS.CLOSE
			SET RS=NOTHING
			
			cadena="select numeroafiliado, monto, cantidadpagos, instid from instrumentos where creditodescargado=0 and tipoinstrumento='CREDITO'"
			
			'nuevo ciere de lote solo afiliados stme
			cadena="SELECT     dbo.instrumentos.numeroafiliado, dbo.instrumentos.monto, dbo.instrumentos.cantidadpagos, dbo.instrumentos.instid"
			cadena=cadena & " FROM         dbo.instrumentos INNER JOIN"
			cadena=cadena & "                       dbo.afiliados ON dbo.instrumentos.numeroafiliado = dbo.afiliados.numeroafiliado"
			cadena=cadena & " WHERE     (dbo.instrumentos.tipoinstrumento = 'CREDITO') AND (dbo.instrumentos.creditodescargado = 0) AND (dbo.afiliados.tipoafiliado = 1)			"
			set rs=conexion.execute(cadena)
			if rs.eof=false then
				do until rs.eof
					afiliado=rs.fields(0)
					cantidadpagos=rs.fields(2)
					monto=rs.fields(1)/cantidadpagos
					vence=dateadd("m", cantidadpagos-1, date)
					MesVencimiento=year(vence) & right("00" & month(vence),2)
					
					cadena="INSERT INTO CIERRES_LOTE (legajo, importe, vencimiento, mes_procesado, userid, tipoinstrumento) values ("
					cadena=cadena & "'" & afiliado & "', "
					cadena=cadena & monto & ", "
					cadena=cadena & MesVencimiento & ", "
					cadena=cadena & mimes & ", "
					cadena=cadena & session("userid") & ","
					cadena=cadena & "'CREDITO')"
					conexion.execute(cadena)
					
					cadena="update instrumentos set creditodescargado=1 where instid=" & rs.fields(3)
					conexion.execute(cadena)
					rs.movenext
				loop
			end if
			rs.close
			set rs=nothing
		
else
	response.write("Se cerro = false")			
END IF			
%>