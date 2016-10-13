<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include FILE="./Includes/cnx.inc" -->
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*27*") =0 then response.redirect("sinacceso.asp")
Server.ScriptTimeout=1000


	
	mimes=int(year(date) & right("00" & month(date),2))
	mimes2=year(date) & right("00" & month(date),2) & right("00" & day(date),2)


CADENA="SELECT MES FROM INSTRUMENTOS_CIERREMES WHERE MES=" & MIMES 
SET RS=CONEXION.EXECUTE(CADENA)
IF RS.EOF=FALSE THEN
	SECERRO=TRUE
ELSE
	SECERRO=FALSE
	CADENA="INSERT INTO INSTRUMENTOS_CIERREMES VALUES (" & MIMES & ")"
	CONEXION.EXECUTE(CADENA)
END IF
RS.CLOSE
SET RS=NOTHING


IF SECERRO=FALSE THEN


			cadena="SELECT     instid, cuotaid, restapagar, fechacobro, tipoinstrumento FROM         dbo.INSTRUMENTOS_CUOTAS"
			cadena=cadena & " WHERE     (fechacobro <= " & mimes & ") AND (restapagar > 0)		"
			
			'NUEVO CIERRE DE LOTE SOLO PARA AFILIADOS STME
			CADENA="SELECT     TOP (100) PERCENT dbo.instrumentos_cuotas.instid, dbo.instrumentos_cuotas.cuotaid, dbo.instrumentos_cuotas.restapagar, dbo.instrumentos_cuotas.fechacobro, "
			CADENA=cADENA & "                       dbo.instrumentos_cuotas.tipoinstrumento"
			CADENA=cADENA & " FROM         dbo.afiliados INNER JOIN"
			CADENA=cADENA & "                       dbo.instrumentos ON dbo.afiliados.numeroafiliado = dbo.instrumentos.numeroafiliado INNER JOIN"
			CADENA=cADENA & "                       dbo.instrumentos_cuotas ON dbo.instrumentos.instid = dbo.instrumentos_cuotas.instid"
			CADENA=cADENA & " WHERE     (dbo.instrumentos_cuotas.fechacobro <= " & mimes & ") AND (dbo.instrumentos_cuotas.restapagar > 0) AND (dbo.afiliados.tipoafiliado = 1)"
			CADENA=cADENA & " ORDER BY dbo.instrumentos.instid			"
			
			
			set rs=conexion.execute(cadena)
			if rs.eof=false then
				response.write("Process " & rs(0) & "<br>")
				response.Flush()
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
										response.write("Updating field" & rs.fields(0) & "<br>")
										response.Flush()									
					end if
					rs.movenext
				loop
			end if
			rs.close
			set rs=nothing

			cadena="SELECT     TOP (100) PERCENT dbo.INSTRUMENTOS.numeroafiliado, SUM(dbo.INSTRUMENTOS_PAGOS.monto) AS Expr1, dbo.INSTRUMENTOS.tipoinstrumento"
			cadena=cadena & " FROM         dbo.INSTRUMENTOS_PAGOS INNER JOIN"
			cadena=cadena & "                       dbo.INSTRUMENTOS ON dbo.INSTRUMENTOS_PAGOS.instid = dbo.INSTRUMENTOS.instid INNER JOIN"
			cadena=cadena & "                       dbo.Afiliados ON dbo.INSTRUMENTOS.numeroafiliado = dbo.Afiliados.numeroafiliado"
			cadena=cadena & " WHERE     (dbo.INSTRUMENTOS_PAGOS.descargada = 0 or dbo.INSTRUMENTOS_PAGOS.descargada is null) and dbo.INSTRUMENTOS_PAGOS.MESDESCUENTO = " & mimes  
			cadena=cadena & " and dbo.INSTRUMENTOS.tipoinstrumento <> 'CREDITO'"
			cadena=cadena & " GROUP BY dbo.INSTRUMENTOS.numeroafiliado, dbo.INSTRUMENTOS.tipoinstrumento"

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
								response.write("Insert in - Proccesing Stage 2" & "<br>")
								response.Flush()							
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
					response.write("Last Update at SQL DB" & "<br>")
								response.Flush()		
				loop
			end if
			rs.close
			set rs=nothing
			
END IF			

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>STME - Intranet de gestion</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 13px;
	font-weight: bold;
	color: #344356;
}
.style15 {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	color: #344453;
}
a:link {
	text-decoration: none;
	color: #00769D;
}
a:visited {
	text-decoration: none;
	color: #00769D;
}
a:hover {
	text-decoration: none;
	color: #00769D;
}
a:active {
	text-decoration: none;
	color: #00769D;
}
.style32 {color: #FF0000}
.style35 {font-family: Arial, Helvetica, sans-serif; font-weight: bold; color: #FFFFFF; }
.style37 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
.style44 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 10px; }
.style46 {font-size: 14px}
.style47 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 14px; }
.style50 {
	font-size: 18px;
	color: #FF0000;
}
.style36 {font-family: Geneva, Arial, Helvetica, sans-serif;
	color: #FFFFFF;
	font-size: 12px;
}
-->
</style>
<script src="../../../../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>

<body>
<table width="1012" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="6" background="images/izquierda.jpg">&nbsp;</td>
    <td width="1000"><table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
      <tr>
        <td height="8" bgcolor="#2F3863"></td>
      </tr>
      <tr>
        <td height="110"><table width="1000" border="0" cellspacing="0" cellpadding="0">
          <tr bgcolor="#FFFFFF">
            <td height="110" background="images/logonu_fondo.jpg"><img src="images/logonu.jpg" width="276" height="199" /></td>
            <td background="images/logonu_fondo.jpg"><p align="right" class="style36"><br />
              INTRANET DE GESTION DEL SINDICATO DE TRABAJADORES MUNICIPALES DE ENSENADA <br />
            </p></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="33" background="images/fondo_menu_top.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="153"><div align="center"><span class="style1"><a href="menu.asp">INICIO</a></span></div></td>
              <td width="9"><div align="center"><img src="images/separador_menu_top.jpg" width="5" height="33" /></div></td>
              <td width="170"><div align="center"><span class="style1"><a href="contenidos.asp">GESTION  DE CONTENIDOS</a></span></div></td>
              <td width="9"><div align="center"><img src="images/separador_menu_top.jpg" width="5" height="33" /></div></td>
              <td width="171"><div align="center"><span class="style1"><a href="financiera.asp"><span class="style32">GESTION FINANCIERA </span></a></span></div></td>
              <td width="10"><div align="center"><img src="images/separador_menu_top.jpg" width="5" height="33" /></div></td>
              <td width="154"><div align="center"><span class="style1"><a href="coseguro.asp">GESTION COSEGURO </a></span></div></td>
              <td width="9"><div align="center"><img src="images/separador_menu_top.jpg" width="5" height="33" /></div></td>
              <td width="155"><div align="center"><span class="style1"><a href="administracion.asp">ADMINISTRACION</a></span></div></td>
              <td width="10"><div align="center"><img src="images/separador_menu_top.jpg" width="5" height="33" /></div></td>
              <td width="152"><div align="center"><span class="style1"><a href="logout.asp">CERRAR SESION</a></span></div></td>
            </tr>
        </table></td>
      </tr>
      
      <tr>
        <td height="2" bgcolor="#2F3863"></td>
      </tr>
      <tr>
        <td height="400" valign="top" bgcolor="#FFFFFF"><p>&nbsp;</p>
          <p align="center" class="style37">&nbsp;</p>
          <p align="center" class="style37">&nbsp;</p>
          <p align="center" class="style15">DESCUENTO MENSUAL 
            PROCESADO. </p>
          <p align="center" class="style37">PARA VISUALIZAR EL ESTADO UTILICE LA OPCION &quot;HISTORICO DE DESCUENTO CD&quot; DEL MENU DE GESTION FINANCIERA </p>
          <p align="center" class="style15"></p>          <p align="center" class="style15">&nbsp;</p>
          <p align="center" class="style15">&nbsp;</p></td>
      </tr>
      <tr>
        <td height="2" bgcolor="#2F3863"></td>
      </tr>      
    </table></td>
    <td width="6" background="images/derecha.jpg">&nbsp;</td>
  </tr>
</table>
</body>
</html>
