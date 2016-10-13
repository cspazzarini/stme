<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*28*") =0 then response.redirect("sinacceso.asp")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>STME - Intranet de gestion</title>
<style type="text/css">
<!-- #include FILE="./Includes/cnx.inc" -->
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
.style33 {	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 12px;
}
.style35 {font-family: Arial, Helvetica, sans-serif; font-weight: bold; color: #FFFFFF; }
.style38 {color: #00769D}
.style40 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; color: #0095C0; }
.style41 {color: #000000}
.style46 {color: #FFFFFF}
.style36 {font-family: Geneva, Arial, Helvetica, sans-serif;
	color: #FFFFFF;
	font-size: 12px;
}
.style48 {color: #666666}
.style49 {
	color: #009900;
	font-weight: bold;
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
          <table width="556" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="800" height="30" bgcolor="#384858"><div align="center"><span class="style35">Exportacion de archivo solicitado por departamento de Haberes (Municipalidad)</span>
                  <%
			  
MES=YEAR(DATE) & RIGHT("00" & MONTH(DATE),2)

cadena="SELECT     TOP (100) PERCENT mes FROM         dbo.INSTRUMENTOS_CIERREMES_salud ORDER BY mes DESC"
			  
set rs=conexion.execute(cadena)
			  
			  
			  %></div></td>
            </tr>
            <tr>
              <td bgcolor="#374957"><table width="556" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td bgcolor="#FFFFFF">
                    <%if rs.eof=false then%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td height="40" bgcolor="#E5E5E5" class="style41 style33"><div align="left"><span class="style41"><strong><span class="style33 style41">&nbsp;</span>MES</strong></span></div></td>
                          <td bgcolor="#E5E5E5"><div align="center"><span class="style41"></span></div></td>
                        </tr>                    
                    <%do until rs.eof

'DEUDA TOTAL DEL MES
CADENA="SELECT     TOP (100) PERCENT SUM(valorcuota) AS Expr1 FROM         dbo.INSTRUMENTOS_CUOTAS GROUP BY fechacobro HAVING      (fechacobro = " & RS.FIELDS(0) & ")"
SET RS2=CONEXION.EXECUTE(CADENA)
IF RS2.EOF=FALSE THEN
	DEUDATOTAL=RS2.FIELDS(0)
ELSE
	DEUDATOTAL=0
END IF
RS2.CLOSE
SET RS2=NOTHING


'PENDIENTE
CADENA="SELECT     TOP (100) PERCENT SUM(restapagar) AS Expr1 FROM         dbo.INSTRUMENTOS_CUOTAS GROUP BY fechacobro HAVING      (fechacobro = " & RS.FIELDS(0) & ")"
SET RS2=CONEXION.EXECUTE(CADENA)
IF RS2.EOF=FALSE THEN
	pendiente=RS2.FIELDS(0)
ELSE
	pendiente=0
END IF
RS2.CLOSE
SET RS2=NOTHING


'PAGO POR VENTANILLA
CADENA="SELECT     TOP (100) PERCENT SUM(monto) AS Expr1 FROM         dbo.INSTRUMENTOS_PAGOS GROUP BY MESDESCUENTO, descargada"
CADENA=CADENA & " HAVING      (MESDESCUENTO = " & RS.FIELDS(0) & ") AND (descargada = 1)"
SET RS2=CONEXION.EXECUTE(CADENA)
IF RS2.EOF=FALSE THEN
	VENTANILLA=RS2.FIELDS(0)
ELSE
	VENTANILLA=0
END IF
RS2.CLOSE
SET RS2=NOTHING



'DESCONTADO
CADENA="SELECT     TOP (100) PERCENT SUM(monto) AS Expr1 FROM         dbo.INSTRUMENTOS_PAGOS GROUP BY MESDESCUENTO HAVING      (MESDESCUENTO = " & RS.FIELDS(0) & ")"

CADENA="SELECT     TOP (100) PERCENT SUM(monto) AS Expr1 FROM         dbo.INSTRUMENTOS_PAGOS GROUP BY MESDESCUENTO, descargada"
CADENA=CADENA & " HAVING      (MESDESCUENTO = " & RS.FIELDS(0) & ") AND (descargada = 0 or descargada is null)"

SET RS2=CONEXION.EXECUTE(CADENA)
IF RS2.EOF=FALSE THEN
	DESCONTADO=RS2.FIELDS(0)
ELSE
	DESCONTADO=0
END IF
RS2.CLOSE
SET RS2=NOTHING
	%>

                        <tr>
                          <td width="148" height="40" bgcolor="#<%=MICOLOR%>" class="style33 style41 style46">&nbsp;
                            <%
						  dia=int(mid(rs.fields(0),5,2))
						  ano=left(rs.fields(0),4)
						  select case dia
						  	case 1
								mimes="Enero"
							case 2
								mimes="Febrero"
							case 3
								mimes="Marzo"
							case 4
								mimes="Abril"
							case 5
								mimes="Mayo"
							case 6
								mimes="Junio"
							case 7
								mimes="Julio"
							case 8
								mimes="Agosto"
							case 9
								mimes="Septiembre"
							case 10
								mimes="Octubre"
							case 11
								mimes="Noviembre"
							case 12
								mimes="Diciembre"
						  end select
						  response.write(mimes & " del " & ano)
						  %>                            </td>
                          <td width="124" bgcolor="#FFFFFF"><div align="center"><a href="financiera_Hdetalle_salud.asp?mes=<%=rs.fields(0)%>" target="_blank" class="style41 style40"><strong><span class="style41"> DETALLE</span></strong></a></div></td>
                        </tr>

                    <%	rs.movenext
					loop
					%>
                    </table>
                      <%else%>
                    <p>&nbsp;</p>
                    <p align="center">  <span class="style40">No se encontró información historica</span></p>
                    <p>&nbsp;</p>
                    <p>
                        <%end if
					rs.close
					set rs=nothing
					conexion.close
					set conexion=nothing%>                    
                    </p></td>
                  </tr>
              </table></td>
            </tr>
          </table>          <p align="center" class="style15">&nbsp;</p>
          </td>
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
