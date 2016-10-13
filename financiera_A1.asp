<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*2A*") =0 then response.redirect("sinacceso.asp")%>
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
.style37 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
.style41 {
	font-size: 12px
}
.style42 {
	color: #000000
}
.style45 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 12px;
	color: #000000;
	font-weight: bold;
	font-style: italic;
}
.style46 {font-size: 10px}
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
        <td width="726" height="400" valign="top" bgcolor="#FFFFFF"><br />
          <table width="700" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="60" bgcolor="#FFFFFF"><div align="center" class="style42"><span class="style15">DETALLE DE OBLIGACIONES DE PAGO PARA 
                      <%

						  dia=int(mid(REQUEST("MIMES"),5,2))
						  ano=left(REQUEST("MIMES"),4)
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
						  response.write(UCASE(mimes & " del " & ano))

					
					%><%

ahoraes=REQUEST("MIMES")
	  
			  
cadena="SELECT     TOP (100) PERCENT dbo.COMERCIOS.Comercio, dbo.COMERCIOS.Direccion, dbo.COMERCIOS.Telefono, "
cadena=cadena & "                       SUM(dbo.INSTRUMENTOS_CUOTAS.valorcuota) AS Expr1, dbo.COMERCIOS.cid"
cadena=cadena & " FROM         dbo.INSTRUMENTOS_CUOTAS INNER JOIN"
cadena=cadena & "                       dbo.INSTRUMENTOS ON dbo.INSTRUMENTOS_CUOTAS.instid = dbo.INSTRUMENTOS.instid INNER JOIN"
cadena=cadena & "                       dbo.COMERCIOS ON dbo.INSTRUMENTOS.comercio = dbo.COMERCIOS.cid"
cadena=cadena & " WHERE     (dbo.INSTRUMENTOS_CUOTAS.fechacobro = " & ahoraes & ")"
cadena=cadena & " GROUP BY dbo.COMERCIOS.Comercio, dbo.COMERCIOS.Direccion, dbo.COMERCIOS.Telefono, dbo.COMERCIOS.cid"
cadena=cadena & " ORDER BY dbo.COMERCIOS.Comercio"


set rs=conexion.execute(cadena)
			  
			  %>
                    </span><br />
              </div></td>
            </tr>
            <tr>
              <td bgcolor="#374957"><table width="700" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td width="57%" height="40" bgcolor="#CCCCCC" class="style37"><div align="center">COMERCIO</div></td>
                    <td width="27%" bgcolor="#CCCCCC" class="style37"><div align="center">MONTO</div></td>
                    <td width="16%" bgcolor="#CCCCCC" class="style37"><div align="center">DETALLE</div></td>
                    </tr>
                  <%IF RS.EOF=FALSE THEN
				  	do until rs.eof%>
                  <tr>
                    <td height="40" bgcolor="#FFFFFF" class="style33"><strong>&nbsp;&nbsp;
                        <%
					=RS.FIELDS(0)
					%>
                        <br />
&nbsp;                        &nbsp;</strong><span class="style46">(<%=RS.FIELDS(1)%>, Tel: <%=rs.fields(2)%>)</span></td>
                    <td bgcolor="#FFFFFF" class="style33"><div align="right">$<%=FORMATNUMBER(RS.FIELDS(3),2)%><%SUMTOT=SUMTOT+RS.FIELDS(3)%>&nbsp;&nbsp;</div></td>
                    <td bgcolor="#FFFFFF" class="style33"><div align="center"><a href="financiera_A2.asp?mimes=<%=request("mimes")%>&amp;cid=<%=rs.fields(4)%>"><img src="images/LUPA.jpg" width="30" height="29" border="0" /></a></div></td>
                    </tr>

                  <%	rs.movenext
				  loop
				  END IF
				  RS.CLOSE
				  SET RS=NOTHING
				  CONEXION.CLOSE
				  SET CONEXION=NOTHING%>
                  <tr>
                    <td height="40" bgcolor="#E5E5E5" class="style33"><div align="center"><strong>TOTAL GENERAL</strong></div></td>
                    <td bgcolor="#E5E5E5" class="style33"><div align="right"><strong>$<%=FORMATNUMBER(SUMTOT,2)%>&nbsp;&nbsp;</strong></div></td>
                    <td bgcolor="#E5E5E5" class="style33">&nbsp;</td>
                  </tr>                  
              </table></td>
            </tr>
          </table>
          <p>&nbsp;</p>
          <p>&nbsp;</p></td>
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
