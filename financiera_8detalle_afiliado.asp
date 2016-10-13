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
          <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="30" bgcolor="#384858"><div align="center"><span class="style35">Detalle de descuentos para 
              
<%

dia=int(mid(request("mimes"),5,2))
						  ano=left(request("mimes"),4)
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
%>              
              
              
              </span>
                  <%
			  
		  


cadena="SELECT     dbo.COMERCIOS.Comercio, dbo.INSTRUMENTOS.tipoinstrumento, dbo.INSTRUMENTOS_CUOTAS.numerodecuota, "
cadena=cadena & "                       dbo.INSTRUMENTOS_PAGOS.monto, dbo.INSTRUMENTOS_PAGOS.MESDESCUENTO, dbo.INSTRUMENTOS.numeroafiliado, "
cadena=cadena & "                       dbo.INSTRUMENTOS_PAGOS.descargada, dbo.INSTRUMENTOS.instid"
cadena=cadena & " FROM         dbo.INSTRUMENTOS INNER JOIN"
cadena=cadena & "                       dbo.INSTRUMENTOS_CUOTAS ON dbo.INSTRUMENTOS.instid = dbo.INSTRUMENTOS_CUOTAS.instid INNER JOIN"
cadena=cadena & "                       dbo.INSTRUMENTOS_PAGOS ON dbo.INSTRUMENTOS_CUOTAS.cuotaid = dbo.INSTRUMENTOS_PAGOS.cuotaid LEFT OUTER JOIN"
cadena=cadena & "                       dbo.COMERCIOS ON dbo.INSTRUMENTOS.comercio = dbo.COMERCIOS.cid"
cadena=cadena & " WHERE     (dbo.INSTRUMENTOS_PAGOS.MESDESCUENTO = " & request("mimes") & ") AND (dbo.INSTRUMENTOS.numeroafiliado = N'" & request("aid") & "') AND "
cadena=cadena & "                       (dbo.INSTRUMENTOS_PAGOS.descargada = 0 or dbo.INSTRUMENTOS_PAGOS.descargada is null)"

set rs=conexion.execute(cadena)
			  
			  %></div></td>
            </tr>
            <tr>
              <td bgcolor="#374957">
              <%IF RS.EOF=FALSE THEN%>
              <table width="800" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td width="391" height="40" bgcolor="#E5E5E5" class="style33"><p align="center"><strong>COMERCIO</strong></p></td>
                    <td width="131" bgcolor="#E5E5E5" class="style33"><div align="center"><strong>TIPO</strong></div></td>
                    <td width="78" bgcolor="#E5E5E5" class="style33"><div align="center"><strong>CUOTA</strong></div></td>
                    <td width="149" bgcolor="#E5E5E5" class="style33"><div align="center"><strong>MONTO</strong></div></td>
                    <td width="45" bgcolor="#E5E5E5" class="style33">&nbsp;</td>
                  </tr>
                  <%DO UNTIL RS.EOF
					if rs.fields(4) <> 0 then
				  %>
                  <tr class="style33">
                    <td height="40" bgcolor="#EAEAEA" ><div align="left"><span >&nbsp;&nbsp;&nbsp;&nbsp;<%
					IF ISNULL(RS.FIELDS(0))=TRUE THEN RESPONSE.WRITE("S.T.M.E.") ELSE RESPONSE.WRITE(RS.FIELDS(0))%></span></div></td>
                    <td bgcolor="#FFFFFF" ><strong>&nbsp;&nbsp;<%=RS.FIELDS(1)%>&nbsp;</strong></td>
                    <td bgcolor="#FFFFFF" ><div align="left"><span >&nbsp;&nbsp;&nbsp;<%=RS.FIELDS(2)%>&nbsp;</span></div></td>
                    <td bgcolor="#FFFFFF" ><div align="right"><span >$<%=FORMATNUMBER(RS.FIELDS(3),2)%>&nbsp;</span></div></td>
                    <td bgcolor="#FFFFFF" ><div align="center"><a href="financiera_8detalle_afiliado_CREDITO.asp?iid=<%=rs.fields(7)%>&amp;legajo=<%=rs.fields(5)%>"><img src="images/LUPA.jpg" width="30" height="29" border="0" /></a></div></td>
                  </tr>
                  <%end if
				  RS.MOVENEXT
				  LOOP%>
              </table>
              <%END IF
			  RS.CLOSE
			  SET RS=NOTHING
			  CONEXION.CLOSE
			  SET CONEXION=NOTHING%>
              </td>
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
