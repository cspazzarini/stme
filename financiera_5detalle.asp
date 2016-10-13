<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*28*") =0 then response.redirect("sinacceso.asp")%>
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

dia=int(mid(request("mes"),5,2))
						  ano=left(request("mes"),4)
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
			  

Dim conexion
Set Conexion = CreateObject("ADODB.Connection")
	conn="Provider=SQLOLEDB;Data source=ARTONEST;Database=stmedata;UID=usuariostme;PWD=Redalert7424"
 
conexion.connectionstring = conn
conexion.open			  

cadena=" SELECT     TOP (100) PERCENT dbo.INSTRUMENTOS.numeroafiliado, dbo.Afiliados.nombre, dbo.Afiliados.documento, dbo.Afiliados.cuentabco, "
cadena=cadena & "                       SUM(dbo.INSTRUMENTOS_PAGOS.monto) AS Expr1"
cadena=cadena & " FROM         dbo.INSTRUMENTOS_PAGOS INNER JOIN"
cadena=cadena & "                       dbo.INSTRUMENTOS ON dbo.INSTRUMENTOS_PAGOS.instid = dbo.INSTRUMENTOS.instid INNER JOIN"
cadena=cadena & "                       dbo.Afiliados ON dbo.INSTRUMENTOS.numeroafiliado = dbo.Afiliados.numeroafiliado"
cadena=cadena & " GROUP BY dbo.Afiliados.nombre, dbo.INSTRUMENTOS.numeroafiliado, dbo.INSTRUMENTOS_PAGOS.MESDESCUENTO, dbo.Afiliados.documento, "
cadena=cadena & "                       dbo.Afiliados.cuentabco"
cadena=cadena & " HAVING      (dbo.INSTRUMENTOS_PAGOS.MESDESCUENTO = " & request("mes") & ")"
cadena=cadena & " ORDER BY dbo.INSTRUMENTOS.numeroafiliado"


cadena="SELECT     TOP (100) PERCENT dbo.INSTRUMENTOS.numeroafiliado, dbo.Afiliados.nombre, dbo.Afiliados.documento, dbo.Afiliados.cuentabco, "
cadena=cadena & "                       SUM(dbo.INSTRUMENTOS_PAGOS.monto) AS Expr1"
cadena=cadena & " FROM         dbo.INSTRUMENTOS_PAGOS INNER JOIN"
cadena=cadena & "                       dbo.INSTRUMENTOS ON dbo.INSTRUMENTOS_PAGOS.instid = dbo.INSTRUMENTOS.instid INNER JOIN"
cadena=cadena & "                       dbo.Afiliados ON dbo.INSTRUMENTOS.numeroafiliado = dbo.Afiliados.numeroafiliado"
cadena=cadena & " where dbo.INSTRUMENTOS.numeroafiliado <> '24485168' "
cadena=cadena & " GROUP BY dbo.Afiliados.nombre, dbo.INSTRUMENTOS.numeroafiliado, dbo.INSTRUMENTOS_PAGOS.MESDESCUENTO, dbo.Afiliados.documento, "
cadena=cadena & "                       dbo.Afiliados.cuentabco"
cadena=cadena & " HAVING      (dbo.INSTRUMENTOS_PAGOS.MESDESCUENTO = " & request("mes") & ")"
cadena=cadena & " ORDER BY dbo.INSTRUMENTOS.numeroafiliado"

			  
set rs=conexion.execute(cadena)
			  
			  %></div></td>
            </tr>
            <tr>
              <td bgcolor="#374957">
              <%IF RS.EOF=FALSE THEN%>
              <table width="800" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td width="83" height="40" bgcolor="#E5E5E5" class="style33"><p align="center"><strong>LEGAJO</strong></p></td>
                    <td width="311" bgcolor="#E5E5E5" class="style33"><div align="center"><strong>AFILIADO</strong></div></td>
                    <td width="117" bgcolor="#E5E5E5" class="style33"><div align="center"><strong>DOCUMENTO</strong></div></td>
                    <td width="148" bgcolor="#E5E5E5" class="style33"><div align="center"><strong>CUENTA BAPRO</strong></div></td>
                    <td width="135" bgcolor="#E5E5E5" class="style33"><div align="center"><strong>MONTO</strong></div></td>
                    <td width="135" bgcolor="#E5E5E5" class="style33">&nbsp;</td>
                  </tr>
                  <%DO UNTIL RS.EOF
					if rs.fields(4) <> 0 then
				  %>
                  <tr class="style33">
                    <td height="40" bgcolor="#EAEAEA" ><div align="center"><span ><%=RS.FIELDS(0)%>&nbsp;</span></div></td>
                    <td bgcolor="#EAEAEA" ><strong>&nbsp;&nbsp;<%=RS.FIELDS(1)%>&nbsp;</strong></td>
                    <td bgcolor="#FFFFFF" ><div align="left"><span >&nbsp;&nbsp;&nbsp;<%=RS.FIELDS(2)%>&nbsp;</span></div></td>
                    <td bgcolor="#FFFFFF" ><div align="left"><span >&nbsp;&nbsp;&nbsp;<%=RS.FIELDS(3)%>&nbsp;</span></div></td>
                    <td bgcolor="#F9F1C8" ><div align="right"><strong><%="$" & FORMATNUMBER(RS.FIELDS(4),2)%>&nbsp;</strong></div></td>
                    <td bgcolor="#FFFFFF" ><div align="center"><a href="financiera_5detalle_afiliado.asp?aid=<%=rs.fields(0)%>&mimes=<%=request("mes")%>&nombre=<%=RS.FIELDS(1)%>"><img src="images/LUPA.jpg" width="30" height="29" border="0" /></a></div></td>
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
