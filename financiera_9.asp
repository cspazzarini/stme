﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*29*") =0 then response.redirect("sinacceso.asp")%>
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
.style37 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
.style41 {
	font-size: 12px
}
.style42 {
	color: #000000
}
.style44 {font-family: Arial, Helvetica, sans-serif; font-size: 14px;}
.style45 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-style: italic;
	font-weight: bold;
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
        <td height="8" colspan="3" bgcolor="#2F3863"></td>
      </tr>
      <tr>
        <td height="110" colspan="3"><table width="1000" border="0" cellspacing="0" cellpadding="0">
          <tr bgcolor="#FFFFFF">
            <td height="110" background="images/logonu_fondo.jpg"><img src="images/logonu.jpg" width="276" height="199" /></td>
            <td background="images/logonu_fondo.jpg"><p align="right" class="style36"><br />
              INTRANET DE GESTION DEL SINDICATO DE TRABAJADORES MUNICIPALES DE ENSENADA <br />
            </p></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="33" colspan="3" background="images/fondo_menu_top.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
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
        <td height="2" colspan="3" bgcolor="#2F3863"></td>
      </tr>
      <tr>
        <td width="261" height="400" valign="top" bgcolor="#e5e5e5"><p class="style41"><br />
            <span class="style15">GESTION DE PROVEEDORES</span></p>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="17%"><div align="center"><img src="images/spacer12.gif" width="10" height="10" /></div></td>
              <td width="83%"><a href="financiera_9.asp" class="style45">Estado general de deuda</a></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td><div align="center"><img src="images/spacer12.gif" width="10" height="10" /></div></td>
              <td><a href="financiera_91.asp" class="style33">Obligaciones a futuro</a></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td><div align="center"><img src="images/spacer12.gif" width="10" height="10" /></div></td>
              <td><a href="financiera_92.asp" class="style33">Registro de pagos</a></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td><div align="center"><img src="images/spacer12.gif" width="10" height="10" /></div></td>
              <td><a href="financiera_93.asp" class="style33">Historial de un proveedor</a></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td><div align="center"><img src="images/spacer12.gif" width="10" height="10" /></div></td>
              <td><a href="financiera_94.asp" class="style33">Historial de pagos</a></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
          </table>
          <p class="style41">&nbsp;</p>
          <p>&nbsp;</p>
          <p align="center" class="style15">&nbsp;</p>          </td>
        <td width="13" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="726" valign="top" bgcolor="#FFFFFF"><br />
          <table width="700" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="60" bgcolor="#FFFFFF"><div align="center" class="style42"><span class="style15">ESTADO GENERAL DE DEUDA A PROVEEDORES
                      <%

ahoraes=year(date) & right("00" & month(date),2) 

Dim conexion
Set Conexion = CreateObject("ADODB.Connection")
conn="Provider=SQLOLEDB;Data source=ARTONEST;Database=stmedata;UID=usuariostme;PWD=Redalert7424"
conexion.connectionstring = conn
conexion.open			  
			  
cadena="SELECT     dbo.COMERCIOS.Comercio, dbo.COMERCIOS.Direccion, dbo.COMERCIOS.Telefono, SUM(dbo.INSTRUMENTOS_CUOTAS.valorcuota) AS Expr1, dbo.COMERCIOS.cid"
cadena=cadena & " FROM         dbo.INSTRUMENTOS INNER JOIN"
cadena=cadena & "                       dbo.INSTRUMENTOS_CUOTAS ON dbo.INSTRUMENTOS.instid = dbo.INSTRUMENTOS_CUOTAS.instid INNER JOIN"
cadena=cadena & "                       dbo.COMERCIOS ON dbo.INSTRUMENTOS.comercio = dbo.COMERCIOS.cid"
cadena=cadena & " WHERE     (dbo.INSTRUMENTOS_CUOTAS.fechacobro >= " & ahoraes & ") and dbo.COMERCIOS.cid <> 101"
cadena=cadena & " GROUP BY dbo.COMERCIOS.Comercio, dbo.COMERCIOS.Direccion, dbo.COMERCIOS.Telefono, dbo.COMERCIOS.cid			  "


set rs=conexion.execute(cadena)
			  
			  %>
                    </span><br />
                      <span class="style44">INCLUYE EL MES EN CURSO</span></div></td>
            </tr>
            <tr>
              <td bgcolor="#374957"><table width="700" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td width="271" height="40" bgcolor="#CCCCCC" class="style37"><div align="center">PROVEEDOR</div></td>
                    <td width="267" bgcolor="#CCCCCC" class="style37"><div align="center">DIRECCION </div></td>
                    <td width="106" bgcolor="#CCCCCC" class="style37"><div align="center">TELEFONO</div></td>
                    <td width="106" bgcolor="#CCCCCC" class="style37"><div align="center">MONTO</div></td>
                    <td width="44" bgcolor="#CCCCCC" class="style37">&nbsp;</td>
                  </tr>
                  <%IF RS.EOF=FALSE THEN
				  	do until rs.eof%>
                  <tr>
                    <td height="40" bgcolor="#FFFFFF" class="style33">&nbsp;<%=RS.FIELDS(0)%></td>
                    <td bgcolor="#FFFFFF" class="style33">&nbsp;<%=RS.FIELDS(1)%></td>
                    <td bgcolor="#FFFFFF" class="style33">&nbsp;<%=RS.FIELDS(2)%></td>
                    <td bgcolor="#FFFFFF" class="style33"><div align="right">$<%tcl=tcl+RS.FIELDS(3)%><%=FORMATNUMBER(RS.FIELDS(3),2)%>&nbsp;&nbsp;</div></td>
                    <td bgcolor="#FFFFFF" class="style33"><div align="center"><a href="financiera_9detalleproveedor.asp?idp=<%=rs.fields(4)%>&amp;PN=<%=RS.FIELDS(0)%>"><img src="images/LUPA.jpg" width="30" height="29" border="0" /></a></div></td>
                  </tr>

                  <%	rs.movenext
				  loop
				  END IF
				  RS.CLOSE
				  SET RS=NOTHING
				  CONEXION.CLOSE
				  SET CONEXION=NOTHING%>
                  <tr>
                    <td height="40" colspan="3" bgcolor="#FFFFFF" class="style33"><div align="center"><strong>TOTAL GENERAL</strong></div></td>
                    <td bgcolor="#FFFFFF" class="style33"><div align="right">$
                          
                      <%=FORMATNUMBER(TCL,2)%>&nbsp;&nbsp;</div></td>
                    <td bgcolor="#FFFFFF" class="style33">&nbsp;</td>
                  </tr>                  
              </table></td>
            </tr>
          </table></td>
      </tr>
      <tr>
        <td height="2" colspan="3" bgcolor="#2F3863"></td>
      </tr>      
    </table></td>
    <td width="6" background="images/derecha.jpg">&nbsp;</td>
  </tr>
</table>
</body>
</html>
