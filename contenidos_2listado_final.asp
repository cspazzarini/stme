<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*1*") =0 then response.redirect("sinacceso.asp")%>
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
.style33 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
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
              <td width="170"><div align="center"><span class="style1"><a href="contenidos.asp"><span class="style32">GESTION  DE CONTENIDOS</span></a></span></div></td>
              <td width="9"><div align="center"><img src="images/separador_menu_top.jpg" width="5" height="33" /></div></td>
              <td width="171"><div align="center"><span class="style1"><a href="financiera.asp">GESTION FINANCIERA </a></span></div></td>
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
<%


orden=request("sort")
if orden="" then orden="nombre"

CADENA="SELECT     TOP (100) PERCENT dbo.Afiliados.numeroafiliado, dbo.Afiliados.nombre, dbo.Afiliados.estadocivil, dbo.tipo_afiliado.tipo, "
CADENA=CADENA & " 						dbo.Afiliados.fechaingreso, "
CADENA=CADENA & "                       dbo.Afiliados.categoria, dbo.Afiliados.id"
CADENA=CADENA & " FROM         dbo.Afiliados INNER JOIN"
CADENA=CADENA & "                       dbo.tipo_afiliado ON dbo.Afiliados.tipoafiliado = dbo.tipo_afiliado.id"
CADENA=CADENA & " where dbo.Afiliados.numeroafiliado <> '24485168' and dbo.Afiliados.numeroafiliado <> '9999999999' and dbo.Afiliados.numeroafiliado <> '999999'  " 

select case(request.form("tipo"))
	case ""
		
	case else
		cadena=cadena & " and tipoafiliado=" & request.form("tipo")
end select


cadena=cadena & " ORDER BY dbo.Afiliados." & orden

SET RS=CONEXION.EXECUTE(CADENA)
IF RS.EOF=FALSE THEN
%>          
          <table width="900" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="30" bgcolor="#384858"><div align="center"><span class="style35">GESTION DE CONTENIDO</span></div></td>
            </tr>
            <tr>
              <td bgcolor="#374957"><table width="900" border="0" cellspacing="1" cellpadding="0">
                <tr>
                  <td bgcolor="#FFFFFF"><table width="900" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="100" height="40" bgcolor="#CCCCCC" class="style33"><div align="center"><strong><a href="contenidos_2listado.asp?sort=numeroafiliado">LEGAJO</a></strong></div></td>
                      <td width="259" bgcolor="#CCCCCC" class="style33"><div align="center"><strong><a href="contenidos_2listado.asp?sort=nombre">AFILIADO</a></strong></div></td>
                      <td width="153" bgcolor="#CCCCCC" class="style33"><div align="center"><a href="contenidos_2listado.asp?sort=estadocivil"><strong>E. CIVIL</strong></a></div></td>
                      <td width="153" bgcolor="#CCCCCC" class="style33"><div align="center"><strong>TIPO AFILIADO</strong></div></td>
                      <td width="108" bgcolor="#CCCCCC" class="style33"><div align="center"><a href="contenidos_2listado.asp?sort=fechaingreso"><strong>FECHA INGRESO</strong></a></div></td>
                      <td width="83" bgcolor="#CCCCCC" class="style33"><div align="center"><strong><a href="contenidos_2listado.asp?sort=categoria">CATEGORIA</a></strong></div></td>
                      <td width="44" bgcolor="#CCCCCC" class="style33"><div align="center"></div></td>
                    </tr>
                    <%do until rs.eof
					if micolor="#e5e5e5" then micolor="#ffffff" else micolor="#e5e5e5"
					%>
                    <tr>
                      <td height="40" bgcolor="<%=micolor%>" class="style33"><div align="center"><%=rs.fields(0)%></div></td>
                      <td bgcolor="<%=micolor%>" class="style33">&nbsp;<%=rs.fields(1)%></td>
                      <td bgcolor="<%=micolor%>" class="style33">&nbsp;<%=rs.fields(2)%></td>
                      <td bgcolor="<%=micolor%>" class="style33">&nbsp;<%=rs.fields(3)%></td>
                      <td bgcolor="<%=micolor%>" class="style33"><div align="center"><%=mid(rs.fields(4),7,2) & "/" & mid(rs.fields(4),5,2) & "/" & left(rs.fields(4),4)%></div></td>
                      <td bgcolor="<%=micolor%>" class="style33"><div align="center"><%=rs.fields(5)%></div></td>
                      <td bgcolor="#FFFFFF"><div align="center"><a href="contenidos_2verafiliado.asp?aid=<%=rs.fields(6)%>"><img src="images/LUPA.jpg" width="30" height="29" border="0" /></a></div></td>
                    </tr>
                    <%rs.movenext
					loop%>
                  </table></td>
                </tr>
              </table></td>
            </tr>
          </table>
          <form id="form1" name="form1" method="post" action="javascript:history.go(-1)">
            <div align="center"><br />
              <input type="submit" name="button" id="button" value="Retornar" />
              </div>
          </form>
          <p align="center">&nbsp;</p>
          <div align="center">
            <p>
              <%ELSE%>
            </p>
            <p>&nbsp;</p>
            <p align="center" class="style15">No hay afiliados para el tipo seleccionado</p>
            <p>&nbsp;</p>
            <p>
              <%END IF
RS.CLOSE
SET RS=NOTHING
CONEXION.CLOSE
SET CONEXION=NOTHING%>
              </p>
          </div>
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
