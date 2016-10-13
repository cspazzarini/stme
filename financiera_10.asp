<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include FILE="./Includes/cnx.inc" -->
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*210*") =0 then response.redirect("sinacceso.asp")%>
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
.style47 {color: #000000; font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px;}
.style48 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; font-style: italic; }
.style51 {font-size: 10px}
.style53 {	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 14px;
}
.style55 {color: #FF0000; font-weight: bold; }
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
          <form id="form1" name="form1" method="post" action="financiera_10detalle.asp">
            <table width="400" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="60" bgcolor="#FFFFFF"><div align="center" class="style42"><span class="style15">REGISTRAR PAGO A PROVEEDOR
                  <%

ahoraes=year(date) & right("00" & month(date),2) 

			  
			  
cadena="SELECT     TOP (100) PERCENT Comercio, Direccion, Telefono, cid FROM         dbo.COMERCIOS ORDER BY Comercio"

set rs=conexion.execute(cadena)
			  
			  %>
                    </span><br />
                </div></td>
              </tr>
              <tr>
                <td bgcolor="#374957"><table width="100%" border="0" cellspacing="1" cellpadding="0">
                    <tr>
                      <td height="60" bgcolor="#FFFFFF"><div align="center">
                        <table width="400" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td>&nbsp;</td>
                            <td class="style33">&nbsp;</td>
                          </tr>
                          <tr>
                            <td width="7%">&nbsp;</td>
                            <td width="93%" class="style33"><div align="left">Proveedor</div></td>
                          </tr>
                          <tr>
                            <td>&nbsp;</td>
                            <td><div align="left">
                              <select name="proveedor" id="proveedor">
                              <%
							  if rs.eof=false then
							  do until rs.eof%>
                                <option value="<%=rs.fields(3)%>"><%=rs.fields(0)%></option>
                              <%rs.movenext
							  loop
							  end if
							  rs.close
							  set rs=nothing
							  conexion.close
							  set conexion=nothing
							  %>  
                              </select>
                            </div></td>
                          </tr>
                          <tr>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                          </tr>
                          <tr>
                            <td>&nbsp;</td>
                            <td class="style47"><div align="left">Período</div></td>
                          </tr>
                          <tr>
                            <td>&nbsp;</td>
                            <td><div align="left">
                              <select name="mes" id="mes">
                                <option <%if month(date)=1 then response.write(" selected ")%> value="1">Enero</option>
                                <option <%if month(date)=2 then response.write(" selected ")%> value="2">Febrero</option>
                                <option <%if month(date)=3 then response.write(" selected ")%> value="3">Marzo</option>
                                <option <%if month(date)=4 then response.write(" selected ")%> value="4">Abril</option>
                                <option <%if month(date)=5 then response.write(" selected ")%> value="5">Mayo</option>
                                <option <%if month(date)=6 then response.write(" selected ")%> value="6">Junio</option>
                                <option <%if month(date)=7 then response.write(" selected ")%> value="7">Julio</option>
                                <option <%if month(date)=8 then response.write(" selected ")%> value="8">Agosto</option>
                                <option <%if month(date)=9 then response.write(" selected ")%> value="9">Septiembre</option>
                                <option <%if month(date)=10 then response.write(" selected ")%> value="10">Octubre</option>
                                <option <%if month(date)=11 then response.write(" selected ")%> value="11">Noviembre</option>
                                <option <%if month(date)=12 then response.write(" selected ")%> value="12">Diciembre</option>
                              </select>
                              <select name="ano" id="ano">
                              <%for t=year(date)-1 to year(date)+1%>
                                <option <%if year(date)=t then response.write(" selected ")%> value="<%=t%>"><%=t%></option>
                              <%next%>
                              </select>
                            </div></td>
                          </tr>
                          <tr>
                            <td>&nbsp;</td>
                            <td><div align="left"></div></td>
                          </tr>
                          <tr>
                            <td>&nbsp;</td>
                            <td><div align="left">
                              <input type="submit" name="button" id="button" value="Continuar" />
                            </div></td>
                          </tr>
                          <tr>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                          </tr>
                        </table>
                      </div></td>
                    </tr>
                </table></td>
              </tr>
            </table>
            </form>
          <p align="center"><span class="style53"><strong>Acción monitoreada. Usuario responsable: </strong><span class="style55"> <%=session("usuario")%></span></span></p>
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
