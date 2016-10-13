<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*23*") =0 then response.redirect("sinacceso.asp")%>
<%legajo=trim(request.form("legajo"))
if isnumeric(legajo) = false then response.redirect("financiera_3.asp")


Dim conexion
Set Conexion = CreateObject("ADODB.Connection")
conn="Provider=SQLOLEDB;Data source=NOTEBOOK\SQL2012;Database=stmedata;UID=sa;PWD=sql2014*"
conexion.connectionstring = conn
conexion.open

cadena="SELECT     numeroafiliado, nombre, documento FROM         dbo.Afiliados"
cadena=cadena & " WHERE     (numeroafiliado = N'" & legajo & "')"

set rs=conexion.execute(cadena)
if rs.eof=false then
	afiliado=rs.fields(0)
	nombre=rs.fields(1)
	documento=rs.fields(2)
else
	afiliado=0
end if
rs.close
set rs=nothing
conexion.close
set conexion=nothing

if afiliado=0 then response.redirect("financiera_3.asp")

%>
<script>javascript:history.go(+1)</script>
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
.style37 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
.style38 {color: #00769D}
.style40 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; color: #000000; }
.style41 {font-family: Arial, Helvetica, sans-serif; font-weight: bold; color: #FF0000; }
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
<script type="text/javascript">
function format(input){
 var num = input.value.replace(/\,/g,'');
  if(!isNaN(num)){
  } else {
   input.value = input.value.substring(0,input.value.length-1);
  }
}


</script>
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
          <form id="form1" name="form1" method="post" action="financiera_11imprime.asp?tipo=CREDITO">
            <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="30" bgcolor="#384858"><div align="center"><span class="style35">ALTA  DE CREDITO COMPRA COMERCIO</span></div></td>
              </tr>
              <tr>
                <td bgcolor="#374957"><table width="600" border="0" cellspacing="1" cellpadding="0">
                    <tr>
                      <td bgcolor="#FFFFFF"><table width="600" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="16" height="40">&nbsp;</td>
                            <td width="22"><img src="images/spacer12.gif" width="10" height="10" /></td>
                            <td colspan="2"><span class="style37">Alta de nuevo crédito (paso 2)</span></td>
                          </tr>
                          <tr>
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td width="119" class="style33 style38">Nombre:</td>
                            <td width="441"><span class="style40"><%=nombre%></span></td>
                          </tr>
                          <tr>
                            <td height="40">&nbsp;</td>
                            <td><input name="afiliado" type="hidden" id="afiliado" value="<%=afiliado%>" /></td>
                            <td class="style33 style38">Nro. afiliado:</td>
                            <td class="style40"><%=afiliado%>&nbsp;</td>
                          </tr>
                          <tr>
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td class="style33 style38">Documento:</td>
                            <td class="style40"><%=documento%>&nbsp;</td>
                          </tr>
                          <tr>
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td class="style33 style38">Importe de la cuota:</td>
                            <td><span class="style41">$</span>
                              <input name="monto" type="text" class="style41" id="monto" size="10" maxlength="10" onKeyUp="format(this);" /></td>
                          </tr>
                          <tr>
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td class="style33 style38">Cantidad de cuotas:</td>
                            <td><select name="cuotas" class="style40" id="cuotas">
                            <%for t=1 to 24%>
                              <option <%if t=1 then response.write(" selected ")%> value="<%=t%>"><%=t%></option>
                             <%next%>
                            </select>                            </td>
                          </tr>
                          <tr>
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td class="style33 style38">Comercio:</td>
                            <td>
<%
	Set Conexion = CreateObject("ADODB.Connection")
	conn="Provider=SQLOLEDB;Data source=NOTEBOOK\SQL2012;Database=stmedata;UID=sa;PWD=sql2014*"
	conexion.connectionstring = conn
	conexion.open
	cadena="SELECT     TOP (100) PERCENT cid, Comercio FROM         dbo.COMERCIOS where estado=1 ORDER BY Comercio"
	set rs=conexion.execute(cadena)
%>

                            
                            <select name="comercio" class="style40" id="comercio">
                            <%if rs.eof=false then
								do until rs.eof%>
                              <option value="<%=rs.fields(0)%>"><%=right("0000" & rs.fields(0),2) & " - " & rs.fields(1)%></option>
                            <%		rs.movenext
								loop
							end if
							rs.close
							set rs=nothing
							conexion.close
							set conexion=nothing%> 
                            </select>                            </td>
                          </tr>
                          <tr>
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td class="style33 style38">Concepto:</td>
                            <td><input name="concepto" type="text" class="style40" id="concepto" size="75" maxlength="75" /></td>
                          </tr>
                          <tr>
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td class="style33 style38">&nbsp;</td>
                            <td><div align="right">
                                  <input type="submit" name="button" id="button" value="Registrar instrumento" onClick="javascript:return confirm('Registrar instrumento?')"/>
                            &nbsp;&nbsp;</div></td>
                          </tr>
                      </table></td>
                    </tr>
                </table></td>
              </tr>
            </table>
            </form>
          <p align="center" class="style15"><span class="style53"><strong>Acción monitoreada. Usuario responsable: </strong><span class="style55"> <%=session("usuario")%></span></span></p>
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
