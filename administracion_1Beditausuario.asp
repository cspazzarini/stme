<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*4*") =0 then response.redirect("sinacceso.asp")
if trim(request("uid"))="" then response.redirect("sinacceso.asp")





Dim conexion
Set Conexion = CreateObject("ADODB.Connection")
conn="Provider=SQLOLEDB;Data source=ARTONEST;Database=stmedata;UID=usuariostme;PWD=Redalert7424"
conexion.connectionstring = conn
conexion.open	

cadena="SELECT     TOP (100) PERCENT nombre, usuario, password, pin, accesos, userid FROM         dbo.Usuarios"
cadena=cadena & " WHERE     (userid = " & request("uid") & ")"
set rs=conexion.execute(cadena)
if rs.eof=false then
	nombre=rs.fields(0)
	sigla=rs.fields(1)
	password=rs.fields(2)
	pin=rs.fields(3)
	accesos=rs.fields(4)
	userid=rs.fields(5)
end if
rs.close
set rs=nothing
conexion.close
set conexion=nothing
	
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
.style33 {	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 12px;
}
.style35 {font-family: Arial, Helvetica, sans-serif; font-weight: bold; color: #FFFFFF; }
.style37 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
.style38 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; color: #FF0000; }
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
              <td width="171"><div align="center"><span class="style1"><a href="financiera.asp">GESTION FINANCIERA </a></span></div></td>
              <td width="10"><div align="center"><img src="images/separador_menu_top.jpg" width="5" height="33" /></div></td>
              <td width="154"><div align="center"><span class="style1"><a href="coseguro.asp">GESTION COSEGURO </a></span></div></td>
              <td width="9"><div align="center"><img src="images/separador_menu_top.jpg" width="5" height="33" /></div></td>
              <td width="155"><div align="center"><span class="style1"><a href="administracion.asp"><span class="style32">ADMINISTRACION</span></a></span></div></td>
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
          <form id="form1" name="form1" method="post" action="administracion_1Beditusuario_save.asp">
            <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="30" bgcolor="#384858"><div align="center"><span class="style35">ALTA DE NUEVO USUARIO</span></div></td>
              </tr>
              <tr>
                <td bgcolor="#374957"><table width="500" border="0" cellspacing="1" cellpadding="0">
                    <tr>
                      <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="24" height="15" bgcolor="#e5e5e5">&nbsp;</td>
                            <td width="152" height="15" bgcolor="#e5e5e5">&nbsp;</td>
                            <td width="322" height="15" bgcolor="#e5e5e5">&nbsp;</td>
                          </tr>
                          <tr>
                            <td height="40" bgcolor="#e5e5e5"><input name="userid" type="hidden" id="userid" value="<%=userid%>" /></td>
                            <td height="40" bgcolor="#e5e5e5" class="style33">Nombre completo:</td>
                            <td bgcolor="#e5e5e5"><input disabled name="F1" type="text" class="style37" id="F1" value="<%=nombre%>" size="40" maxlength="40" /></td>
                          </tr>
                          <tr>
                            <td height="40" bgcolor="#e5e5e5">&nbsp;</td>
                            <td height="40" bgcolor="#e5e5e5" class="style33">Sigla de usuario:</td>
                            <td bgcolor="#e5e5e5"><input disabled name="F2" type="text" class="style37" id="F2" value="<%=sigla%>" size="40" maxlength="40" /></td>
                          </tr>
                          <tr>
                            <td height="40" bgcolor="#e5e5e5">&nbsp;</td>
                            <td height="40" bgcolor="#e5e5e5" class="style33">Contraseña:</td>
                            <td bgcolor="#e5e5e5"><input name="F3" type="password" class="style37" id="F3" value="<%=password%>" size="40" maxlength="40" /></td>
                          </tr>
                          <tr>
                            <td height="40" bgcolor="#e5e5e5">&nbsp;</td>
                            <td height="40" bgcolor="#e5e5e5" class="style33">PIN:</td>
                            <td bgcolor="#e5e5e5"><input name="F4" type="password" class="style37" id="F4" value="<%=pin%>" size="40" maxlength="40" onkeyup="this.value = this.value.replace (/[^0-9]/, ''); "  /></td>
                          </tr>
                          
                          <tr>
                            <td height="15" bgcolor="#e5e5e5">&nbsp;</td>
                            <td height="15" bgcolor="#e5e5e5">&nbsp;</td>
                            <td height="15" bgcolor="#e5e5e5">&nbsp;</td>
                          </tr>
                      </table></td>
                    </tr>
                </table></td>
              </tr>
            </table>
                    <br />
                    <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="30" bgcolor="#384858"><div align="center"><span class="style35">ACCESO A GESTION DE CONTENIDOS</span></div></td>
                      </tr>
                      <tr>
                        <td bgcolor="#374957"><table width="500" border="0" cellspacing="1" cellpadding="0">
                            <tr>
                              <td bgcolor="#FFFFFF"><table width="500" border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <td width="14" height="40">&nbsp;</td>
                                    <td width="19"><img src="images/spacer12.gif" width="10" height="10" /></td>
                                    <td colspan="3"><span class="style37">Emisión de bonos, créditos y adelantos en efectivo</span></td>
                                  </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td width="61">&nbsp;</td>
                                    <td width="27"><input name="checkbox" type="checkbox" id="checkbox" value="11" <%if instr(accesos,"*11*") >0 then %> checked="checked" <%end if%> /></td>
                                    <td width="379"><span class="style33">Dar de alta un nuevo afiliado</span></td>
                                  </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td><input name="checkbox" type="checkbox" id="checkbox15" value="12" <%if instr(accesos,"*12*") >0 then %> checked="checked" <%end if%> /></td>
                                    <td><span class="style33">Administración de afiliados (bajas y modificaciones)</span></td>
                                  </tr>


                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                  </tr>
                              </table></td>
                            </tr>
                        </table></td>
                      </tr>
                    </table>
                    <br />
                    <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="30" bgcolor="#384858"><div align="center"><span class="style35">ACCESO A GESTION FINANCIERA</span></div></td>
                      </tr>
                      <tr>
                        <td bgcolor="#374957"><table width="500" border="0" cellspacing="1" cellpadding="0">
                            <tr>
                              <td bgcolor="#FFFFFF"><table width="500" border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <td width="14" height="40">&nbsp;</td>
                                    <td width="19"><img src="images/spacer12.gif" width="10" height="10" /></td>
                                    <td colspan="3"><span class="style37">Emisión de bonos, créditos y adelantos en efectivo</span></td>
                                  </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td width="61">&nbsp;</td>
                                    <td width="27"><input name="checkbox" type="checkbox" id="checkbox" value="21" <%if instr(accesos,"*21*") >0 then %> checked="checked" <%end if%> /></td>
                                    <td><span class="style33 style32">Emitir Bono</span></td>
                                </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td><input name="checkbox" type="checkbox" id="checkbox" value="22" <%if instr(accesos,"*22*") >0 then %> checked="checked" <%end if%> /></td>
                                    <td><span class="style38">Registrar adelanto  en efectivo</span></td>
                                </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td><input name="checkbox" type="checkbox" id="checkbox" value="23" <%if instr(accesos,"*23*") >0 then %> checked="checked" <%end if%> /></td>
                                    <td width="379"><span class="style38">Emitir  crédito</span></td>
                                </tr>
                                  <tr>
                                    <td height="40">&nbsp;</td>
                                    <td><img src="images/spacer12.gif" width="10" height="10" /></td>
                                    <td colspan="3"><span class="style37">Estado financieros y consolidación de deuda</span></td>
                                  </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td><input name="checkbox" type="checkbox" id="checkbox" value="24" <%if instr(accesos,"*24*") >0 then %> checked="checked" <%end if%> /></td>
                                    <td><span class="style33">Consultar estado de un afiliado</span></td>
                                </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td><input name="checkbox" type="checkbox" id="checkbox" value="25" <%if instr(accesos,"*25*") >0 then %> checked="checked" <%end if%> /></td>
                                    <td><span class="style33">Ingresos mensuales totales (haberesy pago por ventanilla)</span></td>
                                </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td><input name="checkbox" type="checkbox" id="checkbox" value="26" <%if instr(accesos,"*26*") >0 then %> checked="checked" <%end if%>/></td>
                                    <td><span class="style33 style32">Ajuste manual de deuda para proximo cierre de lote </span></td>
                                </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td><input name="checkbox" type="checkbox" id="checkbox" value="27" <%if instr(accesos,"*27*") >0 then %> checked="checked" <%end if%>/></td>
                                    <td><span class="style33 style32">Procesar cierre de lote del mes en curso (Desc. por haberes)</span></td>
                                </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td><input name="checkbox" type="checkbox" id="checkbox" value="28" <%if instr(accesos,"*28*") >0 then %> checked="checked" <%end if%>/></td>
                                    <td><span class="style33">Ver descuento por haberes programado</span></td>
                                </tr>
                                  <tr>
                                    <td height="40">&nbsp;</td>
                                    <td><img src="images/spacer12.gif" width="10" height="10" /></td>
                                    <td colspan="3"><span class="style37">Proveedores</span></td>
                                  </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td><input name="checkbox" type="checkbox" id="checkbox" value="29" <%if instr(accesos,"*29*") >0 then %> checked="checked" <%end if%>/></td>
                                    <td><span class="style33">Estado general de deuda a proveedores</span></td>
                                </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td><input name="checkbox" type="checkbox" id="checkbox" value="210" <%if instr(accesos,"*210*") >0 then %> checked="checked" <%end if%>/></td>
                                    <td><span class="style38">Registrar pago a proveedor</span></td>
                                </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td><input name="checkbox" type="checkbox" id="checkbox" value="2A" <%if instr(accesos,"*2A*") >0 then %> checked="checked" <%end if%>/></td>
                                    <td><span class="style33">Estado de deuda para un mes particular</span></td>
                                </tr>
                                  <tr>
                                    <td height="40">&nbsp;</td>
                                    <td><img src="images/spacer12.gif" width="10" height="10" /></td>
                                    <td colspan="3"><span class="style37">Acciones especiales</span></td>
                                  </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td><input name="checkbox" type="checkbox" id="checkbox" value="2B" <%if instr(accesos,"*2B*") >0 then %> checked="checked" <%end if%>/></td>
                                    <td><span class="style38">Pago de deuda por ventanilla (En efectivo, en sindicato)</span></td>
                                </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td><input name="checkbox" type="checkbox" id="checkbox" value="2C" <%if instr(accesos,"*2C*") >0 then %> checked="checked" <%end if%>/></td>
                                    <td><span class="style38">Modificación de datos de un instrumento</span></td>
                                </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                  </tr>
                              </table></td>
                            </tr>
                        </table></td>
                      </tr>
                    </table>
                    <br />
                    <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="30" bgcolor="#384858"><div align="center"><span class="style35">ACCESO A GESTION COSEGURO</span></div></td>
                      </tr>
                      <tr>
                        <td bgcolor="#374957"><table width="500" border="0" cellspacing="1" cellpadding="0">
                            <tr>
                              <td bgcolor="#FFFFFF"><table width="500" border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <td width="14" height="40">&nbsp;</td>
                                    <td width="19"><img src="images/spacer12.gif" width="10" height="10" /></td>
                                    <td colspan="3"><span class="style37">Emisión de bonos, créditos y adelantos en efectivo</span></td>
                                  </tr>
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td width="61">&nbsp;</td>
                                    <td width="27"><input name="checkbox" type="checkbox" id="checkbox16" value="3" <%if instr(accesos,"*3*") >0 then %> checked="checked" <%end if%>/></td>
                                    <td width="379"><span class="style33">Gestión Coseguro</span></td>
                                  </tr>
                                  
                                  <tr>
                                    <td height="25">&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                  </tr>
                              </table></td>
                            </tr>
                        </table></td>
                      </tr>
                    </table>
                    <p align="center">
                      <input type="submit" name="button" id="button" value=" Modificar Perfil " />
                    </p>
            </form>
          <form id="form2" name="form2" method="post" action="administracion_1Beditusuario_cancelapermisos.asp">
            <div align="center">
              <input name="userid" type="hidden" id="userid" value="<%=userid%>" />
              <input type="submit" name="button2" id="button2" value="Eliminar usuario" />
              <br />
              <br />
            </div>
          </form>          
          <form id="form2" name="form2" method="post" action="javascript:history.go(-1)">
            <div align="center">
              <input type="submit" name="button2" id="button2" value="     Cancelar     " />
              </div>
          </form>          <p align="center" class="style15">&nbsp;</p>
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
