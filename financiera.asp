<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*2*") =0 then response.redirect("sinacceso.asp")%>
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
.style39 {color: #006699}
.style40 {color: #009900}
.style42 {
	color: #006699;
	font-family: Arial, Helvetica, sans-serif;
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
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="14%">&nbsp;</td>
              <td width="14%">&nbsp;</td>
              <td width="14%">&nbsp;</td>
              <td width="14%">&nbsp;</td>
              <td width="14%">&nbsp;</td>
              <td width="14%">&nbsp;</td>
              <td width="14%">&nbsp;</td>
            </tr>
            <tr>
              <td colspan="7" bgcolor="#CCCCCC" class="style15"><div align="center">ACCIONES  MONITOREADAS</div>                
              <div align="center"></div></td>
              </tr>
            <tr>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td><div align="center"><a href="confirmapin.asp?destino=financiera_1"><img src="images/bono.png" width="70" height="70" border="0" /></a></div></td>
              <td><div align="center"><a href="confirmapin.asp?destino=financiera_2"><img src="images/1_efectivo.png" width="70" height="70" border="0" /></a></div></td>
              <td><div align="center"><a href="confirmapin.asp?destino=financiera_3"><img src="images/1_credit.png" width="70" height="70" border="0" /></a></div></td>
              <td><div align="center"><a href="confirmapin.asp?destino=financiera_B"><img src="images/2_ajustemanualdeuda.png" width="70" height="70" border="0" /></a></div></td>
              <td><div align="center"><a href="confirmapin.asp?destino=financiera_6"><img src="images/2_ajustemanualdeuda.png" width="70" height="70" border="0" /></a></div></td>
              <td><div align="center"><a href="confirmapin.asp?destino=financiera_C"><img src="images/2_vercierrelote.png" width="70" height="70" border="0" /></a></div></td>
              <td><div align="center"><a href="confirmapin.asp?destino=financiera_10"><img src="images/3_registrarpagoproveedor.png" width="70" height="70" border="0" /></a></div></td>
            </tr>
            <tr>
              <td class="style33"><div align="center" class="style32">Emitir Bono</div></td>
              <td class="style33"><div align="center" class="style32">Adelanto en Efectivo</div></td>
              <td class="style33"><div align="center" class="style32">Emitir Credito</div></td>
              <td class="style33"><div align="center"><span class="style38">Registrar pago por ventanilla</span></div></td>
              <td class="style33"><div align="center"><span class="style32">Pago por ventanilla antes de cierre de lote</span></div></td>
              <td class="style33"><div align="center"><span class="style32">Eliminar<br />
              Instrumento</span></div></td>
              <td class="style33"><div align="center"><span class="style32">Registrar pago a proveedor</span></div></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td colspan="7" bgcolor="#CCCCCC"><div align="center" class="style15">ACCIONES DE CONSULTA</div></td>
              </tr>
            <tr>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td><div align="center"><a href="financiera_4.asp"><img src="images/2_estado_afiliado.png" width="70" height="70" border="0" /></a></div></td>
              <td><div align="center"><a href="financiera_detalle.asp"><img src="images/2_ingresosmensualestotales.png" width="70" height="70" border="0" /></a></div></td>
              <td><div align="center"><a href="financiera_8.asp"><img src="images/2_vercierrelote.png" width="70" height="70" border="0" /></a></div></td>
              <td><div align="center"><a href="financiera_resumen.asp"><img src="images/bono.png" width="70" height="70" border="0" /></a></div></td>
              <td><div align="center"><a href="financiera_8_salud.asp"><img src="images/2_vercierrelote.png" width="70" height="70" border="0" /></a></div></td>
              <td><div align="center"><a href="financiera_A.asp"><img src="images/3_deudaproveedores.png" width="70" height="70" border="0" /></a><a href="financiera_H_salud.asp"></a></div></td>
              <td><div align="center"><a href="financiera_9.asp"><img src="images/3_deudaproveedores.png" width="70" height="70" border="0" /></a></div></td>
            </tr>
            <tr>
              <td class="style33"><div align="center"><span class="style39">Estado de deuda de un afiliado</span></div></td>
              <td class="style33"><div align="center" class="style39">Total de creditos emitidos en el mes</div></td>
              <td class="style33"><div align="center" class="style32"><span class="style39">Detalle de Cierre de Lote <span class="style32">MUNICIPAL</span></span></div></td>
              <td class="style33"><div align="center"><span class="style39">Resumen de documentos emitidos</span></div></td>
              <td class="style33"><div align="center" class="style39"><span class="style32"><span class="style39">Detalle de Cierre de Lote</span> <span class="style40">SALUD</span></span></div></td>
              <td class="style33"><div align="center" class="style32"><span class="style39">Estado de deuda de un mes particular<span class="style40"></span></span></div></td>
              <td class="style33"><div align="center"><span class="style39">Estado de deuda a proveedores</span></div></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td><div align="center"><a href="financiera_proveedores.asp" target="_blank"><img src="images/1_listarinstrumentos.png" width="70" height="70" border="0" /></a></div></td>
              <td>&nbsp;</td>
              <td><div align="center"></div></td>
              <td>&nbsp;</td>
              <td><div align="center"></div></td>
              <td><div align="center"><a href="financiera_H.asp"><img src="images/2_exportardatos_alboludo.png" width="70" height="70" border="0" /></a></div></td>
              <td><div align="center"><a href="financiera_H_salud.asp"><img src="images/2_exportardatos_alboludo.png" width="70" height="70" border="0" /></a></div></td>
            </tr>
            <tr>
              <td class="style33"><div align="center"><span class="style39"><a href="financiera_proveedores.asp" target="_blank">Listado Imprimible de Proveedores</a></span></div></td>
              <td>&nbsp;</td>
              <td><div align="center"></div></td>
              <td>&nbsp;</td>
              <td class="style33"><div align="center" class="style32"></div></td>
              <td><div align="center"><span class="style42">Exportar Listado de Cierre de Lote <span class="style32">MUNICIPAL</span></span></div></td>
              <td class="style33"><div align="center" class="style32"><span class="style39">Exportar Listado de Cierre de Lote <span class="style40">SALUD</span></span></div></td>
            </tr>
          </table>
          <p>&nbsp;</p>
          <p>&nbsp;</p>
          <%if ftc=1 then%>
          <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="30" bgcolor="#384858"><div align="center"><span class="style35">GESTION FINANCIERA</span></div></td>
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
                          <td width="68">&nbsp;</td>
                          <td width="20"><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td><a href="confirmapin.asp?destino=financiera_1"><span class="style33 style32">Emitir Bono</span></a></td>
                        </tr>
                        <tr>
                          <td height="25">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>&nbsp;</td>
                          <td><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td><a href="confirmapin.asp?destino=financiera_2"><span class="style38">Registrar adelanto  en efectivo</span></a></td>
                        </tr>
                        <tr>
                          <td height="25">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>&nbsp;</td>
                          <td><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td><a href="confirmapin.asp?destino=financiera_3"><span class="style38">Emitir crédito</span></a></td>
                        </tr>
                        <tr>
                          <td height="25">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>x</td>
                          <td><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td width="379"><a href="financiera_detalle.asp"><span class="style33">Listar instrumentos emitidos en el mes</span></a></td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td><img src="images/spacer12.gif" width="10" height="10" /></td>
                          <td colspan="3"><span class="style37">Estado financieros y consolidación de deuda</span></td>
                        </tr>
                        <tr>
                          <td height="25">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>x</td>
                          <td><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td><a href="financiera_4.asp"><span class="style33">Consultar estado de un afiliado</span></a></td>
                        </tr>
                        <tr>
                          <td height="25">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>&nbsp;</td>
                          <td><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td><a href="financiera_5.asp"><span class="style33">Ingresos mensuales totales (haberes y pago en ventanilla)</span></a></td>
                        </tr>
                        <tr>
                          <td height="25">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>&nbsp;</td>
                          <td><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td><a href="confirmapin.asp?destino=financiera_6"><span class="style33 style32">Ajuste manual de deuda para proximo cierre de lote</span></a></td>
                        </tr>
                        <tr>
                          <td height="25">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>&nbsp;</td>
                          <td><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td><a href="confirmapin.asp?destino=financiera_7"><span class="style33 style32">Procesar cierre de lote del mes en curso (Desc. por haberes)</span></a></td>
                        </tr>
                        <tr>
                          <td height="25">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>x</td>
                          <td><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td><a href="financiera_8.asp"><span class="style33">Ver listado de cierre de lote para un mes en particular</span></a></td>
                        </tr>
                        <tr>
                          <td height="25">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>x</td>
                          <td><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td><a href="financiera_H.asp"><span class="style33">Exportar listado de descuento por haberes (Archivo TXT)</span></a></td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td><img src="images/spacer12.gif" width="10" height="10" /></td>
                          <td colspan="3"><span class="style37">Proveedores</span></td>
                        </tr>
                        <tr>
                          <td height="25">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>x</td>
                          <td><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td><a href="financiera_9.asp"><span class="style33">Estado general de deuda a proveedores</span></a></td>
                        </tr>
                        <tr>
                          <td height="25">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>x</td>
                          <td><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td><a href="financiera_4.asp"><span class="style38">Registrar pago a proveedor</span></a></td>
                        </tr>
                        <tr>
                          <td height="25">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>x</td>
                          <td><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td><a href="financiera_A.asp"><span class="style33">Estado de deuda para un mes particular</span></a></td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td><img src="images/spacer12.gif" width="10" height="10" /></td>
                          <td colspan="3"><span class="style37">Acciones especiales</span></td>
                        </tr>
                        <tr>
                          <td height="25">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>x</td>
                          <td><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td><a href="confirmapin.asp?destino=financiera_B"><span class="style38">Pago de deuda por ventanilla (En efectivo, en Sindicato)</span></a></td>
                        </tr>
                        <tr>
                          <td height="25">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>&nbsp;</td>
                          <td><img src="images/spacer09.gif" width="12" height="9" /></td>
                          <td><a href="confirmapin.asp?destino=financiera_C"><span class="style38">Modificación de datos de un instrumento</span></a></td>
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
          </table> <%end if%>         <p align="center" class="style15">&nbsp;</p>
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
