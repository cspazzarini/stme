<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include FILE="./Includes/cnx.inc" -->
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*26*") =0 then response.redirect("sinacceso.asp")


if request.form("legajo")="" then response.redirect("financiera_6.asp")




CADENA="SELECT     dbo.Afiliados.numeroafiliado, dbo.Afiliados.nombre, dbo.Afiliados.documento, dbo.Afiliados.domicilio, dbo.Afiliados.telefono, "
CADENA=CADENA & "                       dbo.Afiliados.fechaingreso, dbo.Afiliados.categoria, dbo.Afiliados.situacionlaboral, "
CADENA=CADENA & " 					  COUNT(dbo.afiliados_parientes.numeroafiliado) AS Expr1, "
CADENA=CADENA & "                       dbo.Afiliados.Estado"
CADENA=CADENA & " FROM         dbo.Afiliados LEFT OUTER JOIN"
CADENA=CADENA & "                       dbo.afiliados_parientes ON dbo.Afiliados.id = dbo.afiliados_parientes.id"
CADENA=CADENA & " GROUP BY dbo.Afiliados.numeroafiliado, dbo.Afiliados.nombre, dbo.Afiliados.documento, dbo.Afiliados.domicilio, dbo.Afiliados.telefono, "
CADENA=CADENA & "                       dbo.Afiliados.fechaingreso, dbo.Afiliados.categoria, dbo.Afiliados.situacionlaboral, dbo.Afiliados.Estado"
CADENA=CADENA & " HAVING      (dbo.Afiliados.numeroafiliado = N'" & request.form("legajo") & "')"




set rs=conexion.execute(cadena)

if rs.eof=true then
	rs.close
	set rs=nothing
	conexion.close
	set conexion=nothing
	response.redirect("financiera_6.asp")
end if

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
.style38 {color: #00769D}
.style40 {	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 14px;
	font-weight: bold;
	color: #000000;
}
.style42 {font-size: 14px}
.style43 {color: #F9760C}
.style44 {color: #FFFFFF}
.style46 {font-size: 12px; }
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
        

<%
MIMES=YEAR(DATE) & RIGHT("00" & MONTH(DATE),2)
CADENA="SELECT MES FROM INSTRUMENTOS_CIERREMES WHERE MES=" & MIMES 
SET RS2=CONEXION.EXECUTE(CADENA)
IF RS2.EOF=FALSE THEN
	SECERRO=TRUE
ELSE
	SECERRO=FALSE
END IF
RS2.CLOSE
SET RS2=NOTHING


IF SECERRO=FALSE THEN
  %>      
        
        
          <table width="900" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="30" bgcolor="#384858"><div align="center"><span class="style35">Registro de pago</span></div></td>
            </tr>
            <tr>
              <td bgcolor="#374957"><table width="900" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td bgcolor="#FFFFFF"><table width="900" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td colspan="4">&nbsp;</td>
                        </tr>
                        <tr>
                          <td width="13" height="40">&nbsp;</td>
                          <td width="18"><img src="images/spacer12.gif" width="10" height="10" /></td>
                          <td colspan="4"><span class="style37"><span class="style42">INFORMACION DEL AFILIADO</span>
                                <%
						  


						  
						  
						  %>
                          </span></td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td width="154" class="style33 style38">Número de Afiliado:</td>
                          <td colspan="3" class="style33"><%=rs.fields(0)%></td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td class="style33 style38">Nombre y apellido:</td>
                          <td class="style33"><%=rs.fields(1)%></td>
                          <td class="style33"><span class="style33 style38">Teléfono:</span></td>
                          <td class="style33"><%=rs.fields(4)%></td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td class="style33 style38">Domicilio:</td>
                          <td width="455" class="style33"><%=rs.fields(3)%></td>
                          <td width="73" class="style33"><span class="style33 style38">DNI:</span></td>
                          <td width="187" class="style33"><%=rs.fields(2)%></td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td class="style33 style38">&nbsp;</td>
                          <td colspan="3" class="style33">&nbsp;</td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td class="style33 style38">Fecha de ingreso:</td>
                          <td class="style33"><%=mid(rs.fields(5),7,2) & "/" & mid(rs.fields(5),5,2) & "/" & left(rs.fields(5),4)%></td>
                          <td class="style33"><span class="style33 style38">Antigüedad:</span></td>
                          <td class="style33"><%
						  antig=year(date) - int(left(rs.fields(5),4))
						  
						  response.write(antig)
						  
						  %>
                            &nbsp;años</td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td class="style33 style38">Situación laboral:</td>
                          <td class="style33"><%=rs.fields(7)%></td>
                          <td class="style33"><span class="style33 style38">Categoría:</span></td>
                          <td class="style33"><%=rs.fields(6)%></td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td class="style33 style38">Familiares a cargo:</td>
                          <td colspan="3" class="style33"><%=rs.fields(8)%></td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td class="style33 style38"><%RS.CLOSE
						  SET RS=NOTHING%>
                            &nbsp;</td>
                          <td colspan="3">&nbsp;</td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td><img src="images/spacer12.gif" width="10" height="10" /></td>
                          <td colspan="4" class="style33 style38"><span class="style40"> INGRESO DE PAGOS MANUALES </span><span class="style33 style43">(Se muestran pagos atrasados y del mes en curso. LOS CREDITOS NO PUEDEN SER MODIFICADOS)</span><span class="style40">
                          <%

limite=year(date) & right("00" & month(date),2)
						  
cadena="SELECT     TOP (100) PERCENT dbo.INSTRUMENTOS.tipoinstrumento, dbo.INSTRUMENTOS_CUOTAS.valorcuota, dbo.INSTRUMENTOS_CUOTAS.restapagar, "
cadena=cadena & "                       dbo.INSTRUMENTOS_CUOTAS.fechacobro, dbo.COMERCIOS.Comercio, dbo.INSTRUMENTOS_CUOTAS.numerodecuota,"
cadena=cadena & "                       dbo.INSTRUMENTOS_CUOTAS.instid, dbo.INSTRUMENTOS_CUOTAS.cuotaid"
cadena=cadena & "  FROM         dbo.INSTRUMENTOS INNER JOIN"
cadena=cadena & "                       dbo.INSTRUMENTOS_CUOTAS ON dbo.INSTRUMENTOS.instid = dbo.INSTRUMENTOS_CUOTAS.instid LEFT OUTER JOIN"
cadena=cadena & "                       dbo.COMERCIOS ON dbo.INSTRUMENTOS.comercio = dbo.COMERCIOS.cid"
cadena=cadena & " WHERE     (dbo.INSTRUMENTOS.numeroafiliado = N'" & request.form("legajo") & "') AND (dbo.INSTRUMENTOS_CUOTAS.fechacobro <= " & limite & ") AND "
cadena=cadena & "                       (dbo.INSTRUMENTOS_CUOTAS.restapagar > 0) and dbo.instrumentos.tipoinstrumento <> 'CREDITO'"
cadena=cadena & " ORDER BY dbo.INSTRUMENTOS_CUOTAS.fechacobro, dbo.INSTRUMENTOS_CUOTAS.numerodecuota"



SET RS=CONEXION.EXECUTE(CADENA)

%>
                                                    </span></td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td colspan="4" class="style33 style38"><p>
                            <%if rs.eof=false then%>
                            
                            
                                </p>
                            <form id="form1" name="form1" method="post" action="financiera_61save.asp">
                              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                  <td bgcolor="#384858"><table width="100%" border="0" cellspacing="1" cellpadding="0">
                                      <tr>
                                        <td width="11%" height="40"><div align="center" class="style46"><span class="style44">INSTRUMENTO</span></div></td>
                                        <td width="6%"><div align="center" class="style46"><span class="style44">CUOTA</span></div></td>
                                        <td width="11%"><div align="center" class="style46"><span class="style44">MONTO ORIGINAL</span></div></td>
                                        <td width="10%"><div align="center" class="style46"><span class="style44">MONTO ADEUDADO</span></div></td>
                                        <td width="11%"><div align="center" class="style46"><span class="style44">FECHA<br />
                                        </span></div></td>
                                        <td width="36%"><div align="center" class="style46"><span class="style44">COMERCIO</span></div></td>
                                        <td width="15%"><div align="center" class="style46"><span class="style44">ESTE PAGO</span></div></td>
                                      </tr>
                                      <%DO UNTIL RS.EOF
								  	IF INT(RS.FIELDS(3)) < INT(YEAR(DATE) & RIGHT("00" & MONTH(DATE),2)) THEN
										XCOLOR="FF0000"
									ELSE
										XCOLOR="FFFFFF"
									END IF
									
									X=X+1
								  %>
                                      <tr>
                                        <td height="40" bgcolor="#FFFFFF" class="style33">&nbsp;&nbsp;<%=RS.FIELDS(0)%></td>
                                        <td bgcolor="#FFFFFF" class="style33"><div align="center"><%=RS.FIELDS(5)%></div></td>
                                        <td bgcolor="#FFFFFF" class="style33"><div align="right">$<%=FORMATNUMBER(RS.FIELDS(1),2)%>&nbsp;&nbsp;&nbsp;&nbsp;</div></td>
                                        <td bgcolor="#FFFFFF" class="style33"><div align="right" class="style32">$<%=FORMATNUMBER(RS.FIELDS(2),2)%>&nbsp;&nbsp;&nbsp;&nbsp;</div></td>
                                        <td bgcolor="#<%=XCOLOR%>" class="style33"><div align="center"><%=RIGHT(RS.FIELDS(3),2) & " / " & LEFT(RS.FIELDS(3),4)%></div></td>
                                        <td bgcolor="#FFFFFF" class="style33">&nbsp;&nbsp;<%=RS.FIELDS(4)%></td>
                                        <td bgcolor="#FFFFFF" class="style33"><div align="center"> 
                                          <input name="MAXIMO<%=X%>" type="hidden" id="MAXIMO<%=X%>" value="<%=RS.FIELDS(2)%>" />
                                          <input name="INSTID<%=X%>" type="hidden" id="INSTID<%=X%>" value="<%=RS.FIELDS(6)%>" />
                                          <input name="CUOTAID<%=X%>" type="hidden" id="CUOTAID<%=X%>" value="<%=RS.FIELDS(7)%>" />
                                          $
                                          <input name="MONTO<%=X%>" type="text" id="MONTO<%=X%>" size="10" maxlength="10" onkeyup="format(this);"/>
                                        </div></td>
                                      </tr>
                                      <%RS.MOVENEXT
								  LOOP%>
                                  </table></td>
                                </tr>
                              </table>
                              <p align="center"><span class="style53"><strong>Acción monitoreada. Usuario responsable: </strong><span class="style55"> <%=session("usuario")%></span></span></p>
                              <p align="center">
                                <input name="MAXLINES" type="hidden" id="MAXLINES" value="<%=X%>" />
                                <input type="submit" name="button" id="button" value="Registrar Pago(s)" onclick="javascript:return confirm('Registrar pagos?')"/>
                              </p>
                              <p>&nbsp;</p>
                            </form>
                            <p>
                              <%else%>
                            </p>
                            <p>No se encontró deuda actual ni pendiente para este afiliado                              </p>
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
              </table></td>
            </tr>
          </table>
          <p>
            <%ELSE%>
            </p>
          <p>&nbsp;</p>
          <p align="center" class="style15 style32">EL MES EN CURSO HA SIDO CERRADO. </p>
          <p align="center" class="style37">ESTA OPCION SE ENCUENTRA DESHABILITADA HASTA EL MES ENTRANTE </p>
          <p>&nbsp;</p>
          <p>
              <%END IF%>          
          </p>
          <p align="center" class="style15">&nbsp;</p>
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
