<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include FILE="./Includes/cnx.inc" -->
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*24*") =0 then response.redirect("sinacceso.asp")
if isnumeric(request.form("numdoc"))=false then response.redirect("financiera_c1.asp")




cadena="SELECT     numeroafiliado AS Expr1 FROM         dbo.INSTRUMENTOS WHERE     (instid = " & request.form("numdoc") & ")"
set rs=conexion.execute(cadena)
if rs.eof=false then
	legajo=rs.fields(0)
end if
rs.close
set rs=nothing

CADENA="SELECT     dbo.Afiliados.numeroafiliado, dbo.Afiliados.nombre, dbo.Afiliados.documento, dbo.Afiliados.domicilio, dbo.Afiliados.telefono, "
CADENA=CADENA & "                       dbo.Afiliados.fechaingreso, dbo.Afiliados.categoria, dbo.Afiliados.situacionlaboral "
CADENA=CADENA & " FROM         dbo.Afiliados "
CADENA=CADENA & " WHERE (dbo.Afiliados.numeroafiliado = N'" & legajo & "')"



set rs=conexion.execute(cadena)

if rs.eof=true then
	rs.close
	set rs=nothing
	conexion.close
	set conexion=nothing
	response.redirect("financiera_c1.asp")
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
.style40 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 14px;
	font-weight: bold;
	color: #000000;
}
.style42 {font-size: 14px}
.style52 {color: #000000}
.style53 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; color: #000000; }
.style54 {color: #FFFFFF}
.style56 {color: #333333}
.style57 {color: #000000; font-weight: bold; }
.style59 {color: #FF0000; font-weight: bold; }
.style60 {
	color: #009900;
	font-weight: bold;
}
.style36 {font-family: Geneva, Arial, Helvetica, sans-serif;
	color: #FFFFFF;
	font-size: 12px;
}
-->
</style>
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
          <table width="900" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="30" bgcolor="#384858"><div align="center"><span class="style35">Estado de cuenta de afiliado</span></div></td>
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
                            <td width="19" height="40">&nbsp;</td>
                            <td width="17"><img src="images/spacer12.gif" width="10" height="10" /></td>
                            <td colspan="4"><span class="style37"><span class="style42">INFORMACION 
                              DEL AFILIADO</span> 
                              <%
						  


						  
						  
						  %>
                              </span></td>
                          </tr>
                          <tr> 
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td width="153" class="style33 style38">Número de 
                              Afiliado:</td>
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
                            <td width="393" class="style33"><%=rs.fields(3)%></td>
                            <td width="63" class="style33"><span class="style33 style38">DNI:</span></td>
                            <td width="255" class="style33"><%=rs.fields(2)%></td>
                          </tr>
                          <tr> 
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td class="style33 style38">&nbsp;</td>
                            <td colspan="3" class="style33">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td height="40"><span class="style33 style38"> 
                              <%RS.CLOSE
						  SET RS=NOTHING%>
                              </span></td>
                            <td><img src="images/spacer12.gif" width="10" height="10" /></td>
                            <td colspan="4" class="style33 style38"><span class="style40"> 
                              DETALLE DE TRANSACCION. 
                              <%
						  
cadena="SELECT     dbo.INSTRUMENTOS.tipoinstrumento, dbo.INSTRUMENTOS.fechasolicitud, dbo.INSTRUMENTOS.monto, dbo.INSTRUMENTOS.restapagar, "
cadena=cadena & "                       dbo.INSTRUMENTOS.cantidadpagos, dbo.COMERCIOS.Comercio, dbo.INSTRUMENTOS.concepto"
cadena=cadena & " FROM         dbo.INSTRUMENTOS LEFT OUTER JOIN"
cadena=cadena & "                       dbo.COMERCIOS ON dbo.INSTRUMENTOS.comercio = dbo.COMERCIOS.cid"
cadena=cadena & " WHERE     (dbo.INSTRUMENTOS.numeroafiliado = N'" & legajo & "') AND (dbo.INSTRUMENTOS.instid = " & request.form("numdoc") & ")"

set rs=conexion.execute(cadena) 

if rs.eof=false then
	tipoinstrumento=rs.fields(0)
	fecha=mid(rs.fields(1),7,2) & "/" & mid(rs.fields(1),5,2) & "/" & left(rs.fields(1),4)
	monto=rs.fields(2)
	deuda=rs.fields(3)
	cuotas=rs.fields(4)
	comercio=rs.fields(5)
	concepto=rs.fields(6)
end if
rs.close
set rs=nothing
						  
						  
						  %>
                              </span></td>
                          </tr>
                          <tr> 
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td class="style33 style38">Tipo de transacción:</td>
                            <td class="style33 style38"><span class="style33 style52"><%=tipoinstrumento%></span></td>
                            <td class="style33 style38">Fecha:</td>
                            <td class="style33 style38"><span class="style53"><%=fecha%></span></td>
                          </tr>
                          <tr> 
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td class="style33 style38">Monto total:</td>
                            <td class="style33 style38"><span class="style33 style52">$<%=FORMATNUMBER(monto,2)%></span></td>
                            <td class="style33 style38">Deuda:</td>
                            <td class="style33 style38"><span class="style53">$<%=FORMATNUMBER(deuda)%></span></td>
                          </tr>
                          <tr> 
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td class="style33 style38">Comercio:</td>
                            <td colspan="3" class="style33 style38"><span class="style53"><%=comercio%></span></td>
                          </tr>
                          <tr> 
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td class="style33 style38">Concepto:</td>
                            <td colspan="3" class="style33 style38"><span class="style53"><%=concepto%></span></td>
                          </tr>
                          <tr> 
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td class="style33 style38">Instrumento:</td>
                            <td colspan="3"><font color="#FF0000" size="5" face="Geneva, Arial, Helvetica, sans-serif">
                              <%=request.form("numdoc")%>
                              </font></td>
                          </tr>
                          <tr> 
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td class="style33 style38">&nbsp;</td>
                            <td colspan="3" class="style33 style38">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td height="40">&nbsp;</td>
                            <td><img src="images/spacer12.gif" width="10" height="10" /></td>
                            <td colspan="4" class="style33 style38"><span class="style40">CUOTAS 
                              Y PAGOS 
                              <%
						  

CADENA="SELECT     TOP (100) PERCENT numerodecuota, valorcuota, fechacobro, cuotaid, INSTID  FROM         dbo.INSTRUMENTOS_CUOTAS"
CADENA=CADENA & " WHERE     (instid = " & request.form("numdoc") & ") ORDER BY numerodecuota asc"

pagos=0

CADENA="SELECT     TOP (100) PERCENT fechacobro, numerodecuota, valorcuota, restapagar, cuotaid"
CADENA=CADENA & " FROM         dbo.INSTRUMENTOS_CUOTAS WHERE     (instid = " & request.form("numdoc") & ") ORDER BY fechacobro"

SET RS=CONEXION.EXECUTE(CADENA)
						  
						  %>
                              </span></td>
                          </tr>
                          <tr> 
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td colspan="4" class="style33 style38"> <%IF RS.EOF=FALSE THEN%> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td width="96%" bgcolor="#384858"><table width="100%" border="0" cellspacing="1" cellpadding="0">
                                      <tr> 
                                        <td width="8%" height="40"><div align="center"><span class="style54">CUOTA</span></div></td>
                                        <td width="17%"><div align="center"><span class="style54">MONTO 
                                            CUOTA</span></div></td>
                                        <td width="17%"><div align="center"><span class="style54">FECHA 
                                            DE PAGO</span></div></td>
                                        <td width="19%" bgcolor="#FF0000"><div align="center"> 
                                            <blockquote> 
                                              <p><span class="style54">DEUDA</span></p>
                                            </blockquote>
                                          </div></td>
                                        <td width="30%"><div align="center"><span class="style54">DETALLE 
                                            DE PAGO(S)</span></div></td>
                                      </tr>
                                      <%do until rs.eof

							
								%>
                                      <tr> 
                                        <td height="40" bgcolor="#E5E5E5"><div align="center"><%=rs.fields(1)%></div></td>
                                        <td bgcolor="#FFFFFF"><div align="right" class="style52"><strong>
                                            <%
								  TETA=TETA+rs.fields(2)
								  
								  %>
                                            </strong>$<%=formatnumber(rs.fields(2),2)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div></td>
                                        <td bgcolor="#FFFFFF"><div align="center"><%=MID(rs.fields(0),5,2) & "/" & LEFT(rs.fields(0),4)%></div></td>
                                        <td bgcolor="#ffffff"><div align="center">
                                            <%CULO=CULO+RS.FIELDS(3)%>
                                            $
                                            <%response.write(formatnumber(RS.FIELDS(3),2))%>
                                          </div></td>
                                        <td bgcolor="#FFFFFF"><div align="left"> 
                                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                              <%
								  
								  CADENA="SELECT     TOP (100) PERCENT fecha, monto, descargada FROM         dbo.INSTRUMENTOS_PAGOS WHERE     (cuotaid = " & RS.FIELDS(4) & ")"

								  SET RS2=CONEXION.EXECUTE(CADENA)
								  IF RS2.EOF=FALSE THEN
								  		pagos=1
								  		DO UNTIL RS2.EOF
								  %>
                                              <%if rs2.fields(1) <> 0 then%>
                                              <tr> 
                                                <td width="69%"><span class="style56">&nbsp;&nbsp;&nbsp;Fecha 
                                                  dto: &nbsp;<%=MID(RS2.FIELDS(0),7,2) & "/" & MID(RS2.FIELDS(0),5,2) & "/" & LEFT(RS2.FIELDS(0),4)%> 
                                                  <%if rs2.fields(2)=1 then%>
                                                  <span class="style59">(M)</span> 
                                                  <%else%>
                                                  <span class="style60">(D)</span> 
                                                  <%end if%>
                                                  </span></td>
                                                <td width="31%"><div align="right" class="style56">
                                                    <%MOCO=MOCO+RS2.FIELDS(1)%>
                                                    $<%=FORMATNUMBER(RS2.FIELDS(1),2)%>&nbsp;&nbsp;</div></td>
                                              </tr>
                                              <%end if
								  		RS2.MOVENEXT
								  		LOOP
								  END IF
								  RS2.CLOSE
								  SET RS2=NOTHING%>
                                            </table>
                                          </div></td>
                                      </tr>
                                      <%rs.movenext
								loop%>
                                      <tr> 
                                        <td height="40" bgcolor="#E5E5E5">&nbsp;</td>
                                        <td bgcolor="#FFFFFF"><div align="right" class="style57">$<%=formatnumber(teta,2)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div></td>
                                        <td bgcolor="#FFFFFF"><div align="center" class="style57">---</div></td>
                                        <td bgcolor="#FFFFFF"><div align="center" class="style57">$<%=formatnumber(culo,2)%></div></td>
                                        <td bgcolor="#FFFFFF"><div align="center" class="style57"> 
                                            <div align="right">$<%=formatnumber(MOCO,2)%><span class="style56">&nbsp;</span></div>
                                          </div></td>
                                      </tr>
                                    </table></td>
                                  <td width="4%">&nbsp;</td>
                                </tr>
                              </table>
                                </p>
                              </div>
                              <p>&nbsp; </p>
                              <p> 
                                <%ELSE%>
                              </p>
                              <p>No se encontro información de pagos efectuados para este instrumento 
                              </p>
                              <p> 
                                <%END IF
							rs.close
							set rs=nothing%>
                            
                              </p>
<%if pagos=0 then%> <form name="form1" id="form1" method="post" action="financiera_c12delete.asp">
                                <div align="center"> 
                                  <input name="NUMDOC" type="hidden" id="NUMDOC" value="<%=request.form("numdoc")%>" />
                                  <input type="submit" name="Submit" value="Eliminar Instrumento"  onclick="javascript:return confirm('Eliminar instrumento?')"/>
                                </div>
                              </form>
                              <div align="center"> 
                                <p> 
                                  <%else%>
                                </p>
                                <p><span class="style40">Este instrumento no puede 
                                  ser eliminado</span> </p>
                                <p> 
                                  <%end if%>
                              </td>
                          </tr>
                          <tr> 
                            <td height="40">&nbsp;</td>
                            <td>&nbsp;</td>
                            <td colspan="4" class="style33 style38">&nbsp;</td>
                          </tr>
                        </table></td>
                  </tr>
              </table></td>
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
