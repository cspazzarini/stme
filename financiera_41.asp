<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include FILE="./Includes/cnx.inc" -->
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*24*") =0 then response.redirect("sinacceso.asp")

LEGAJO=REQUEST.FORM("LEGAJO")
IF LEGAJO="" THEN
	LEGAJO=REQUEST("AID")
END IF

if LEGAJO="" then response.redirect("financiera_4.asp")




CADENA="SELECT     dbo.Afiliados.numeroafiliado, dbo.Afiliados.nombre, dbo.Afiliados.documento, dbo.Afiliados.domicilio, dbo.Afiliados.telefono, "
CADENA=CADENA & "                       dbo.Afiliados.fechaingreso, dbo.Afiliados.categoria, dbo.Afiliados.situacionlaboral, "
CADENA=CADENA & " 					  COUNT(dbo.afiliados_parientes.numeroafiliado) AS Expr1, "
CADENA=CADENA & "                       dbo.Afiliados.Estado, Informado,tipoafiliado"
CADENA=CADENA & " FROM         dbo.Afiliados LEFT OUTER JOIN"
CADENA=CADENA & "                       dbo.afiliados_parientes ON dbo.Afiliados.id = dbo.afiliados_parientes.id"
CADENA=CADENA & " GROUP BY dbo.Afiliados.numeroafiliado, dbo.Afiliados.nombre, dbo.Afiliados.documento, dbo.Afiliados.domicilio, dbo.Afiliados.telefono, "
CADENA=CADENA & "                       dbo.Afiliados.fechaingreso, dbo.Afiliados.categoria, dbo.Afiliados.situacionlaboral, dbo.Afiliados.Estado, Informado,tipoafiliado"
CADENA=CADENA & " HAVING      (dbo.Afiliados.numeroafiliado = N'" & LEGAJO & "')"




set rs=conexion.execute(cadena)

if rs.eof=true then
	rs.close
	set rs=nothing
	conexion.close
	set conexion=nothing
	response.redirect("financiera_4.asp")
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
.style43 {color: #F9760C}
.style47 {color: #000000; font-size: 18px; }
.style49 {color: #FFFFFF; font-size: 18px; }
.style50 {color: #FFFFFF}
.style51 {font-size: 14px; color: #FFFFFF; }
.style52 {color: #000000}
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
                          <td width="13" height="40">&nbsp;</td>
                          <td width="18"><img src="images/spacer12.gif" width="10" height="10" /></td>
                          <td colspan="4"><span class="style37"><span class="style42">INFORMACION DEL AFILIADO</span> 
                            <%
						  


						  
						  
						  %></span></td>
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
                          <td class="style33 style38">&nbsp;</td>
                          <td colspan="3" class="style33">&nbsp;</td>
                        </tr>
						<% if (rs.fields(11)=8) then%>
						 <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td class="style33 style38">Fecha Informado:</td>
                          <td colspan="3" class="style33">
						  <% if (rs.fields(10) <> "") then%>
								<%=rs.fields(10)%> 
							<%else%>  
							<%response.write("No informado")%>
								
						  <% end if%>
						  </td>
                        </tr>
						<% end if%>
						
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td class="style33 style38">
						  
						  <%RS.CLOSE
						  SET RS=NOTHING%>
                          &nbsp;</td>
                          <td colspan="3">&nbsp;</td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td><img src="images/spacer12.gif" width="10" height="10" /></td>
                          <td colspan="4" class="style33 style38"><span class="style40"> CONSOLIDACION DE DEUDA Y ESTADO DE CREDITO. 
                              <%
						  
CADENA="SELECT     tipoinstrumento, SUM(monto) AS Expr1, SUM(restapagar) AS Expr2 FROM         dbo.INSTRUMENTOS"
CADENA=CADENA & " WHERE     (numeroafiliado = N'" & LEGAJO & "') GROUP BY tipoinstrumento						  "

SET RS=CONEXION.EXECUTE(CADENA)

%>
                          </span></td>
                          </tr>
                        <tr>
                          <td height="20">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td colspan="4" class="style33 style38">&nbsp;</td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td colspan="4" class="style33 style38">
<%IF RS.EOF=FALSE THEN%>                          
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td width="95%" height="40" bgcolor="#384858"><table width="100%" border="0" cellspacing="1" cellpadding="0">
                                <tr>
                                  <td width="25%" height="40" bgcolor="#384858">&nbsp;</td>
                                  <td width="25%" bgcolor="#384858"><div align="center" class="style42 style50">TOTAL HISTORICO</div></td>
                                  <td width="25%" bgcolor="#384858"><div align="center" class="style42">
                                    <blockquote>
                                      <p class="style50">TOTAL SALDADO</p>
                                    </blockquote>
                                  </div></td>
                                  <td width="25%" bgcolor="#384858"><div align="center" class="style51">DEUDA ACTUAL</div></td>
                                </tr>
<%DO UNTIL RS.EOF%>                                
                                <tr>
                                  <td height="40" bgcolor="#FFFFFF"><span class="style43">&nbsp;&nbsp;&nbsp;<%=RS.FIELDS(0)%></span></td>
                                  <td bgcolor="#FFFFFF"><div align="right" class="style47">$<%=FORMATNUMBER(RS.FIELDS(1),2)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%C1=C1+RS.FIELDS(1)%></div></td>
                                  <td bgcolor="#FFFFFF"><div align="right" class="style47">$<%=FORMATNUMBER(RS.FIELDS(1)-RS.FIELDS(2),2)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%C2=C2+(RS.FIELDS(1)-RS.FIELDS(2))%></div></td>
                                  <td bgcolor="#DBDBDB"><div align="right" class="style47">$<%=FORMATNUMBER(RS.FIELDS(2),2)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%C3=C3+RS.FIELDS(2)%></div></td>
                                </tr>

<%		RS.MOVENEXT
LOOP 
RS.CLOSE
SET RS=NOTHING
%>         
                                <tr>
                                  <td height="40" bgcolor="#FFFFFF"><span class="style43">&nbsp;&nbsp;&nbsp;</span>TOTALES CONSOLIDADOS</td>
                                  <td bgcolor="#FFFFFF"><div align="right"><span class="style47">$<%=FORMATNUMBER(C1,2)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></div></td>
                                  <td bgcolor="#FFFFFF"><div align="right"><span class="style47">$<%=FORMATNUMBER(C2,2)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></div></td>
                                  <td bgcolor="#FF0000"><div align="right"><span class="style49">$<%=FORMATNUMBER(C3,2)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></div></td>
                                </tr>                       
                              </table></td>
                              <td width="5%">&nbsp;</td>
                            </tr>
                          </table>
                          <p>
                            <%ELSE%> 
                            </p>
                          <p>No existe información financiera o de crédito para este afiliado.</p>
                          <p>
                            <%END IF%>                          
                          </p></td>
                          </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td class="style33 style38">&nbsp;</td>
                          <td colspan="3">&nbsp;</td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td><img src="images/spacer12.gif" width="10" height="10" /></td>
                          <td colspan="4" class="style33 style38"><span class="style40">DETALLE HISTORICO Y ACTUAL (ORDENADO POR FECHA DESCENDENTE)
                              <%


CADENA="SELECT     dbo.INSTRUMENTOS.fechasolicitud, dbo.INSTRUMENTOS.tipoinstrumento, dbo.INSTRUMENTOS.monto, dbo.INSTRUMENTOS.restapagar, "
CADENA=CADENA & "                       dbo.INSTRUMENTOS.cantidadpagos, dbo.Usuarios.nombre, dbo.COMERCIOS.Comercio,  dbo.INSTRUMENTOS.instid"
CADENA=CADENA & " FROM         dbo.INSTRUMENTOS INNER JOIN"
CADENA=CADENA & "                       dbo.Usuarios ON dbo.INSTRUMENTOS.userid = dbo.Usuarios.userid LEFT OUTER JOIN"
CADENA=CADENA & "                       dbo.COMERCIOS ON dbo.INSTRUMENTOS.comercio = dbo.COMERCIOS.cid "
CADENA=CADENA & " WHERE     (dbo.INSTRUMENTOS.numeroafiliado = N'" & LEGAJO & "') ORDER BY dbo.INSTRUMENTOS.fechasolicitud DESC "



SET RS=CONEXION.EXECUTE(CADENA)
						  
						  %></span></td>
                          </tr>
                        <tr>
                          <td height="20">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td colspan="4" class="style33 style38">&nbsp;</td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td colspan="4" class="style33 style38">
                          
                          <%IF RS.EOF=FALSE THEN%>
                          
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td width="95%" height="40" bgcolor="#384858"><table width="100%" border="0" cellspacing="1" cellpadding="0">
                                <tr>
                                  <td width="11%" height="40"><div align="center"><span class="style42 style50">FECHA</span></div></td>
                                  <td width="12%"><div align="center"><span class="style42 style50">TIPO</span></div></td>
                                  <td width="23%"><div align="center"><span class="style42 style50">A FAVOR DE</span></div></td>
                                  <td width="12%"><div align="center"><span class="style42 style50">MONTO</span></div></td>
                                  <td width="10%"><div align="center"><span class="style42 style50">DEUDA</span></div></td>
                                  <td width="6%"><div align="center"><span class="style42 style50">CTS</span></div></td>
                                  <td width="21%"><div align="center"><span class="style42 style50">EMITIDO POR</span></div></td>
                                  <td width="5%">&nbsp;</td>
                                </tr>
                                <%do until rs.eof
								if rs.fields(3) >0 then 
									micolor="#FCB772"
								else
									micolor="#ffffff"
								end if
								%>
                                <tr>
                                  <td height="40" bgcolor="<%=micolor%>"><div align="center" class="style52"><%=MID(RS.FIELDS(0),7,2) & "/" & MID(RS.FIELDS(0),5,2) & "/" & LEFT(RS.FIELDS(0),4)%>&nbsp;</div></td>
                                  <td bgcolor="<%=micolor%>"><div align="center" class="style52">
                                    <div align="left">&nbsp;<%=RS.FIELDS(1)%></div>
                                  </div></td>
                                  <td bgcolor="<%=micolor%>"><div align="center" class="style52">
                                    <div align="left">&nbsp;<%
									IF ISNULL(RS.FIELDS(6))=TRUE THEN
										RESPONSE.WRITE("S.T.M.E.")
									ELSE
										RESPONSE.WRITE(RS.FIELDS(6))
									END IF
									
									%></div>
                                  </div></td>
                                  <td bgcolor="<%=micolor%>"><div align="center" class="style52">
                                    <div align="right">$<%=FORMATNUMBER(RS.FIELDS(2),2)%>&nbsp;&nbsp;&nbsp;</div>
                                  </div></td>
                                  <td bgcolor="<%=micolor%>"><div align="center" class="style52">
                                    <div align="right">$<%=FORMATNUMBER(RS.FIELDS(3),2)%>&nbsp;&nbsp;&nbsp;</div>
                                  </div></td>
                                  <td bgcolor="<%=micolor%>"><div align="center" class="style52"><%=RS.FIELDS(4)%>&nbsp;</div></td>
                                  <td bgcolor="<%=micolor%>"><div align="center" class="style52">
                                    <div align="left">&nbsp;&nbsp;&nbsp;<%=RS.FIELDS(5)%>&nbsp;</div>
                                  </div></td>
                                  <td bgcolor="#FFFFFF"><div align="center"><a href="financiera_41detalle.asp?iid=<%=rs.fields(7)%>&amp;legajo=<%=LEGAJO%>"><strong><img src="images/LUPA.jpg" width="30" height="29" border="0" /></strong></a></div></td>
                                </tr>
                                <%rs.movenext
								loop%>
                              </table></td>
                              <td width="5%">&nbsp;</td>
                            </tr>
                          </table>
                          
                          <p>
                            <%ELSE%>
                            </p>
                          <p>No existe información financiera o de crédito para este afiliado.                            </p>
                          <p>
                            <%END IF
							rs.close
							set rs=nothing%>
                          </p></td>
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
