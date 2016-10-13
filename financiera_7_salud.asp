<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*27*") =0 then response.redirect("sinacceso.asp")%>
<%
dia=int(day(date))
dia=21
if dia <=20 then
	response.redirect("financiera.asp")
end if

%>
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
.style37 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
.style46 {font-size: 14px}
.style49 {font-size: 12px; font-family: Geneva, Arial, Helvetica, sans-serif;}
.style36 {font-family: Geneva, Arial, Helvetica, sans-serif;
	color: #FFFFFF;
	font-size: 12px;
}
.style50 {color: #006633}
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
          <p align="center" class="style15 style50">PROCESO DE CIERRE DE LOTE MENSUAL PARA EMPLEADOS DE SALUD.</p>
          <p align="center" class="style15"><%
  
		  
MIMES=YEAR(DATE) & RIGHT("00" & MONTH(DATE),2)
CADENA="SELECT MES FROM INSTRUMENTOS_CIERREMES_SALUD WHERE MES=" & MIMES 
SET RS=CONEXION.EXECUTE(CADENA)
IF RS.EOF=FALSE THEN
	SECERRO=TRUE
ELSE
	SECERRO=FALSE
END IF


IF SECERRO=FALSE THEN
  %>      &nbsp;</p>
          <p align="center" class="style15 style32 style46">ATENCION!</p>
          <p align="center" class="style37">Una vez llevado a cabo este proceso, no puede ser cancelado ni revertido. </p>
          <p align="center" class="style49">Este proceso generará el listado para solicitar el descuento sobre créditos, bonos y adelantos de efectivo en los recibos de sueldo de los afiliados.</p>
          <p align="center" class="style49">Presione el siguiente botón para comenzar el proceso.</p>
          <form id="form1" name="form1" method="post" action="financiera_71_SALUD.asp">
            <div align="center"><br />
			
			
			
			 <%

CADENA="select numeroafiliado, nombre, documento,cuentabco, 'ALTA' as tipoinstrumento, 0 restapagar from [dbo].[afiliados] where tipoafiliado=8 and Informado is null"
CADENA = CADENA & " UNION ALL "
								CADENA=CADENA &" SELECT     TOP (100) PERCENT  dbo.afiliados.numeroafiliado, nombre, documento,cuentabco,"
								CADENA=CADENA & " dbo.instrumentos_cuotas.tipoinstrumento,dbo.instrumentos_cuotas.restapagar"
								CADENA=CADENA & " FROM         dbo.afiliados INNER JOIN"
								CADENA=CADENA & "                       dbo.instrumentos ON dbo.afiliados.numeroafiliado = dbo.instrumentos.numeroafiliado INNER JOIN"
								CADENA=CADENA & "                       dbo.instrumentos_cuotas ON dbo.instrumentos.instid = dbo.instrumentos_cuotas.instid"
								CADENA=CADENA & " WHERE     (dbo.instrumentos_cuotas.fechacobro =" & MIMES &") AND (dbo.instrumentos_cuotas.restapagar > 0) AND (dbo.afiliados.tipoafiliado = 8)"
CADENA=cADENA & " UNION ALL	"
CADENA=CADENA &"SELECT     dbo.afiliados.numeroafiliado, nombre, documento,cuentabco, 'CREDITO' as tipoinstrumento, monto restapagar"
			CADENA=CADENA & " FROM         dbo.instrumentos INNER JOIN"
			CADENA=CADENA & "                       dbo.afiliados ON dbo.instrumentos.numeroafiliado = dbo.afiliados.numeroafiliado"
			CADENA=CADENA & " WHERE     (dbo.instrumentos.tipoinstrumento = 'CREDITO') AND (dbo.instrumentos.creditodescargado = 0) AND (dbo.afiliados.tipoafiliado = 8)			"
'response.write(cADENA)
SET RS=CONEXION.EXECUTE(CADENA)
cantAltas = 0
cantPrestamos=0
totalPrestamos = 0
						  %>
			
				<%IF RS.EOF=FALSE THEN%>
                          
                          <table width="100%" border="1" cellspacing="0" cellpadding="0">
                            <tr>
                              <td width="95%" height="40" bgcolor="#ffffff"><table width="100%" border="0" cellspacing="1" cellpadding="0">
                                <tr>
                                  <td width="10%" height="40"><div align="center"><span class="style42 style50">LEGAJO</span></div></td>
                                  <td width="35%"><div align="center"><span class="style42 style50">NOMBRE</span></div></td>
                                  <td width="15%"><div align="center"><span class="style42 style50">DOCUMENTO</span></div></td>
                                  <td width="20%"><div align="center"><span class="style42 style50">CUENTABCO</span></div></td>
                                  <td width="10%"><div align="center"><span class="style42 style50">CONCEPTO</span></div></td>
								  <td width="10%"><div align="center"><span class="style42 style50">IMPORTE</span></div></td>
                              
                                </tr>
                              <%do until rs.eof%>
							  <% cantAltas = cantAltas +1%>
                                <tr>
                                  <td height="40" ><div align="left" class="style52"><%=RS.FIELDS(0)%>&nbsp;</div></td>
								  
                                  <td ><div align="left" class="style52">
                                    <div align="left">&nbsp;<%=RS.FIELDS(1)%></div>
                                  </div></td>
								  
                                  <td><div align="left" class="style52">
                                    <div align="left"><%=RS.FIELDS(2)%>&nbsp;&nbsp;&nbsp;</div>
                                  </div></td>
								  
								   <td><div align="left" class="style52">
                                    <div align="left"><%=RS.FIELDS(3)%>&nbsp;&nbsp;&nbsp;</div>
                                  </div></td>
								  
                                  <td ><div align="left" class="style52">
                                    <div align="left"><%=RS.FIELDS(4)%>&nbsp;&nbsp;&nbsp;</div>
                                  </div></td>
								  
                                   <td ><div align="left" class="style52">
                                    <div align="left"><%=RS.FIELDS(5)%>&nbsp;&nbsp;&nbsp;</div>
                                  </div></td>
                                </tr>
                                <%rs.movenext
								loop%>
								
								
							  </table></td>
                              
                            </tr>
                          </table>
                          
                          <p>
                            <%ELSE%>
                            </p>
                          <p>No existe información a generar.                            </p>
                          <p> 
                            <%END IF%>
			
                          </p></td>
                        </tr>
                        <tr>
                          <td height="40">&nbsp;</td>
                          <td>&nbsp;</td>
                          <td colspan="4" class="style33 style38">&nbsp;</td>
                        </tr>
                    </table>
				<span>Cantidad de altas informadas: <%=cantAltas%></span><BR/>
				<span>Cantidad de prestamos: <%=cantPrestamos%></span><BR/>
				<span>Total en prestamos: $<%=FORMATNUMBER(totalPrestamos,2)%></span><BR/><BR/><BR/>
				<center>
              <input type="submit" name="button" id="button" value="Iniciar proceso unico mensual" onclick="javascript:return confirm('Cerrar cierre de lote?')" />
			  </center>
            </div>
          </form>
          <form id="form2" name="form2" method="post" action="financiera.asp">
            <div align="center"><br />
              <input type="submit" name="button2" id="button2" value="Cancelar" />
              </div>
          </form>
          <p align="center" class="style37">&nbsp;</p>
          <p>
            <%ELSE%>
          </p>
          <p>&nbsp;</p>
          <p>&nbsp;</p>
          <p align="center" class="style15 style32">EL MES EN CURSO HA SIDO CERRADO. </p>
          <p align="center" class="style37">ESTA OPCION SE ENCUENTRA DESHABILITADA HASTA EL MES ENTRANTE </p>
          <p align="center" class="style37">&nbsp;</p>
          <p align="center" class="style37">&nbsp;</p>
          <p align="center" class="style37">&nbsp;</p>
          <%END IF%>
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
<%
RS.CLOSE
SET RS=NOTHING
CONEXION.CLOSE
SET CONEXION=NOTHING%>
