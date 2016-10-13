<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*24*") =0 then response.redirect("sinacceso.asp")%>
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
        <td height="400" valign="top" bgcolor="#FFFFFF"><p>&nbsp;</p>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                  <td width="96%" bgcolor="#384858"><table width="100%" border="0" cellspacing="1" cellpadding="0">
                                      <tr>
                                        <td width="7%" height="40"><div align="center"><span class="style35">Mov</span></div></td>
                                        <td width="7%" height="40"><div align="center"><span class="style35">Cod Rech</span></div></td>
										<td width="15%"><div align="center"><span class="style35"> Legajo</span></div></td>
                                        <td width="15%"><div align="center"><span class="style35">Apellido y Nombre</span></div></td>
                                        <td width="13%"><div align="center"><span class="style35">Total </span></div></td>
										<td width="13%"><div align="center"><span class="style35">Secuencia </span></div></td>
                                      </tr>
									  
									  <%
										Const ForReading = 1
										Const Create = False
										Dim FSysObj
										Dim TS
										Dim strLine
										Dim strFileName
										Dim actual
										actual = DateAdd("m",-1,Now())
										'nombre del fichero a mostrar
										'strFileName = Server.MapPath("salud\ficherofuente.txt")
										fileName = "salud\p\" & year(actual) & right("0"& month(actual),2) & ".txt"
										strFileName = Server.MapPath(fileName)
										'Creación del objeto FileSystemObject
										Set FSysObj = Server.CreateObject("Scripting.FileSystemObject")

										' Abrimos el fichero
										Set TS = FSysObj.OpenTextFile(strFileName, ForReading, Create)

										If not TS.AtEndOfStream Then
										Response.Write "<FONT FACE=Verdana SIZE=1>"
										Do While not TS.AtendOfStream

										' Leemos el fichero linea a linea y lo mostramos
										strLine = TS.ReadLine
										
																				
										strMov = Mid(strLine,1,1) '1 Código Movimiento
										strCodigo = Mid(strLine,2,3) 'Código Entidad	A	3
										strMesAnio = Mid(strLine,5,6) '2 + 4 'Mes de Proceso (MMAAAA)	N	6
										strLeg =Mid(strLine,11,10) 'Legajo	N	10
										strCodigo = Mid(strLine,21,3)'Concepto	N	3
										strTotal = Mid(strLine,24,9) /100 'Importe (2 decimales sin “coma”)	N	9
										strVto = Mid(strLine, 33, 6) 'Mes Vencimiento (MMAAAA)	N	6
										strComent = Mid(strLine, 39, 10) 'Comentario	A	10
										strApyn = Mid(strLine,49,20) 'Apellido y Nombre	A	20
										strRech = Mid(strLine,69,2) 'Código Rechazo	A	2
										'response.write("AAAA " & Trim(strRech))
										if (Trim(strRech) = "")then
										color = "#E5E5E5"
										else
										color = "RED"
										end if
										strLegN = Mid(strLine,71,10)'Nuevo Legajo	N	10
										
										'Devolución	A	1
										'Dependencia Pago	A	4
										'Código de Partido	A	3
										'Blancos	A	35
										strSec = Mid(strLine,124,5)'Número Secuencial	N	5
										
									  %>
                                      <tr>
                                        <td  height="40" bgcolor="<%= color %>"><div align="center"><%=strMov%></div></td>
										<td  height="40" bgcolor="<%= color %>"><div align="center"><%=strRech%></div></td>
                                        <td  height="40" bgcolor="<%= color %>"><div align="center"><%=strLeg%></div></td>
										<td  height="40" bgcolor="<%= color %>"><div align="center"><%=strApyn%></div></td>
										<td  height="40" bgcolor="<%= color %>"><div align="center"><%=strTotal%></div></td>	
										<td  height="40" bgcolor="<%= color %>"><div align="center"><%=strSec%></div></td>												
  <%
  
											loop
										End If

										' cerramos y destruimos los objetos
										TS.Close
										Set TS = Nothing
										Set FSysObj = Nothing

%> 
                                        
									</tr>
									
				</table>
				<center>
				<form method=post action="financiera_salud_procesar.asp"> 
				<input type="hidden" id = "nombreArchivo" value="<%= fileName %>"/>
				<input type="submit" value="<< Procesar >>"/>
				</form>
				</center>
</td>
</tr>
</table>
		</body> </html>
