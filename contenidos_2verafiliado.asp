<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*11*") =0 then response.redirect("sinacceso.asp")
IF REQUEST("AID")="" THEN RESPONSE.REDIRECT("contenidos.asp")
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
.style35 {font-family: Arial, Helvetica, sans-serif; font-weight: bold; color: #FFFFFF; }
.style46 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; }
.style52 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; color: #384858; }
.style53 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 14px;
}
.style55 {color: #FF0000; font-weight: bold; }
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
        <td height="400" valign="top" bgcolor="#FFFFFF"><p><%
	


cadena="SELECT     TOP (100) PERCENT nombre, numeroafiliado, documento, fechanacimiento, estadocivil, tipoafiliado, cantidadhijos, email, domicilio, telefono, "
cadena=cadena & "                       fechaingreso, situacionanses, mutual, grupofamiliar, agrupamiento, categoria, lugartrabajo, situacionlaboral, cuentabco"
cadena=cadena & ", beneficiario FROM         dbo.Afiliados"
cadena=cadena & " WHERE     (id = " & request("aid") & ")"

set rs=conexion.execute(cadena)
if rs.eof=false then
	nombre=rs.fields(0)
	legajo=rs.fields(1)
	documento=rs.fields(2)
	fechanacimiento=rs.fields(3)
	estadocivil=rs.fields(4)
	tipoafiliado=rs.fields(5)
	hijos=rs.fields(6)
	email=rs.fields(7)
	domicilio=rs.fields(8)
	telefono=rs.fields(9)
	fechaingreso=rs.fields(10)
	anses=rs.fields(11)
	mutual=rs.fields(12)
	grupofamiliar=rs.fields(13)
	agrupamiento=rs.fields(14)
	categoria=rs.fields(15)
	lugartrabajo=rs.fields(16)
	situacionlaboral=rs.fields(17)
	cuentabanco=rs.fields(18)
	beneficiario=rs.fields(19)
end if
rs.close
set rs=nothing

		
		%>&nbsp;</p>
          <form id="form1" name="form1" method="post" action="contenidos_2save.asp" >

            <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="30" bgcolor="#384858"><div align="center"><span class="style35">MODIFICACION  DE AFILIADO</span>
                  <input name="ID" type="hidden" id="ID" value="<%=request("aid")%>" />
                </div></td>
              </tr>
              <tr>
                <td bgcolor="#374957"><table width="800" border="0" cellspacing="1" cellpadding="0">
                    <tr>
                      <td bgcolor="#FFFFFF"><table width="800" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0">&nbsp;</td>
                          </tr>
                          <tr>
                            <td width="22" height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td width="180" bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Nombre y apellido:</span></td>
                            <td width="598" bgcolor="#F0F0F0"><input name="F1" type="text" id="F1" value="<%=nombre%>" size="50" maxlength="50"  /></td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Número de afiliado:</span></td>
                            <td bgcolor="#F0F0F0"><input name="F40" type="text" id="F2"  onkeyup="this.value = this.value.replace (/[^0-9]/, ''); " value="<%=legajo%>" size="10" maxlength="10"/></td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Documento de identidad (DNI):</span></td>
                            <td bgcolor="#F0F0F0"><input name="F2" type="text" id="textfield49" onkeyup="this.value = this.value.replace (/[^0-9]/, ''); " value="<%=documento%>" size="50" maxlength="50" /></td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Fecha de Nacimiento:</span></td>
                            <td bgcolor="#F0F0F0"><%
							if len(fechanacimiento) > 4 then
							xdia=int(mid(fechanacimiento,7,2))
							xmes=int(mid(fechanacimiento,5,2))
							xano=int(mid(fechanacimiento,1,4))
							end if
							%><select name="F3_2" class="style46" id="select58">
                              <%for t=1 to 31%>
                              <option  <%if int(t)=int(xdia) then response.write("selected")%> value="<%=t%>"><%=t%></option>
                              <%next%>
                            </select>
/
<select name="F3_1" class="style46" id="select57">
                                <option <%if xmes=1 then response.write("selected")%> value="1">Enero</option>
                                <option <%if xmes=2 then response.write("selected")%> value="2">Febrero</option>
                                <option <%if xmes=3 then response.write("selected")%> value="3">Marzo</option>
                                <option <%if xmes=4 then response.write("selected")%> value="4">Abril</option>
                                <option <%if xmes=5 then response.write("selected")%> value="5">Mayo</option>
                                <option <%if xmes=6 then response.write("selected")%> value="6">Junio</option>
                                <option <%if xmes=7 then response.write("selected")%> value="7">Julio</option>
                                <option <%if xmes=8 then response.write("selected")%> value="8">Agosto</option>
                                <option <%if xmes=9 then response.write("selected")%> value="9">Septiembre</option>
                                <option <%if xmes=10 then response.write("selected")%> value="10">Octubre</option>
                                <option <%if xmes=11 then response.write("selected")%> value="11">Noviembre</option>
                                <option <%if xmes=12 then response.write("selected")%> value="12">Diciembre</option>
                              </select>
                              /
                              
                              <select name="F3_3" class="style46" id="select59">
                                <%for t=1950 to year(date)%>
                                <option <%if int(t)=int(xano) then response.write("selected")%> value="<%=t%>"><%=t%></option>
                                <%next%>
                              </select></td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Estado Civil:</span></td>
                            <td bgcolor="#F0F0F0"><select name="F4" class="style46" id="select56">
                              <option <%if estadocivil="Soltero" then response.write(" selected ")%> value="Soltero">Soltero/a</option>
                              <option <%if estadocivil="Casado" then response.write(" selected ")%> value="Casado">Casado/a</option>
                              <option <%if estadocivil="Separado" then response.write(" selected ")%> value="Separado">Separado/a</option>
                              <option <%if estadocivil="Divorciado" then response.write(" selected ")%> value="Divorciado">Divorciado/a</option>
                              <option <%if estadocivil="Viudo" then response.write(" selected ")%> value="Viudo">Viudo</option>
                              <option <%if estadocivil="Concubinato" then response.write(" selected ")%> value="Concubinato">En concubinato</option>
                              <option <%if estadocivil="Union Civil" then response.write(" selected ")%> value="Union Civil">Union Civil (mismo sexo)</option>
                                                          </select>                            </td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Tipo de Afiliado:</span></td>
                            <td bgcolor="#F0F0F0"><%
cadena="SELECT     TOP (100) PERCENT id, tipo FROM         dbo.tipo_afiliado ORDER BY tipo"
set rs=conexion.execute(cadena)
							%><select name="F5" class="style46" id="select55" <% if tipoafiliado = "8" then response.write(" disabled=disabled " ) end if %>>
                            <%do until rs.eof%>
                              <option <%if tipoafiliado= rs.fields(0) then response.write(" selected ")%> value="<%=rs.fields(0)%>"><%=rs.fields(1)%></option>
                            <%rs.movenext
							loop%>  
                            </select>
<%rs.close
set rs=nothing%>                            </td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Cantidad de hijos:</span></td>
                            <td bgcolor="#F0F0F0"><select name="F6" class="style46" id="select54">
                                <%FOR T=0 TO 15%>
                                <option <%if int(hijos)=int(t) then response.write(" selected ")%> value="<%=T%>"><%=T%></option>
                                <%NEXT%>
                            </select></td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52">Email:</span></td>
                            <td bgcolor="#F0F0F0"><input name="F7" type="text" id="textfield18" value="<%=email%>" size="50" maxlength="50" /></td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Domicilio:</span></td>
                            <td bgcolor="#F0F0F0"><input name="F8" type="text" id="textfield50" value="<%=domicilio%>" size="80" maxlength="100" /></td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52">Teléfono:</span></td>
                            <td bgcolor="#F0F0F0"><input name="F9" type="text" id="textfield17" value="<%=telefono%>" size="50" maxlength="50" /></td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Fecha de ingreso:</span></td>
                            <td bgcolor="#F0F0F0"><%
							if len(fechaingreso) > 4 then							
								xdia=int(mid(fechaingreso,7,2))
								xmes=int(mid(fechaingreso,5,2))
								xano=int(mid(fechaingreso,1,4))
							end if
							%><select name="F10_2" class="style46" id="select52">
                              <%for t=1 to 31%>
                              <option <%if int(t)=int(xdia) then response.write("selected")%> value="<%=t%>"><%=t%></option>
                              <%next%>
                            </select>
/
<select name="F10_1" class="style46" id="select51">
                                <option <%if xmes=1 then response.write("selected")%> value="1">Enero</option>
                                <option <%if xmes=2 then response.write("selected")%> value="2">Febrero</option>
                                <option <%if xmes=3 then response.write("selected")%> value="3">Marzo</option>
                                <option <%if xmes=4 then response.write("selected")%> value="4">Abril</option>
                                <option <%if xmes=5 then response.write("selected")%> value="5">Mayo</option>
                                <option <%if xmes=6 then response.write("selected")%> value="6">Junio</option>
                                <option <%if xmes=7 then response.write("selected")%> value="7">Julio</option>
                                <option <%if xmes=8 then response.write("selected")%> value="8">Agosto</option>
                                <option <%if xmes=9 then response.write("selected")%> value="9">Septiembre</option>
                                <option <%if xmes=10 then response.write("selected")%> value="10">Octubre</option>
                                <option <%if xmes=11 then response.write("selected")%> value="11">Noviembre</option>
                                <option <%if xmes=12 then response.write("selected")%> value="12">Diciembre</option>
                              </select>
                              /
                              
                              <select name="F10_3" class="style46" id="select53">
                                <%for t=1950 to year(date)%>
                                <option <%if int(t)=int(xano) then response.write("selected")%> value="<%=t%>"><%=t%></option>
                                <%next%>
                              </select></td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Situacion IPS:</span></td>
                            <td bgcolor="#F0F0F0"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                  <td width="50%" class="style46"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td width="7%"><input name="F11_1" type="radio" id="radio11" value="activo" <%if anses="activo" then%>checked="checked"<%end if%> /></td>
                                      <td width="33%">Trabajador activo</td>
                                      <td width="7%"><input type="radio" name="F11_1" id="radio" value="jubilado" <%if anses="jubilado" then%>checked="checked"<%end if%>/></td>
                                      <td width="21%">Jubilado</td>
                                      <td width="7%"><input type="radio" name="F11_1" id="radio2" value="pensionado" <%if anses="pensionado" then%>checked="checked"<%end if%>/></td>
                                      <td width="25%">Pensionado</td>
                                    </tr>
									
                                  </table></td>
                                  <td width="3%" class="style46">&nbsp;</td>
                                  <td width="19%" class="style46">&nbsp;</td>
                                  <td width="28%" class="style46">&nbsp;</td>
                                </tr>
                            </table></td>
                          </tr>

                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Afiliado a mutual:</span></td>
                            <td bgcolor="#F0F0F0"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td width="4%" class="style46"><input name="F12" type="radio" id="radio3" value="IOMA" <%if mutual="IOMA" then%>checked="checked"<%end if%>/></td>
                                <td width="16%" class="style46">IOMA</td>
                                <td width="4%" class="style46"><input type="radio" name="F12" id="radio4" value="PAMI" <%if mutual="PAMI" then%>checked="checked"<%end if%> /></td>
                                <td width="11%" class="style46">PAMI</td>
                                <td width="17%" class="style46"><span class="style52"><span class="style32">*</span></span> Nº   Beneficiario:</td>
                                <td width="48%" class="style46"><input name="F11_2" type="text" id="F11_2" value="<%=beneficiario%>" size="20" maxlength="20" /></td>
                              </tr>
                            </table></td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52">Grupo Familiar:</span></td>
                            <td bgcolor="#F0F0F0"><textarea name="F13" cols="60" rows="5" id="textfield13"><%=grupofamiliar%></textarea></td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Agrupamiento:</span></td>
                            <td bgcolor="#F0F0F0"><%
cadena="SELECT     TOP (100) PERCENT id, agrupamiento FROM         dbo.agrupamientos ORDER BY agrupamiento"
set rs=conexion.execute(cadena)
							%><select name="F14" class="style46" id="select55">
                            <%do until rs.eof%>
                              <option <%if int(agrupamiento)=rs.fields(0) then response.write(" selected ")%> value="<%=rs.fields(0)%>"><%=rs.fields(1)%></option>
                            <%rs.movenext
							loop%>  
                            </select>
<%rs.close
set rs=nothing%>     </td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Categoria:</span></td>
                            <td bgcolor="#F0F0F0"><select name="F15" class="style46" id="select15">
                              <option <%if categoria="1" then response.write("selected")%> value="1" selected="selected">1</option>
                              <option <%if categoria="2" then response.write("selected")%> value="2">2</option>
                              <option <%if categoria="3" then response.write("selected")%> value="3">3</option>
                              <option <%if categoria="4" then response.write("selected")%> value="4">4</option>
                              <option <%if categoria="5" then response.write("selected")%> value="5">5</option>
                              <option <%if categoria="6" then response.write("selected")%> value="6">6</option>
                              <option <%if categoria="Otro" then response.write("selected")%> value="Otro">Otro</option>
                            </select></td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Lugar de trabajo:</span></td>
                            <td bgcolor="#F0F0F0"><input name="F16" type="text" id="textfield14" value="<%=lugartrabajo%>" size="50" maxlength="50" /></td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Situación laboral:</span></td>
                            <td bgcolor="#F0F0F0"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td width="50%" class="style46"><table width="50%" border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td width="8%"><input name="F17" type="radio" id="radio7" value="efectivo" <%if situacionlaboral="efectivo" then%> checked="checked" <%end if%> /></td>
                                      <td width="24%">Efectivo</td>
                                      <td width="8%"><input type="radio" name="F17" id="radio8" value="contratado" <%if situacionlaboral="contratado" then%> checked="checked" <%end if%> /></td>
                                      <td width="60%">Contratado</td>
                                    </tr>
                                </table></td>
                                </tr>
                            </table></td>
                          </tr>
                          
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0"><span class="style52"><span class="style32">*</span> Nº de cuenta Banco Pcia:</span></td>
                            <td bgcolor="#F0F0F0"><input name="F18" type="text" id="textfield15" value="<%=cuentabanco%>" size="50" maxlength="50" /> </td>
                          </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0">&nbsp;</td>
                            <td bgcolor="#F0F0F0">&nbsp;</td>
                          </tr>
                          <tr>
                            <td height="30" colspan="3" bgcolor="#EC7A00"><div align="center"><span class="style35">FAMILIARES A CARGO</span></div><%
							

cadena="SELECT     TOP (100) PERCENT id, sublegajo, nombre, parentesco, sexo, fechanacimiento FROM         dbo.afiliados_parientes"
cadena=cadena & " WHERE     (id = " & request("aid") & ") ORDER BY sublegajo"
set rs=conexion.execute(cadena)
dim f1(15)
dim f2(15)
dim f3(15)
dim f4(15)

for t=1 to 15
	f1(t)=""
	f2(t)=""
	f3(t)=""
	f4(t)=""
next
x=0
if rs.eof=false then
	do until rs.eof
		x=x+1
		f1(x)=rs.fields(2)
		f2(x)=rs.fields(3)
		f3(x)=rs.fields(4)
		f4(x)=rs.fields(5)
		
		rs.movenext
	loop
end if
rs.close
set rs=nothing
conexion.close
set conexion=nothing							
							
							%></td>
                            </tr>
                          <tr>
                            <td height="30" bgcolor="#F0F0F0">&nbsp;</td>
                            <td colspan="2" bgcolor="#F0F0F0"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td height="30" colspan="2" class="style52"><div align="center">Apellido y nombre</div></td>
                                <td width="19%" class="style52"><div align="center">Parentesco</div></td>
                                <td width="13%" class="style52"><div align="center">Sexo</div></td>
                                <td width="27%" class="style52"><div align="center">Fecha Nacimiento</div></td>
                                <td width="3%">&nbsp;</td>
                              </tr>
                              
                              <%FOR T=1 TO 15%>
                              
                              <tr>
                                <td width="3%" height="30" class="style52">/<%=T%></td>
                                <td width="35%" class="style52"><input name="FACN<%=T%>" type="text" class="style46" id="FACN<%=T%>" value="<%=f1(t)%>" size="50" maxlength="50" /></td>
                                <td class="style52"><div align="center">
                                  <select name="FACP<%=T%>" id="FACP<%=T%>">
                                    <option value=" " selected="selected"> </option>
                                    <option <%if f2(t)="Padre/Madre" then response.write(" selected ")%> value="Padre/Madre">Padre/Madre</option>
                                    <option <%if f2(t)="Hijo/a" then response.write(" selected ")%> value="Hijo/a">Hijo/a</option>
                                    <option <%if f2(t)="Esposo/a" then response.write(" selected ")%> value="Esposo/a">Esposo/a</option>
                                    <option <%if f2(t)="Abuelo/a" then response.write(" selected ")%> value="Abuelo/a">Abuelo/a</option>
                                    <option <%if f2(t)="Otro" then response.write(" selected ")%> value="Otro">Otro</option>
                                  </select>
                                </div></td>
                                <td class="style52"><div align="center">
                                  <select name="FACS<%=T%>" class="style46" id="FACS<%=T%>">
                                    <option  value=" " selected="selected"> </option>
                                    <option <%if f3(t)="Femenino" then response.write(" selected ")%> value="Femenino">Femenino</option>
                                    <option <%if f3(t)="Masculino" then response.write(" selected ")%> value="Masculino">Masculino</option>
                                  </select>
                                </div></td>
                                <td class="style52"><div align="center">
  <%
	if f4(t) <> "" then
		xdia=int(mid(f4(t),7,2))
		xmes=int(mid(f4(t),5,2))
		xano=int(mid(f4(t),1,4))  
	else
		xdia=0
		xmes=0
		xano=0
	end if
  
  %><select name="FACDIA<%=T%>" class="style46" id="FACDIA<%=T%>">
    <option value="" selected="selected"></option>
    <%for X=1 to 31%>
          <option <%if int(x)=int(xdia) then response.write("selected")%> value="<%=x%>"><%=x%></option>
    <%next%>
  </select>
                                  /
  <select name="FACMES<%=T%>" class="style46" id="FACMES<%=T%>">
    <option value="" selected="selected"></option>
                                <option <%if xmes=1 then response.write("selected")%> value="1">Enero</option>
                                <option <%if xmes=2 then response.write("selected")%> value="2">Febrero</option>
                                <option <%if xmes=3 then response.write("selected")%> value="3">Marzo</option>
                                <option <%if xmes=4 then response.write("selected")%> value="4">Abril</option>
                                <option <%if xmes=5 then response.write("selected")%> value="5">Mayo</option>
                                <option <%if xmes=6 then response.write("selected")%> value="6">Junio</option>
                                <option <%if xmes=7 then response.write("selected")%> value="7">Julio</option>
                                <option <%if xmes=8 then response.write("selected")%> value="8">Agosto</option>
                                <option <%if xmes=9 then response.write("selected")%> value="9">Septiembre</option>
                                <option <%if xmes=10 then response.write("selected")%> value="10">Octubre</option>
                                <option <%if xmes=11 then response.write("selected")%> value="11">Noviembre</option>
                                <option <%if xmes=12 then response.write("selected")%> value="12">Diciembre</option>
  </select>
                                  /
  <select name="FACANO<%=t%>" class="style46" id="FACANO<%=t%>">
      <option value="" selected="selected"></option>

    <%for X=1950 to year(date)%>
		<option <%if int(x)=int(xano) then response.write("selected")%> value="<%=x%>"><%=x%></option>
    <%next%>
  </select>
                                </div></td>
                                <td>&nbsp;</td>
                              </tr>
                              <%NEXT%>
                              <tr>
                                <td height="30" class="style52">&nbsp;</td>
                                <td class="style52">&nbsp;</td>
                                <td class="style52">&nbsp;</td>
                                <td class="style52">&nbsp;</td>
                                <td class="style52">&nbsp;</td>
                                <td>&nbsp;</td>
                              </tr>                              
                            </table></td>
                            </tr>
                      </table></td>
                    </tr>
                </table></td>
              </tr>
            </table>
                    <p align="center" class="style53"><span class="style53"><strong>Acción monitoreada. Usuario responsable: </strong></span><span class="style55"> <%=session("usuario")%></span> </p>
                    <p align="center" class="style53">
                      <input type="button" name="button" id="btnAceptar" value="Continuar" onclick= "document.form1.action = 'contenidos_2save.asp'; 
document.form1.submit()" />
					   <input type="button" name="button" id="btnBaja" value="Dar de baja" onclick= "document.form1.action = 'contenidos_2baja.asp'; 
document.form1.submit()"  />
</p>
          </form>
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
