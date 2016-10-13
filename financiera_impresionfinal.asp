<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include FILE="./Includes/cnx.inc" -->
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*21*") =0 then response.redirect("sinacceso.asp")%>
<%legajo=trim(request("legajo"))
id=request("id")
if isnumeric(legajo) = false then response.redirect("financiera_1.asp")
if isnumeric(id) = false then response.redirect("financiera_1.asp")




'CADENA="SELECT     dbo.INSTRUMENTOS.monto, dbo.Afiliados.nombre, dbo.Afiliados.numeroafiliado, dbo.Afiliados.documento, dbo.INSTRUMENTOS.instid, "
'CADENA=CADENA & "                       dbo.INSTRUMENTOS.tipoinstrumento FROM         dbo.INSTRUMENTOS INNER JOIN"
'CADENA=CADENA & "                       dbo.Afiliados ON dbo.INSTRUMENTOS.numeroafiliado = dbo.Afiliados.numeroafiliado"
'CADENA=CADENA & " WHERE     (dbo.INSTRUMENTOS.tipoinstrumento = N'" & request("tipo") & "') AND (dbo.INSTRUMENTOS.instid = " & id & ") "
'CADENA=CADENA & " AND (dbo.Afiliados.numeroafiliado = N'" & legajo & "')"


CADENA="SELECT     dbo.INSTRUMENTOS.monto, dbo.Afiliados.nombre, dbo.Afiliados.numeroafiliado, dbo.Afiliados.documento, dbo.INSTRUMENTOS.instid, "
CADENA=CADENA & "                       dbo.INSTRUMENTOS.tipoinstrumento, dbo.INSTRUMENTOS.cantidadpagos, dbo.INSTRUMENTOS.concepto, dbo.COMERCIOS.Comercio, dbo.Afiliados.domicilio"
CADENA=CADENA & " FROM         dbo.INSTRUMENTOS INNER JOIN"
CADENA=CADENA & "                       dbo.Afiliados ON dbo.INSTRUMENTOS.numeroafiliado = dbo.Afiliados.numeroafiliado LEFT OUTER JOIN"
CADENA=CADENA & "                       dbo.COMERCIOS ON dbo.INSTRUMENTOS.comercio = dbo.COMERCIOS.cid"
CADENA=CADENA & " WHERE     (dbo.INSTRUMENTOS.tipoinstrumento = N'" & request("tipo") & "') AND (dbo.INSTRUMENTOS.instid = " & id & ") AND (dbo.Afiliados.numeroafiliado = N'" & legajo & "')"



set rs=conexion.execute(cadena)
if rs.eof=false then
	monto=rs.fields(0)
	nombre=rs.fields(1)
	legajo=rs.fields(2)
	dni=rs.fields(3)
	CUOTAS=RS.FIELDS(6)
	CONCEPTO=RS.FIELDS(7)
	COMERCIO=RS.FIELDS(8)
	DOMICILIO=RS.FIELDS(9)
end if
rs.close
set rs=nothing
conexion.close
set conexion=nothing







Dim xcen(9) 'centenas
Dim xdec(9) 'decenas
Dim xuni(9) 'unidades
Dim xexc(6) 'except
Dim ceros(9)

Function CONVERTIR(pnumero)

Dim letras
Dim i
Dim c
Dim j
Dim xnumero
Dim xnum
Dim num
Dim digito
Dim numero_ent
Dim entero
Dim decimales
Dim temp
  
  xcen(2) = "dosc"
  xcen(3) = "tresc"
  xcen(4) = "cuatrosc"
  xcen(5) = "quin"
  xcen(6) = "seisc"
  xcen(7) = "setec"
  xcen(8) = "ochoc"
  xcen(9) = "novec"
  xdec(2) = "veinti"
  xdec(3) = "trei"
  xdec(4) = "cuare"
  xdec(5) = "cincue"
  xdec(6) = "sese"
  xdec(7) = "sete"
  xdec(8) = "oche"
  xdec(9) = "nove"
  xuni(1) = "uno"
  xuni(2) = "dos"
  xuni(3) = "tres"
  xuni(4) = "cuatro"
  xuni(5) = "cinco"
  xuni(6) = "seis"
  xuni(7) = "siete"
  xuni(8) = "ocho"
  xuni(9) = "nueve"
  xexc(1) = "diez"
  xexc(2) = "once"
  xexc(3) = "doce"
  xexc(4) = "trece"
  xexc(5) = "catorce"
  xexc(6) = "quince"
  ceros(1) = "0"
  ceros(2) = "00"
  ceros(3) = "000"
  ceros(4) = "0000"
  ceros(5) = "00000"
  ceros(6) = "000000"
  ceros(7) = "0000000"
  ceros(8) = "00000000"
  
  c = 1
  i = 1
  j = 0
  
  xnumero = cStr(pnumero)
If Cdbl(LTrim(RTrim(pnumero))) < 999999999.99 Then
    numero_ent = Cdbl(Int(pnumero))
    If Len(numero_ent) < 9 Then
        numero_ent = ceros(9 - Len(numero_ent)) & numero_ent
    End If
    entero = Cdbl(Int(numero_ent))
    decimales = (Cdbl(xnumero) - entero) * 100
    
	Do While i < 8
        temp = 0
        num = Cdbl(Mid(numero_ent, i, 3))
        xnum = Mid(numero_ent, i, 3)
        digito = Cdbl(Mid(xnum, 1, 1))
        
        '/* analizo el numero entero de a 3 */
        If xnum = "000" Then
            j = 0
        Else
        	j = 1
            If digito > 1 Then
                letras = letras & xcen(digito) & "ientos "
            End If
            If Mid(xnum, 1, 1) = "1" And Mid(xnum, 2, 2) <> "00" Then
                letras = letras & "ciento "
            ElseIf Mid(xnum, 1, 1) = "1" Then
                letras = letras & "cien "
            End If
  
  			'/* analisis de las decenas */
            digito = Cdbl(Mid(xnum, 2, 1))
            If digito > 2 And Mid(xnum, 3, 1) = "0" Then
                letras = letras & xdec(digito) & "nta "
				temp = 1
            End If
            
			If digito > 2 And Mid(xnum, 3, 1) <> "0" Then
                letras = letras & xdec(digito) & "nta y "
				
            End If
            
			If digito = 2 And Mid(xnum, 3, 1) = "0" Then
                letras = letras & "veinte "
				temp = 1
            ElseIf digito = 2 And Mid(xnum, 3, 1) <> "0" Then
                letras = letras & "veinti"
				
            End If
            
			If digito = 1 And Mid(xnum, 3, 1) >= "6" Then
                letras = letras & "dieci"
            ElseIf digito = 1 And Mid(xnum, 3, 1) < "6" Then
                letras = letras & xexc(Cdbl(Mid(xnum, 3, 1) + 1))
            	temp = 1
			End If
        End If
   
   		if temp = 0 then
   	'/* analisis del ultimo digito */
        digito = Cdbl(Mid(xnum, 3, 1))
            	If ((c = 1) Or (c = 2)) And xnum = "001" Then
                	letras = letras & "un"
            	Else
                	If ((c = 1) Or (c = 2)) And xnum >= "020" And Mid(xnum, 3, 1) = "1" Then
                    	letras = letras & "un"
                	Else
                    	If digito <> 0 Then
                        	letras = letras & xuni(digito)
                    	End If
                	End If
            	End If
		end if
  
  If j = 1 And i = 1 And xnum = "001" And c = 1 Then
    letras = letras & " millon "
  ElseIf j = 1 And i = 1 And xnum <> "001" And c = 1 Then
    letras = letras & " millones "
  ElseIf j = 1 And i = 4 And c = 2 Then
    letras = letras & " mil "
  End If
  i = i + 3
  c = c + 1
  Loop
  If letras = "" Then
  letras = "cero "
  End If
  If decimales <> 0 Then
    decimales = Round(decimales)
    
    letras = letras & " con " & CStr(decimales) & "/100"
  End If
  
End If

'EN ESTA VARIABLE SESSION SE GUARDA EL NUMERO EN LETRAS 
session("Valor") = letras
End Function



x=convertir(monto)


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>STME - Intranet de gestion</title>
<style type="text/css">
<!--
.style2 {
	font-weight: bold;
	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 11px;
}
.style3 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 18px;
	line-height:26px;
}
.style12 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 10px;
}
.style13 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 14px;
	font-weight: bold;
}
.Estilo3 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 18px; color: #000000; }
.Estilo4 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 14px; font-weight: bold; color: #FF0000; }
.Estilo5 {color: #FF0000}
.Estilo6 {font-size: 16px}
.Estilo7 {font-weight: bold; font-family: Geneva, Arial, Helvetica, sans-serif;}
.style15 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 14px; }
.style17 {font-size: 14px}
.style18 {font-size: 14}
.style19 {font-size: 14px; font-weight: bold; }
-->
</style>
</head>

<body onload="window.print(); alert('Acción concluida. Presione el boton ACEPTAR para continuar.');window.close()" >
<div align="center">
  <table width="1000" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="550" valign="top"><p align="left"><img src="images/solologo.jpg" width="200" height="173" /><br />
          </p>
        <p align="right"><span class="Estilo3">
          <%if REQUEST("TIPO")="ADELANTO" then%>
          PAGARÉ NUMERO:
          <%ELSE%>
          AUTORIZACION NUMERO:
          <%END IF%>
          <%
		  
		  AHORA=MID(YEAR(DATE),3,2) & RIGHT("00" & MONTH(DATE),2) & RIGHT("00" & DAY(DATE),2)
		  midia= day(date) & "/" & month(date) & "/" & year(date) 
		  RESPONSE.WRITE(AHORA & "-" & LEGAJO & "-" & ID)%>
          </span> <br />
          <span class="style2">      UNICO DOMICILIO DE PAGO: Calle Sidotti Nro. 428, Ciudad de Ensenada (CP 1925, Pcia. de Buenos Aires)</span><br />
        </p>
<%if request("tipo")="CREDITO" then%>        <p align="left" class="style3">Ensenada, <%=midia%>. Por el presente, se autoriza al afiliado <%=nombre%>, legajo número <%=legajo%>, a obtener un crédito en <span class="Estilo5"><%=comercio%></span> para la adquisición de <%=concepto%>, en un todo de acuerdo a las condiciones de pago vigentes que declaro conocer y aceptar, en <%=cuotas%> cuotas consecutivas de $ <%=monto/cuotas%> por un monto total de compra de <span class="Estilo5">$<%=monto%> (Pesos <%=session("valor")%>)</span></p>
<p align="left" class="style3">&nbsp;</p>
<p align="left" class="style3"><br />
  <br />
  <br />
</p>
<p align="left" class="style3">&nbsp;</p><p align="left" class="style3">&nbsp;</p>
<p align="left" class="style3">&nbsp;</p>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="1" bgcolor="#000000" class="style12"></td>
    <td></td>
    <td bgcolor="#000000" class="style12"></td>
  </tr>
  <tr>
    <td width="45%" class="style12"><div align="left" class="style17"><strong>FIRMA AFILIADO</strong></div></td>
    <td width="10%">&nbsp;</td>
    <td width="45%" class="style12"><div align="left" class="style17"><strong>AUTORIZACION STME</strong></div></td>
  </tr>
  <tr>
    <td width="45%" class="style12"><div align="left"><span class="style17"><span class="style18"><span class="style17"><span class="style17"></span></span></span></span></div></td>
    <td width="10%">&nbsp;</td>
    <td width="45%" class="style12"><div align="left"><span class="style17"><span class="style18"><span class="style17"><span class="style17"></span></span></span></span></div></td>
  </tr>
  <tr>
    <td width="45%" class="style12"><div align="left" class="style17">NOMBRE Y APELLIDO: <%=nombre%></div></td>
    <td width="10%">&nbsp;</td>
    <td width="45%" class="style12"><div align="left" class="style17">NOMBRE Y APELLIDO:</div></td>
  </tr>
  <tr>
    <td width="45%" class="style12"><div align="left"><span class="style17"><span class="style18"><span class="style17"><span class="style17"></span></span></span></span></div></td>
    <td width="10%">&nbsp;</td>
    <td width="45%" class="style15">&nbsp;</td>
  </tr>
  <tr>
    <td width="45%" class="style12"><div align="left" class="style17">LEGAJO MUNICIPAL: <%=legajo%></div></td>
    <td width="10%">&nbsp;</td>
    <td width="45%" class="style15">&nbsp;</td>
  </tr>
</table>
<p>
    <%end if%>
        <%if request("tipo")="BONO" then%>
</p>
<p>&nbsp;      </p>
<p align="left" class="style3">Este bono ha sido emitido por el STME el día <%=day(date) & "/" & month(date) & "/" & year(date)%>, por la suma de  $<%=monto%> (Pesos <%=trim(session("valor"))%>) al afiliado/a <%=nombre%>, legajo número <%=legajo%></p>
      <p align="left" class="style13">&nbsp;</p>
      <p align="left" class="style13">&nbsp;</p>
      <p align="left" class="style13">&nbsp;</p>
      <p align="left" class="style13"><br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
      </p>
      <p align="left" class="style13">&nbsp;</p>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="1" bgcolor="#000000" class="style12"></td>
          <td></td>
          <td bgcolor="#000000" class="style12"></td>
        </tr>
        <tr>
          <td width="45%" class="style12"><div align="left" class="style19">FIRMA AFILIADO</div></td>
          <td width="10%">&nbsp;</td>
          <td width="45%" class="style12"><div align="left" class="style19">AUTORIZACION STME</div></td>
        </tr>
        <tr>
          <td width="45%" class="style12"><div align="left"><span class="style17"><span class="style18"><span class="style17"><span class="style17"></span></span></span></span></div></td>
          <td width="10%">&nbsp;</td>
          <td width="45%" class="style12"><div align="left"><span class="style17"><span class="style18"><span class="style17"><span class="style17"></span></span></span></span></div></td>
        </tr>
        <tr>
          <td width="45%" class="style12"><div align="left" class="style17">NOMBRE Y APELLIDO: <%=nombre%></div></td>
          <td width="10%">&nbsp;</td>
          <td width="45%" class="style12"><div align="left" class="style17">NOMBRE Y APELLIDO:</div></td>
        </tr>
        <tr>
          <td width="45%" class="style12"><div align="left"><span class="style17"><span class="style18"><span class="style17"><span class="style17"></span></span></span></span></div></td>
          <td width="10%">&nbsp;</td>
          <td width="45%" class="style12"><div align="left"><span class="style17"><span class="style18"><span class="style17"><span class="style17"></span></span></span></span></div></td>
        </tr>
        <tr>
          <td width="45%" class="style12"><div align="left" class="style17">LEGAJO MUNICIPAL: <%=legajo%></div></td>
          <td width="10%">&nbsp;</td>
          <td width="45%" class="style12"><div align="left"><span class="style17"><span class="style18"><span class="style17"><span class="style17"></span></span></span></span></div></td>
        </tr>
      </table>
      <br />
      <%end if%>
      <%if request("tipo")="ADELANTO" then%><p align="justify"><span class="style3">Pagaré sin protesto (Art. 50, Decreto Ley 5965/63) al Sindicato de Trabajadores Municipales de Ensenada, la suma de $<%=monto%> (Pesos <%=session("valor")%>) por igual valor recibido a mi entera satisfacción. Tipo de instrumento: <%=request("tipo")%><br />
      </span><span class="style12"><strong>CLAUSULA ESPECIAL: </strong>Autorizo al Sindicato de Trabajadores Municipales de Ensenada (STME) a retener la suma arriba indicada, mas $3 (pesos tres) en concepto de gastos administrativos, hasta los primeros $100 (Pesos cien) y a partir de allí $3 (pesos tres) por cada $ o fracción de $50 (pesos cincuenta) por igual concepto, de mis haberes mensuales o a debitarlo de mi caja de ahorro en el Banco de la Provincia de Buenos Aires, Sucursal Ensenada, de acuerdo a las fechas de inicio y cierre de cada período el cual se estima ocurrirá en el mes en curso. Se fija la jurisdicción de los Tribunales Ordinarios de la Ciudad de La Plata, con exclusión de todo otro que pudiere corresponder. EN CASO DE NO EFECTUARSE LA RETENCION MENSUAL CORRESPONDIENTE ME COMPROMETO A ABONAR LAS SUMAS ADEUDADAS AL STME O EN SU DEFECTO ACEPTAR LOS CARGOS ADMINISTRATIVOS QUE CORREPONDAN A CADA CIERRE MENSUAL POR LAS SUMAS ADEUDADAS.</span></p>
      <p align="justify">&nbsp;</p>
      <p align="justify">&nbsp;</p>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="1" bgcolor="#000000" class="style12"></td>
          <td></td>
          <td bgcolor="#FFFFFF" class="style12"></td>
        </tr>
        <tr>
          <td width="45%" class="style12"><div align="left"><strong>FIRMA AFILIADO (de puño y letra)</strong></div></td>
          <td width="10%">&nbsp;</td>
          <td width="45%" class="style12">&nbsp;</td>
        </tr>
        <tr>
          <td width="45%" class="style12"><div align="left"></div></td>
          <td width="10%">&nbsp;</td>
          <td width="45%" class="style12">&nbsp;</td>
        </tr>
        <tr>
          <td width="45%" class="style12"><div align="left">NOMBRE Y APELLIDO (DE PUÑO Y LETRA):</div></td>
          <td width="10%">&nbsp;</td>
          <td width="45%" class="style12">&nbsp;</td>
        </tr>
        <tr>
          <td width="45%" class="style12"><div align="left"></div></td>
          <td width="10%">&nbsp;</td>
          <td width="45%" class="style12">&nbsp;</td>
        </tr>
        <tr>
          <td width="45%" class="style12"><div align="left">LEGAJO MUNICIPAL:</div></td>
          <td width="10%">&nbsp;</td>
          <td width="45%" class="style12">&nbsp;</td>
        </tr>
        <tr>
          <td width="45%" class="style12"><div align="left"></div></td>
          <td width="10%">&nbsp;</td>
          <td width="45%" class="style12">&nbsp;</td>
        </tr>
        <tr>
          <td width="45%" class="style12"><div align="left">DIRECCION:</div></td>
          <td width="10%">&nbsp;</td>
          <td width="45%" class="style12">&nbsp;</td>
        </tr>
      </table>
      <%end if%>
      <p align="center" class="Estilo5 Estilo6"><span class="Estilo7">PARA COMPROBAR LA AUTENTICIDAD DE ESTE INSTRUMENTO, INGRESE A HTTP://BONOS.STME.ORG Y SIGA LAS INSTRUCCIONES SOLCITADAS POR NUESTRO SISTEMA. </span>
      </p></td>
    </tr>
  </table>
  <p>&nbsp;</p>
  
  <%if request("tipo")<>"ADELANTO" then%>
  <table width="1000" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="550" valign="top"><p align="left"><img src="images/solologo.jpg" width="200" height="173" /><br />
        <br />
            <span class="Estilo3">
            
            PAGARÉ NUMERO: 

              <%
		  
		  AHORA=MID(YEAR(DATE),3,2) & RIGHT("00" & MONTH(DATE),2) & RIGHT("00" & DAY(DATE),2)
		  
		  RESPONSE.WRITE(AHORA & "-" & LEGAJO & "-" & ID)%>
            </span> <br />
            <span class="style2">UNICO DOMICILIO DE PAGO: Calle Sidotti Nro. 428, Ciudad de Ensenada (CP 1925, Pcia. de Buenos Aires)</span><br />
          </p>
        <p align="justify"><span class="style3"><span class="style3">Pagaré sin protesto (Art. 50, Decreto Ley 5965/63) al Sindicato de Trabajadores Municipales de Ensenada, la suma de $<%=monto%> (Pesos <%=session("valor")%>) por igual valor recibido a mi entera satisfacción.</span> Tipo de instrumento: <%=request("tipo")%><br />
          <br />
        </span><span class="style15"><strong>CLAUSULA ESPECIAL: </strong>Autorizo al Sindicato de Trabajadores Municipales de Ensenada (STME) a retener la suma arriba indicada, mas $3 (pesos tres) en concepto de gastos administrativos, hasta los primeros $100 (Pesos cien) y a partir de allí $3 (pesos tres) por cada $ o fracción de $50 (pesos cincuenta) por igual concepto, de mis haberes mensuales o a debitarlo de mi caja de ahorro en el Banco de la Provincia de Buenos Aires, Sucursal Ensenada, de acuerdo a las fechas de inicio y cierre de cada período el cual se estima ocurrirá en el mes en curso. Se fija la jurisdicción de los Tribunales Ordinarios de la Ciudad de La Plata, con exclusión de todo otro que pudiere corresponder. EN CASO DE NO EFECTUARSE LA RETENCION MENSUAL CORRESPONDIENTE ME COMPROMETO A ABONAR LAS SUMAS ADEUDADAS AL STME O EN SU DEFECTO ACEPTAR LOS CARGOS ADMINISTRATIVOS QUE CORREPONDAN A CADA CIERRE MENSUAL POR LAS SUMAS ADEUDADAS.</span></p>
        <p align="justify">&nbsp;</p>
          <p align="justify"><br />
          </p>
          <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="1" bgcolor="#000000" class="style12"></td>
              <td></td>
              <td bgcolor="#000000" class="style12"></td>
            </tr>
            <tr>
              <td width="45%" class="style12"><div align="left" class="style17"><strong>FIRMA AFILIADO (de puño y letra)</strong></div></td>
              <td width="10%">&nbsp;</td>
              <td width="45%" class="style12"><div align="left" class="style17"><strong>FIRMA GARANTE</strong></div></td>
            </tr>
            <tr>
              <td width="45%" class="style12"><div align="left"><span class="style17"><span class="style18"><span class="style17"><span class="style17"></span></span></span></span></div></td>
              <td width="10%">&nbsp;</td>
              <td width="45%" class="style12"><div align="left"><span class="style17"><span class="style18"><span class="style17"><span class="style17"></span></span></span></span></div></td>
            </tr>
            <tr>
              <td width="45%" class="style12"><div align="left" class="style17">NOMBRE Y APELLIDO:</div></td>
              <td width="10%">&nbsp;</td>
              <td width="45%" class="style12"><div align="left" class="style17">NOMBRE Y APELLIDO:</div></td>
            </tr>
            <tr>
              <td width="45%" class="style12"><div align="left"><span class="style17"><span class="style18"><span class="style17"><span class="style17"></span></span></span></span></div></td>
              <td width="10%">&nbsp;</td>
              <td width="45%" class="style12"><div align="left"><span class="style17"><span class="style18"><span class="style17"><span class="style17"></span></span></span></span></div></td>
            </tr>
            <tr>
              <td width="45%" class="style12"><div align="left" class="style17">LEGAJO MUNICIPAL:</div></td>
              <td width="10%">&nbsp;</td>
              <td width="45%" class="style12"><div align="left" class="style17">LEGAJO MUNICIPAL:</div></td>
            </tr>
            <tr>
              <td width="45%" class="style12"><div align="left"><span class="style17"><span class="style18"><span class="style17"><span class="style17"></span></span></span></span></div></td>
              <td width="10%">&nbsp;</td>
              <td width="45%" class="style12"><div align="left"><span class="style17"><span class="style18"><span class="style17"><span class="style17"></span></span></span></span></div></td>
            </tr>
            <tr>
              <td width="45%" class="style12"><div align="left" class="style17">DIRECCION:</div></td>
              <td width="10%">&nbsp;</td>
              <td width="45%" class="style12"><div align="left" class="style17">DIRECCION:</div></td>
            </tr>
          </table>
      </td>
    </tr>
  </table>
<%END IF%>
</div>
</body>
</html>
