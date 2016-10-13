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
Session("Valor") = letras
End Function



x=convertir(monto)


set tn = server.createObject("briz.AspThumb")

'*************************************************************
'**********           IMPRESION DE BONOS             *********
'*************************************************************

IF REQUEST("tipo")="BONO" then

		tn.load "D:\IISDATA\intempo\bonovacio.jpg"
		tn.ResizeQuality = 0
		tn.Resize 800, 376
		tn.EncodingQuality = 100
		
		tn.SetFont "Arial", 10, True, False, False, vbblack
		
		a=now()
		
		cbono=year(a) & right("00" & month(a),2) & right("00" & day(a),2) & "*" & id
		control=right("00" & session("userid"),3) & "*" & right("0000000" & id,7)
		pagare="DP*" & right("0000000000000000000" & id,8)

		tn.drawtext 350,160, a

		tn.drawtext 640,160, left("$" & formatnumber(monto,2) & "//////////////////////////////////////////////////////////////////////////////////", 27)


		tn.drawtext 110,211, left(nombre & "/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////",60)
		
		tn.drawtext 540,211, legajo

		tn.drawtext 110,185, left(session("valor") & " con 0/100 centavos///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////",100)

		tn.drawtext 640,211, dni

		tn.SetFont "Arial", 34, True, False, False, vbwhite
		tn.drawtext 550,25, "$" & formatnumber(monto,2) & "**********"


		tn.SetFont "Arial", 10, True, False, False, vbwhite
		tn.drawtext 720,15, "$" & formatnumber(monto,2) 


		tn.SetFont "Arial", 7, True, False, False, vbwhite
		tn.drawtext 290,87, ucase("$" & formatnumber(monto,2) & " * " & session("valor") & " * " & "$" & formatnumber(monto,2) & " * " & session("valor") & " * " & "$" & formatnumber(monto,2) )


		tn.SetFont "Arial", 30, True, False, False, vbwhite
		tn.drawtext 400,290, "$" & formatnumber(monto,2) 

		tn.SetFont "Arial", 20, True, False, False, vbwhite
		tn.drawtext 650,100, "$" & formatnumber(monto,2) 

		tn.SetFont "Arial", 30, True, False, False, vbwhite
		tn.drawtext 20,250, "$" & formatnumber(monto,2) 




		tn.SetFont "Arial", 12, True, False, False, vbwhite
		tn.drawtext 670,350, left("$" & formatnumber(monto,2) & "***************",20)
		
		tn.SetFont "Arial", 10, True, False, False, vbwhite
		tn.drawtext 420,22, cbono

		tn.SetFont "Arial", 10, True, False, False, vbwhite
		tn.drawtext 420,45, control

		tn.SetFont "Arial", 10, True, False, False, vbwhite
		tn.drawtext 420,69, pagare
		
		
		tn.Save "D:\IISDATA\tempfolder\" & trim(id) & "-" & legajo & ".jpg"

		tn.SetFont "Arial", 40, True, False, False, vbblack
		tn.drawtext 1,240, "**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**"

		tn.SetFont "Arial", 20, True, False, False, vbblack
		tn.drawtext 1,290, "**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**"

		tn.SetFont "Arial", 14, True, False, False, vbblack
		tn.drawtext 1,320, "**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**"

		tn.SetFont "Arial", 10, True, False, False, vbblack
		tn.drawtext 1,340, "**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**"


		tn.Save "D:\IISDATA\tempfolder\" & trim(id) & "-" & legajo & "_SPECIMEN.jpg"

end if







IF REQUEST("tipo")="ADELANTO" then

		tn.load "D:\IISDATA\intempo\pagarevacio.jpg"
		tn.ResizeQuality = 0
		tn.Resize 800, 376
		tn.EncodingQuality = 100
		
		tn.SetFont "Arial", 10, True, False, False, vbblack
		
		a=now()
		
		cbono=year(a) & right("00" & month(a),2) & right("00" & day(a),2) & "*" & id
		control=right("00" & session("userid"),3) & "*" & right("0000000" & id,7)
		pagare="PG*" & right("0000000000000000000" & id,8)

		tn.SetFont "Arial", 10, True, False, False, vbwhite
		tn.drawtext 360,77, a

		tn.drawtext 670,74, pagare


		tn.SetFont "Arial", 10, True, False, False, vbblack
		tn.drawtext 700,145, "$" & formatnumber(monto,2)


		tn.drawtext 137,295, nombre
		
		tn.drawtext 137,320, legajo

		tn.drawtext 137,340, DOMICILIO

		tn.drawtext 25,200, "DINERO EN EFECTIVO POR LA SUMA DE $ " & formatnumber(monto,2)

		
		tn.Save "D:\IISDATA\tempfolder\" & trim(id) & "-" & legajo & ".jpg"

end if










'*************************************************************
'**********           IMPRESION DE CREDITOS          *********
'*************************************************************

IF REQUEST("tipo")="CREDITO" then

		tn.load "D:\IISDATA\intempo\creditovacio.jpg"
		tn.ResizeQuality = 0
		tn.Resize 800, 376
		tn.EncodingQuality = 100
		
		tn.SetFont "Arial", 10, True, False, False, vbblack
		
		a=now()
		
		cbono=year(a) & right("00" & month(a),2) & right("00" & day(a),2) & "*" & id
		control=right("00" & session("userid"),3) & "*" & right("0000000" & id,7)
		pagare= cuotas & " CUOTAS"

		tn.drawtext 150,165, COMERCIO & "/////////////////////////////////////////////////////"


		tn.drawtext 150,187, CONCEPTO & "/////////////////////////////////////////////////////"

		tn.SetFont "Arial", 8, True, False, False, vbblack

		tn.drawtext 505,208, CUOTAS  & " CUOTAS"

		tn.SetFont "Arial", 10, True, False, False, vbblack

'		tn.drawtext 640,160, left("$" & formatnumber(monto,2) & "//////////////////////////////////////////////////////////////////////////////////", 27)


		tn.drawtext 240,145, nombre & " (DNI: " & DNI & ") /////////////"
		tn.drawtext 730,145, legajo

		tn.drawtext 150,225, left(session("valor") & " con 0/100 centavos///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////",140)
		'tn.drawtext 110,185, left(session("valor") & " con 0/100 centavos///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////",100)

		'tn.drawtext 640,211, dni

		tn.SetFont "Arial", 34, True, False, False, vbwhite
		tn.drawtext 550,25, "$" & formatnumber(monto,2) & "**********"


		tn.SetFont "Arial", 10, True, False, False, vbwhite
		tn.drawtext 720,15, "$" & formatnumber(monto,2) 


		tn.SetFont "Arial", 7, True, False, False, vbwhite
		tn.drawtext 290,87, ucase("$" & formatnumber(monto,2) & " * " & session("valor") & " * " & "$" & formatnumber(monto,2) & " * " & session("valor") & " * " & "$" & formatnumber(monto,2) & "$" & formatnumber(monto,2) & " * " & session("valor") & " * " & "$" & formatnumber(monto,2) & " * " & session("valor") & " * ")


		tn.SetFont "Arial", 30, True, False, False, vbwhite
		tn.drawtext 400,290, "$" & formatnumber(monto,2) 

		tn.SetFont "Arial", 20, True, False, False, vbwhite
		tn.drawtext 650,100, "$" & formatnumber(monto,2) 

		tn.SetFont "Arial", 30, True, False, False, vbwhite
		tn.drawtext 20,250, "$" & formatnumber(monto,2) 




		
		tn.SetFont "Arial", 10, True, False, False, vbwhite
		tn.drawtext 420,22, cbono

		tn.SetFont "Arial", 10, True, False, False, vbwhite
		tn.drawtext 420,45, control

		tn.SetFont "Arial", 10, True, False, False, vbwhite
		tn.drawtext 420,64, pagare
		

		tn.SetFont "Arial", 8, True, False, False, VBWHITE
		tn.drawtext 595,360, cbono


		'detalle de cuotas
		

		Set Conexion = CreateObject("ADODB.Connection")
		conn="Provider=SQLOLEDB;Data source=ARTONEST;Database=stmedata;UID=usuariostme;PWD=Redalert7424"
		conexion.connectionstring = conn
		conexion.open

		cadena="SELECT     TOP (100) PERCENT numerodecuota, valorcuota"
		cadena=cadena & " FROM         dbo.INSTRUMENTOS_CUOTAS"
		cadena=cadena & " WHERE     (instid = " & id & ")"
		cadena=cadena & " ORDER BY numerodecuota"
		
		set rs=conexion.execute(cadena)
		if rs.eof=false then
			dim Vector(48)
			do until rs.eof
				vector(rs.fields(0))=rs.fields(1)
				rs.movenext
			loop
		end if
		FOR T=1 TO 48
			IF VECTOR(T)="" THEN VECTOR(T)="------"
		NEXT
		rs.close
		set rs=nothing
		conexion.close
		set conexion=nothing
		
		tn.SetFont "Arial", 8, True, False, False, vbblack
		posY=232
		posX=415
		for roll=1 to 48
			posy=posy+14
			c=c+1
     		tn.SetFont "Arial", 8, True, False, False, vbblack
			tn.drawtext posX,posy, roll & ")"
     		tn.SetFont "Arial", 8, true, False, False, vbblack
			IF ISNUMERIC(VECTOR(ROLL))=TRUE THEN MONTOC=formatnumber(vector(roll),2) ELSE MONTOC=VECTOR(ROLL)
			tn.drawtext posX+18,posy, "$" & MONTOC
			
			if c=9 then
				posX=posX+64
				posy=232
				c=0
			end if
			
		next
		
		
		tn.Save "D:\IISDATA\tempfolder\" & trim(id) & "-" & legajo & ".jpg"

		tn.SetFont "Arial", 40, True, False, False, vbblack
		tn.drawtext 1,240, "**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**"

		tn.SetFont "Arial", 20, True, False, False, vbblack
		tn.drawtext 1,290, "**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**"

		tn.SetFont "Arial", 14, True, False, False, vbblack
		tn.drawtext 1,320, "**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**"

		tn.SetFont "Arial", 10, True, False, False, vbblack
		tn.drawtext 1,340, "**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**NO VALIDO**"


		tn.Save "D:\IISDATA\tempfolder\" & trim(id) & "-" & legajo & "_SPECIMEN.jpg"

end if
ahoraes=year(date) & month(date) & day(date) & hour(time) & minute(time) & second(time)
session("valor")=""
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>STME - Intranet de gestion</title>
<script>
function makepage(src)
{
  // We break the closing script tag in half to prevent
  // the HTML parser from seeing it as a part of
  // the *main* page.

  return "<html>\n" +
    "<head>\n" +
    "<title>Impresion oficial de documentos STME</title>\n" +
    "<script>\n" +
    "function step1() {\n" +
    "  setTimeout('step2()', 10);\n" +
    "}\n" +
    "function step2() {\n" +
    "  window.print();\n" +
    "  window.close();\n" +
    "}\n" +
    "</scr" + "ipt>\n" +
    "</head>\n" +
    "<body onLoad='step1()'>\n" +
    "<img src='" + src + "'/>\n" +
	"<br>" +
	"<br>" +
	"<br>" +
	"<br>" +
	"<br>" +
	"<br>" +
	"<br>" +
	"<br>" +
	"<br>" +
	"<br>" +
	"<br>" +
	"<br>" +
	"<br>" +
	"<br>" +
    "<img src='" + src + "'/>\n" +
    "</body>\n" +
    "</html>\n";
}

function printme(evt)
{
  if (!evt) {
    // Old IE
    evt = window.event;
  }    
  var image = evt.target;
  if (!image) {
    // Old IE
    image = window.event.srcElement;
  }
  src = image.src;
  link = "about:blank";
  var pw = window.open(link, "_new");
  pw.document.open();
  pw.document.write(makepage(src));
  pw.document.close();
}
</script>
</head>

<body>
<div align="center">
  <p>&nbsp;</p>
  <p><font size="4" face="Geneva, Arial, Helvetica, sans-serif">CLICKEE SOBRE 
    LA IMAGEN PARA IMPRIMIR ESTE DOCUMENTO</font></p>
  <p>&nbsp;</p>
  <p><img src="tempfolder/<%=trim(id) & "-" & legajo & ".jpg?ahoraes=" & ahoraes%>" width="342" height="171" onClick="printme(event);window.close()" /></p>
  
</div>
</body>
</html>
