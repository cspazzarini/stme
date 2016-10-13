<!-- #include FILE="./Includes/cnx.inc" -->
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")
if instr(session("accesos"), "*26*") =0 then response.redirect("sinacceso.asp")
ierror=0
X=INT(REQUEST.FORM("MAXLINES"))
dim maximo, monto 
FOR T=1 TO X
	MAXIMO=REQUEST.FORM("MAXIMO" & T)
	INSTRUMENTO=REQUEST.FORM("INSTID" & T)
	CUOTAID=REQUEST.FORM("CUOTAID" & T)
	MONTO=REQUEST.FORM("MONTO" & T)
	if isnumeric(monto)=true then
		monto=cdbl(monto)
		maximo=cdbl(maximo)
		response.write(monto & "-" & maximo & "//////")
		IF MONTO > MAXIMO THEN
			ierror=1
		END IF
	end if
NEXT

ahoraes=year(date) & right("00" & month(date),2) & right("00" & day(date),2)

mimes=year(date) & right("00" & month(date),2) 

if ierror=0 then
		
		
		
		FOR T=1 TO X
			INSTRUMENTO=REQUEST.FORM("INSTID" & T)
			CUOTAID=trim(REQUEST.FORM("CUOTAID" & T))
			if cuotaid="" then cuotaid=1
			MONTO=REQUEST.FORM("MONTO" & T)
			if isnumeric(monto)=true then
				monto=cdbl(monto)
				cadena="insert into instrumentos_pagos (instid, cuotaid, monto, userid,  fecha, mesdescuento) values ("
				cadena=cadena & instrumento & ", "
				cadena=cadena & cuotaid & ", "
				cadena=cadena & monto & ", "
				cadena=cadena & session("userid") & ", "
				cadena=cadena & ahoraes & ","
				cadena=cadena & mimes & ")"
				conexion.execute(cadena)
				
				cadena="update instrumentos set restapagar=restapagar - " & monto & " where instid=" & instrumento
				conexion.execute(cadena)
				
				cadena="update instrumentos_cuotas set restapagar=restapagar - " & monto & " where instid=" & instrumento & " and cuotaid=" & cuotaid
				conexion.execute(cadena)
				
			end if
		next
		conexion.close
		set conexion=nothing
		RESPONSE.REDIRECT("done.asp?destino=financiera")
else
			%>
			<script>javascript:alert("El monto ingresado no puede superar el monto de deuda");history.go(-1);</script>
			<%
end if
%>