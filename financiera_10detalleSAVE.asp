<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<!-- #include FILE="./Includes/cnx.inc" -->
<%if instr(session("accesos"), "*210*") =0 then response.redirect("sinacceso.asp")


proveedor=request.form("proveedor")
mimes=request.form("mimes")
deuda=request.form("deuda")
cheque=request.form("cheque")
pago=request.form("pago")

ok=true

if proveedor="" then ok=false
if mimes="" then ok=false
if deuda="" then ok=false
if cheque="" then ok=false
if pago="" then ok=false
if isnumeric(pago)=false then ok=false

if ok=true then
	
	
	cadena="select deudaoriginal from comercios_fechapago where cid=" & proveedor & " and fechapago=" & mimes
	set rs=conexion.execute(cadena)
	if rs.eof=false then
		rs.close
		set rs=nothing
		conexion.close
		set conexion=nothing
		response.redirect("mensaje.asp?destino=financiera_10&mensaje=Ya existe un pago registrado para el proveedor correspondiente a la fecha seleccionada")
	end if
	rs.close
	set rs=nothing
	
	
	cadena="insert into comercios_fechapago (cid, fechapago, deudaoriginal, monto, cheque, userid, fpagotexto) values ("
	cadena=cadena & proveedor & ", "
	cadena=cadena & mimes & ", "
	cadena=cadena & deuda & ", "
	cadena=cadena & pago & ", "
	cadena=cadena & "'" & cheque & "', "
	cadena=cadena & session("userid") & ", "
	cadena=cadena & "'" & now() & "')"

	conexion.execute(cadena)
	conexion.close
	set conexion=nothing

	response.redirect("mensaje.asp?destino=financiera_10&mensaje=Pago registrado satisfactoriamente")

end if
%>
<script>javascript:history.go(-1)</script>