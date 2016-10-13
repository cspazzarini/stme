<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*2C*") =0 then response.redirect("sinacceso.asp")%>

<!-- #include FILE="./Includes/cnx.inc" -->
<%if isnumeric(request.form("NUMDOC"))=TRUE THEN


		cadena="delete from instrumentos where instid=" & request.form("NUMDOC")
		conexion.execute(cadena)
		
		cadena="delete from instrumentos_cuotas where  instid=" & request.form("NUMDOC") 
		conexion.execute(cadena)
	conexion.close
	set conexion=nothing
	response.redirect("done.asp?destino=financiera")
END IF%>
<script>javascript:history.go(-1)</script>