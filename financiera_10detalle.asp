<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include FILE="./Includes/cnx.inc" -->
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*210*") =0 then response.redirect("sinacceso.asp")


proveedor=request.form("proveedor")
mes=request.form("mes")
ano=request.form("ano")

mimes=ano & right("00" & mes,2)

if mimes="" then response.redirect("financiera_10.asp")
if proveedor="" then response.redirect("financiera_10.asp")



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
.style37 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
.style42 {
	color: #000000
}
.style47 {color: #000000; font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px;}
.style52 {color: #FF0000; font-family: Geneva, Arial, Helvetica, sans-serif; }
.style53 {	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 14px;
}
.style55 {color: #FF0000; font-weight: bold; }
.style36 {font-family: Geneva, Arial, Helvetica, sans-serif;
	color: #FFFFFF;
	font-size: 12px;
}
-->
</style>
<script src="../../../../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script type="text/javascript">
function format(input){
 var num = input.value.replace(/\,/g,'');
  if(!isNaN(num)){
  } else {
   input.value = input.value.substring(0,input.value.length-1);
  }
}


</script>
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
        <td width="726" height="400" valign="top" bgcolor="#FFFFFF"><br />
          <form id="form1" name="form1" method="post" action="financiera_10detalleSAVE.asp">
            <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="60" bgcolor="#FFFFFF"><div align="center" class="style42"><span class="style15">REGISTRAR PAGO A PROVEEDOR
                  <%

ahoraes=year(date) & right("00" & month(date),2) 

		  
			  
cadena="SELECT     SUM(dbo.INSTRUMENTOS_CUOTAS.valorcuota) AS Expr1"
cadena=cadena & " FROM         dbo.INSTRUMENTOS_CUOTAS INNER JOIN"
cadena=cadena & "                       dbo.INSTRUMENTOS ON dbo.INSTRUMENTOS_CUOTAS.instid = dbo.INSTRUMENTOS.instid"
cadena=cadena & " WHERE     (dbo.INSTRUMENTOS.comercio = " & proveedor & ") AND (dbo.INSTRUMENTOS_CUOTAS.fechacobro = " & mimes & ")"

deuda=0
set rs=conexion.execute(cadena)
if rs.eof=false then
	deuda=rs.fields(0)
end if
rs.close
set rs=nothing
if isnull(deuda)=true then deuda=0
			  
			  %>
                    </span><br />
                </div></td>
              </tr>
              <tr>
                <td bgcolor="#374957"><table width="100%" border="0" cellspacing="1" cellpadding="0">

                    <tr>
                      <td height="40" bgcolor="#e5e5e5"><span class="style47">&nbsp;Pago registrado por:</span></td>
                      <td bgcolor="#FFFFFF" class="style37 style32">&nbsp; <%=session("usuario")%></td>
                    </tr>
                    <tr>
                      <td width="33%" height="40" bgcolor="#e5e5e5"><span class="style47">&nbsp;Monto adeudado para este mes:</span></td>
                      <td width="67%" bgcolor="#FFFFFF" class="style37 style32">&nbsp; $<%=formatnumber(deuda,2)%> </td>
                    </tr>
                    <tr>
                      <td height="40" bgcolor="#e5e5e5"><span class="style47">&nbsp;Cheque número:</span></td>
                      <td bgcolor="#FFFFFF">&nbsp; <input type="text" name="cheque" id="cheque" /> <input name="proveedor" type="hidden" id="proveedor" value="<%=proveedor%>" />
                        <input name="mimes" type="hidden" id="mimes" value="<%=mimes%>" />
                        <input name="deuda" type="hidden" id="deuda" value="<%=deuda%>" /></td>
                    </tr>
                    <tr>
                      <td height="40" bgcolor="#e5e5e5"><span class="style47">&nbsp;Monto cheque:</span></td>
                      <td bgcolor="#FFFFFF">&nbsp; <input name="pago" type="text" id="pago" value="<%=deuda%>" onkeyup="format(this);" /></td>
                    </tr>
                </table></td>
              </tr>
            </table>
            <p align="center">
              <%if deuda >0 then%>
              <span class="style53"><strong>Acción monitoreada. Usuario responsable: </strong><span class="style55"> <%=session("usuario")%></span></span></p>
            <p align="center">
              <input type="submit" name="button" id="button" value="Registrar pago" onclick="javascript:return confirm('Registrar pago a proveedor?')"/>
              <%else%>
              <span class="style52">No existe deuda pendiente para este proveedor en el mes seleccionado</span>
              <%end if%>
            </p>
          </form>
          <p>&nbsp;</p>
          <p>&nbsp;</p></td>
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
