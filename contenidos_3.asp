<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*1*") =0 then response.redirect("sinacceso.asp")%>
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
.style33 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 12px;
}
.style35 {font-family: Arial, Helvetica, sans-serif; font-weight: bold; color: #FFFFFF; }
.style36 {font-family: Geneva, Arial, Helvetica, sans-serif;
	color: #FFFFFF;
	font-size: 12px;
}
-->
</style>



<script language="JavaScript" type="text/javascript">
<!--

function checkform ( Form1 )
{
  // ** START **
  
  if (Form1.F1.value == "") {
    alert( "Ingrese todos los campos requeridos" );
    Form1.F1.focus();
    return false ;
  }
  
  if (Form1.F2.value == "") {
    alert( "Ingrese todos los campos requeridos" );
    Form1.F2.focus();
    return false ;
  }  


  if (Form1.F3.value == "") {
    alert( "Ingrese todos los campos requeridos" );
    Form1.F3.focus();
    return false ;
  }

  if (Form1.F4.value == "") {
    alert( "Ingrese todos los campos requeridos" );
    Form1.F4.focus();
    return false ;
  }

  // ** END **
  return true ;
}
//-->
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
        <td height="400" valign="top" bgcolor="#FFFFFF"><p>&nbsp;</p>
          <form id="form1" name="form1" method="post" action="contenidos_3_alta.asp" onsubmit="return checkform(this);">
            <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="30" bgcolor="#384858"><div align="center"><span class="style35">GESTION DE CONTENIDO - ALTA PROVEEDOR</span></div></td>
              </tr>
              <tr>
                <td bgcolor="#374957"><table width="500" border="0" cellspacing="1" cellpadding="0">
                    <tr>
                      <td bgcolor="#FFFFFF"><table width="500" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="13" height="40">&nbsp;</td>
                            <td width="126" class="style33">Nombre Comercio</td>
                            <td width="361"><input name="F1" type="text" id="F1" size="50" maxlength="50" /></td>
                          </tr>
                          <tr>
                            <td height="40">&nbsp;</td>
                            <td class="style33">Direccion</td>
                            <td><input name="F2" type="text" id="F2" size="50" maxlength="50" /></td>
                          </tr>
                          <tr>
                            <td height="40">&nbsp;</td>
                            <td class="style33">Telefono</td>
                            <td><input name="F3" type="text" id="F3" size="50" maxlength="50" /></td>
                          </tr>
                          <tr>
                            <td height="40">&nbsp;</td>
                            <td class="style33">Email</td>
                            <td><input name="F4" type="text" id="F4" size="50" maxlength="50" /></td>
                          </tr>
                      </table></td>
                    </tr>
                </table></td>
              </tr>
            </table>
                    <p align="center">
                      <input type="submit" name="button" id="button" value="Dar de Alta" />
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
