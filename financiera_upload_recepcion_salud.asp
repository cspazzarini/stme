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
<script type="application/javascript">
function validateForm() {
  var errors = 0;
  
  /* 
   * Store appropriate form elements.
   * There are many ways to do this, but I'm keeping it 
   * simple for the example
   */
  var firstName = document.getElementById("file").value;
 
  // etc..

  // validate required fields
  if (firstName === "") {
	alert('Debe seleccionar un archivo a procesar');
    errors++;
  }

  // etc...

  return errors === 0;
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
        <td height="400" valign="top" bgcolor="#FFFFFF"><p>&nbsp;</p>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                  <td width="96%" bgcolor="#384858">
 
<form action="upload.asp" method="post" enctype="multipart/form-data" onsubmit="return validateForm()">
     <p class="style36" >Seleccione el archivo recibido del Ministerio de Salud: <br /> <br />
	 <input type="file" name="file" id="file"/><input type="submit" value="Procesar resultados" /></p>
</form>
</td>
</tr>
</table>
		</body> </html>