<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%'response.redirect("http://www.artonestudio.com/WEBSITE97.ASPX")%>
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
.style16 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
}
.style29 {font-size: 10}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
.style30 {font-family: Arial, Helvetica, sans-serif; font-weight: bold; color: #344453; font-size: 24px; }
.style31 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: #344454; }
.style32 {font-size: 10px}
.style33 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	color: #FFFFFF;
	font-size: 12px;
}
.style34 {
	font-size: 24
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
        <td width="1000" height="8" colspan="2" bgcolor="#2F3863"></td>
      </tr>
      <tr>
        <td height="110" colspan="2"><table width="1000" border="0" cellspacing="0" cellpadding="0">
          <tr bgcolor="#FFFFFF">
            <td height="110" bgcolor="#CC0000"><p align="center" class="style33"><br />
                    <br />
                    <br />
                    <br />
                    <br />
                    <br />
                    <br />
              El acceso no autorizado a este sitio se encuentra prohibido. Su direccion IP esta siendo registrada:</p>              <p align="center" class="style33 style34">
                <%
Dim UserIPAddress
UserIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If UserIPAddress = "" Then
  UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
  response.write(UserIPAddress)
End If
%>
                &nbsp;</p></td>
            </tr>
        </table></td>
        </tr>
      
      <tr>
        <td height="33" colspan="2" background="images/fondo_menu_top.jpg"><table width="123" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="114"><div align="center"><span class="style1"><a href="default.asp">PRINCIPAL</a></span></div></td>
              <td width="9"><div align="center"><img src="images/separador_menu_top.jpg" width="5" height="33" /></div></td>
              </tr>
        </table></td>
      </tr>
      
      <tr>
        <td height="2" colspan="2" bgcolor="#2F3863"></td>
      </tr>
      <tr>
        <td height="480" colspan="2" valign="top" background="images/spacer.jpg" bgcolor="#FFFFFF"><p>&nbsp;</p>
          <p align="center" class="style15">BIENVENIDO A LA INTRANET DE GESTION DEL </p>
          <p align="center" class="style30">SINDICATO DE TRABAJADORES MUNICIPALES DE ENSENADA</p>
          <p align="center" class="style31">Ingrese su nombre de usuario y password, le recordamos que el acceso a esta área   sólo esta permitida a personal administrativo del STME.</p>
          <form id="form1" name="form1" method="post" action="loginprocess.asp">
            <table width="400" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td bgcolor="#2F3863"><table width="400" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td bgcolor="#FFFFFF"><table width="400" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr>
                        <td width="128"><div align="right"><span class="style16">Usuario:</span></div></td>
                        <td width="10">&nbsp;</td>
                        <td width="260"><input name="f1" type="text" class="style16" id="f1" size="30" maxlength="30" /></td>
                      </tr>
                      <tr>
                        <td><div align="right"><span class="style29"></span></div></td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr>
                        <td><div align="right"><span class="style16">Contraseña:</span></div></td>
                        <td>&nbsp;</td>
                        <td><input name="f2" type="password" class="style16" id="f2" size="30" maxlength="30" /></td>
                      </tr>
                      <tr>
                        <td><div align="right"><span class="style29"></span></div></td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr>
                        <td><div align="right"><span class="style16">PIN:</span></div></td>
                        <td>&nbsp;</td>
                        <td><input name="f3" type="password" class="style16" id="f3" size="30" maxlength="30" /></td>
                      </tr>
                      <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                      
                      <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td><input type="submit" name="button" id="button" value="Ingresar a intranet" /></td>
                      </tr>
                      <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                    </table></td>
                  </tr>
                </table></td>
              </tr>
            </table>
          </form>          <p align="center" class="style16">&nbsp;</p>
          <p align="center" class="style16"><br />
            <br />
          </p></td>
      </tr>
      <tr>
        <td height="2" colspan="2" bgcolor="#2F3863"></td>
      </tr>      
    </table></td>
    <td width="6" background="images/derecha.jpg">&nbsp;</td>
  </tr>
</table>
</body>
</html>
