<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*24*") =0 then response.redirect("sinacceso.asp")%>
 <%
 'creamos el nombre del archivo 
archivo= request.serverVariables("APPL_PHYSICAL_PATH") & "/salud/"&mes&"/281-Trabajadores Ensenada STME.txt" 

'conectamos con el FSO 
set confile = createObject("scripting.filesystemobject") 

'creamos el objeto TextStream 
set fich = confile.CreateTextFile(archivo) 

'escribimos los números del 0 al 9 
for i=0 to 9 
   fich.write(i) 
next 

'cerramos el fichero 
fich.close() 

'volvemos a abrir el fichero para lectura 
set fich = confile.OpenTextFile(archivo) 

'leemos el contenido del fichero 
texto_fichero = fich.readAll() 

'imprimimos en la página el contenido del fichero 
response.write(texto_fichero) 

'cerramos el fichero 
fich.close() 
%> 