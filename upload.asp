<%
Set fs = Server.CreateObject("Scripting.FileSystemObject")

Set tfolder = fs.GetSpecialFolder(2)
tname = fs.GetTempName

'Declare variables
Dim fileSize
Dim filename
Dim file
Dim fileType
Dim p
Dim newPath
Dim actual
actual = DateAdd("m",-1,Now())
'Assign variables
fileSize       = Request.TotalBytes
fileName       = Request.form("filename")
file           = request.form("file")
fileType       = fs.GetExtensionName(file)
fileOldPath    = tfolder
newPath        = Server.MapPath("./salud/p/")

fs.CopyFile "D:\_Cristian\workspace\stme\stme\salud\soeme.txt", "D:\_Cristian\workspace\stme\stme\salud\p\" & year(actual) & right("0"&month(actual),2) &".txt"

set fs = nothing

response.redirect("financiera_recepcion_salud.asp")
%>