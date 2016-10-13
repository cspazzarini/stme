<%
PIN=REQUEST.FORM("PIN")
IF PIN <> SESSION("PIN") THEN RESPONSE.REDIRECT("financiera.asp") else response.redirect(request("destino") & ".asp")
%>