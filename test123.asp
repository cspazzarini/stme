<html>
<head>
<title>Print Image Only</title>
<script>
function makepage(src)
{
  // We break the closing script tag in half to prevent
  // the HTML parser from seeing it as a part of
  // the *main* page.

  return "<html>\n" +
    "<head>\n" +
    "<title>Temporary Printing Window</title>\n" +
    "<script>\n" +
    "function step1() {\n" +
    "  setTimeout('step2()', 10);\n" +
    "}\n" +
    "function step2() {\n" +
    "  window.print();\n" +
    "  window.close();\n" +
    "}\n" +
    "</scr" + "ipt>\n" +
    "</head>\n" +
    "<body onLoad='step1()'>\n" +
    "<img src='" + src + "'/>\n" +
    "</body>\n" +
    "</html>\n";
}

function printme(evt)
{
  if (!evt) {
    // Old IE
    evt = window.event;
  }    
  var image = evt.target;
  if (!image) {
    // Old IE
    image = window.event.srcElement;
  }
  src = image.src;
  link = "about:blank";
  var pw = window.open(link, "_new");
  pw.document.open();
  pw.document.write(makepage(src));
  pw.document.close();
}
</script>
</head>
<body>
<h1>Print Image Only</h1>
When you click on the image below, just the image should print.
And the temporary window used in the process should go away
afterwards, whether you allow the print operation to go ahead
or not.
<p>
<noscript>
<b>You have JavaScript turned off.</b> So this won't work.
That is to be expected.
</noscript>
<p> <img src="tempfolder/130-4008.jpg" width="799" height="376" onClick="printme(event)"/> 
</body>
</html>

