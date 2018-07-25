<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="PrintDemo.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
  <script language="VBScript">
sub Print()
OLECMDID_PRINT = 6
OLECMDEXECOPT_DONTPROMPTUSER = 2
OLECMDEXECOPT_PROMPTUSER = 1
call WB.ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER,1)
End Sub
document.write "<object id='WB' width='0' height='0' classid='CLSID:8856F961-340A-11D0-A96B-00C04FD705A2'></object>"
</script>
    <title></title>
    <style>
        @media all {
	.page-break	{ display: none; }
}

@media print {
	.page-break	{ display: block; page-break-before: always; }
}
    </style>
</head>
 
<body>
    <object id="WebBrowser1" width="0" height="0" classid="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"> </object>
    <a href="javascript:window.print();">Print</a>
    <div>
  <div>
    Do not print
  </div>
        <div id="printable" style="display:none">
            Print this div default
              <div id="printable1"  class="page-break" style="background-color: pink">
                Print this div 1
              </div>
                    <div id="printable2"  class="page-break" style="background-color: pink">
                Print this div 2
              </div>
        </div>
        <input type="hidden" runat="server" id="userId"/>
        <div id="printerDiv" style="display:none"></div>
  <button onClick="printdiv();">Print Div</button>
</div>
    

    <script>
  function printdiv()
  {
      //var printContents = document.getElementById("printable").innerHTML;

      //var div = document.getElementById("printerDiv");
      //div.innerHTML = '<iframe onload="this.contentWindow.print();">' + printContents + '</iframe>';
  //    var printContents2 = document.getElementById("printable").innerHTML;
  ////var head = document.getElementById("head").innerHTML;
  ////var popupWin = window.open('', '_blank');
  //var popupWin = window.open('', 'blank');
  ////popupWin.document.open();
  //popupWin.document.write('' + '<html>' + '<head></head> <style>@media all {	.page-break	{ display: none; }}@media print { header nav, footer {display: none;}   .page-break{ display: block; page-break-before: always; }a[href]:after { content: none !important; }img[src]:after { content: none !important; }} </style>' + '<body onload="window.print();window.close()">' + printContents2 + '</body>' + '</html>');
  //popupWin.document.close();
 return false;
  };
    </script>
</body>
    </html>

<%--/<form>
//    <h1>Page Title</h1>
//    <p id="printArea">
//<!-- content block -->
//<!-- content block -->
//<div class="page-break">First Page </div>
//<!-- content block -->
//<!-- content block -->
//<div class="page-break">Second Page</div>
//        </p>
//<!-- content block -->
//<!-- content -->
//<input type="button" value="Print Page" onClick="window.print()"> 
//    Print without print dialog box testing........
//    <div style="page-break-after:always;">You can indeed....
//        </div>
//</form>--%>

<%--<script language="VBScript">
// THIS VB SCRIP REMOVES THE PRINT DIALOG BOX AND PRINTS TO YOUR DEFAULT PRINTER
Sub window_onunload()
On Error Resume Next
Set WB = nothing
On Error Goto 0
End Sub

Sub Print()
OLECMDID_PRINT = 6
OLECMDEXECOPT_DONTPROMPTUSER = 2
OLECMDEXECOPT_PROMPTUSER = 1


On Error Resume Next

If DA Then
call WB.ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER,1)

Else
call WB.IOleCommandTarget.Exec(OLECMDID_PRINT ,OLECMDEXECOPT_DONTPROMPTUSER,"","","")

End If

If Err.Number <> 0 Then
If DA Then 
Alert("Nothing Printed :" & err.number & " : " & err.description)
Else
HandleError()
End if
End If
On Error Goto 0
End Sub

If DA Then
wbvers="8856F961-340A-11D0-A96B-00C04FD705A2"
Else
wbvers="EAB22AC3-30C1-11CF-A7EB-0000C05BAE0B"
End If

document.write "<object ID=""WB"" WIDTH=0 HEIGHT=0 CLASSID=""CLSID:"
document.write wbvers & """> </object>"

    --%>

