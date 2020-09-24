<div align="center">

## Get Operating System Version


</div>

### Description

A few Lines of code to show you the version of windows you are running.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Venkat Mani](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/venkat-mani.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/venkat-mani-get-operating-system-version__1-23322/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>
<body>
<p>The Api used for this are</p>
<p>&nbsp;</p>
<p><b>Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long<br>
</b></p>
<p><b>Private Type OSVERSIONINFO<br>
    dwOSVersionInfoSize As Long<br>
    dwMajorVersion As Long<br>
    dwMinorVersion As Long<br>
    dwBuildNumber As Long<br>
    dwPlatformId As Long<br>
    szCSDVersion As String * 128   ' Maintenance string for PSS usage<br>
End Type</b></p>
<p>The GetVersion Api just takes one parameter of type OSVERSIONINFO. The
OSVERSIONINFO structure will contain all the details about the OS after
GetVersionApi has been successfully executed. The parameters of&nbsp;
OSVERSIONINFO are&nbsp;</p>
<ul>
 <li>dwMajorVersion which gives info about the major version of the OS .This
  value is&nbsp; 3 for win Nt 3.51, 4 for win95/98/me and win nt4 and it is 5
  for win2k.</li>
 <li>dwMinorVersion ,another parameter to differentiate the OS further .It is 0
  for win 95,10 for win 98 ,98 for win ME,0&nbsp;&nbsp; for win2k ,0 for win
  nt4 and 51 for win nt 3.51</li>
 <li>    dwPlatformId&nbsp; .This is an important parameter which helps in
  further differentiating the varios win OS.It is 1 for win 95/ 98/ME ,and 2
  for win NT&nbsp;&nbsp;</li>
</ul>
<p>Once Declared we can use this in the following way</p>
<p><b>Dim os As OSVERSIONINFO&nbsp;</b></p>
<p><br>
<b>os.dwOSVersionInfoSize = Len(os)&nbsp;</b>&nbsp;&nbsp;&nbsp; 'Assign some
size to store the received information<br>
<br>
<b>Dim m As Long<br>
Dim mv As Long<br>
Dim pd As Long<br>
Dim miv As Long</b><br>
<b>m = GetVersionEx(os)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; '</b>The
actual API call to GetVersionEx<br>
<b>mv = os.dwMajorVersion<br>
pd = os.dwPlatformId<br>
miv = os.dwMinorVersion</b><br>
<b>If pd = 2 Then MsgBox " OS is Windows NT" &amp; mv &amp; "." &amp; miv<br>
If pd = 1 Then<br>
If miv = 10 Then MsgBox " OS is Windows 98 "<br>
If miv = 0 Then MsgBox " OS is Windows 95 "<br>
If miv = 90 Then MsgBox " OS is Windows ME "<br>
End If</b></p>
<p>This can be quite useful if you are making OS specific Applications. Send
your comments to <a href="mailto:venky_dude@yahoo.com">venky_dude@yahoo.com</a>.
.Visit my <a href="http://www.geocities.com/venky_dude/venkwork.htm">homepage</a>
for some cool VBcodes.&nbsp;&nbsp; </p>
<p>&nbsp;</p>
</body>
</html>

