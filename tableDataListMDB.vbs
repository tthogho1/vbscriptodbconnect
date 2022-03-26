Dim oConn
Dim oRS

'
' %windir%\SysWOW64\cscript.exe
'

 Set oConn = CreateObject("ADODB.Connection")
 
 oConn.ConnectionString="Provider=Microsoft.Ace.OLEDB.12.0;Data Source=C:\temp\Database4.mdb"
 oConn.Open 

 '
 ' https://www.w3schools.com/asp/met_conn_openschema.asp 
 ' 
 table_name = Wscript.Arguments(0)
 output_file = table_name & ".csv"
 'f_e_pdf_storage
 
 command = "SELECT * FROM " & table_name
 
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objText = objFS.CreateTextFile(output_file, 2)

Set rs = CreateObject("ADODB.Recordset")
rs.Open command,oConn
 
Do Until rs.EOF
	line=""
	i = 0
	For Each field In rs.Fields
	   if ( i = 0 ) Then
	      line = field.Value
	      i = 1
	   else 
	      line = line & ","  & field.Value
	   end if
	Next 
	
	WScript.Echo line
	objtext.WriteLine(line)
	i = 0
	rs.MoveNext
Loop

objText.Close
