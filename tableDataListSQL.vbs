Dim oConn
Dim oRS

 Set oConn = CreateObject("ADODB.Connection")

 oConn.ConnectionString ="Provider=SQLOLEDB;Data Source=127.0.0.1,1433;Initial Catalog=XXXXXX;User ID=XXXXXX;Password=XXXXXX"

 oConn.Open 
 
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