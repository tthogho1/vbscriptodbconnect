Dim oConn
Dim oRS

 Set oConn = CreateObject("ADODB.Connection")
 
 oConn.ConnectionString ="Provider=SQLOLEDB;Data Source=127.0.0.1,1433;Initial Catalog=XXXXXXXX;User ID=XXXXXXXX;Password=XXXXXXXX"

 oConn.Open 

 '
 ' https://www.w3schools.com/asp/met_conn_openschema.asp 
 '
 Set oRS = oConn.OpenSchema(4,Array("XXXXXXXX","XXXXXXXX",Empty,Empty))

 while not oRS.EOF
  ' Table Name , Column_Name , Data_Type , length
     'For Each field In oRS.Fields
     'WScript.Echo field.Name & ": " & field.Value
      'Next
       Wscript.Echo  oRS("Table_Name") & ","_  
              & oRS("Column_Name") & ","_ 
              & oRS("Data_Type") & ","_
              & oRS("Character_Maximum_Length") & ","_
              & oRS("IS_NULLABLE") & ","_
              & oRS("COLUMN_DEFAULT") 
              
  oRS.movenext
 Wend
 
 table_name = Wscript.Arguments(0)
 output_file = table_name & ".csv"
 'f_e_pdf_storage
  
 'ファイルシステムオブジェクト作成
Set objFS = CreateObject("Scripting.FileSystemObject")
' ファイルオープン
Set objText = objFS.CreateTextFile(output_file, 2)

'
' adSchemaPrimaryKeys = 28 
'
Set objRecordSet = oConn.OpenSchema(28,Array(Empty,Empty,Empty))

Do Until objRecordset.EOF
   For Each field In objRecordset.Fields
     objtext.WriteLine field.Name & ": " & field.Value
    Next
    objtext.WriteLine
    objRecordset.MoveNext
Loop
 
objText.Close 
 
