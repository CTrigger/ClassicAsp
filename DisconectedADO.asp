Const adVarChar = 200  'the SQL datatype is varchar

'Create a disconnected recordset
Set rs = CreateObject("ADODB.RECORDSET")
rs.Fields.append "Name", adVarChar, 25

rs.CursorType = adOpenStatic
rs.Open

'include data
rs.AddNew "Name", "Some data"
rs.Update

'Sort the ADO a-z
rs.Sort = "Name"


'here delete duplicated values after sort
rs.MoveFirst 
check = ""
Do Until rs.EOF
    if check = rs.fields("Name").value then
      rs.delete
    else
      check = rs.fields("name").value
    rs.MoveNext
Loop 

