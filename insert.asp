<%
 Option Explicit
 Dim con 'for connection object
 Dim rec ' for Recordset object 
 Dim nm '  to hold username from submitted data
 Dim rl ' to hold roll from submitted data

 ' create  a Connection Object
  Set con =  Server.createObject("ADODB.Connection")
 ' create a recordset Object 
  Set rec = Server.createObject("Adodb.recordset")

  'open the connection 
  con.Open "Provider=SQLOLEDB; Data Source = (local); Initial Catalog = newpro; User Id = Joy; Password=1234"

   
  'collect data from the submitted form data 
    nm =  Request.form("name")
    rl = Request.form("roll")

    ' Execute an SQL for insertion 
    con.execute("insert into student values(" & rl &",'"& nm & "')")





Response.write("Connection successfull")




%>