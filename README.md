<div align="center">

## Compact and Repair Access Database<br/>by Chris Peneguy

</div>

### Description

Compact and Repair

You can use the "Compact and Repair" function in Access from ASP code. The following code is an example of how this can be done. Note that when you decide to "Compact and Repair" your Access database, some autonumbers can be changed. Access makes all autonumbers consecutive.

This code uses one database, but I'm sure the code can easily be changed so that the listbox displays, for example, all the databases in one folder.

### API Declarations

Use as you wish

### Source Code

<%
Const Jet_Conn_Partial = "Provider=Microsoft.Jet.OLEDB.4.0; Data source="
Dim strDatabase, strFolder, strFileName
'#################################################
'# Edit the following two lines
'# Define the full path to where your database is
strFolder = "F:\InetPub\wwwroot\_db\"
'# Enter the name of the database
strDatabase = "YourAccessDatabase.mdb"
'# Stop editing here
'##################################################
Private Sub dbCompact(strDBFileName)
Dim SourceConn
Dim DestConn
Dim oJetEngine
Dim oFSO
SourceConn = Jet_Conn_Partial & strFolder & strDatabase
DestConn = Jet_Conn_Partial & strFolder & "Temp" & strDatabase
Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
Set oJetEngine = Server.CreateObject("JRO.JetEngine")
With oFSO
    If Not .FileExists(strFolder & strDatabase) Then
      Response.Write ("Not Found: " & strFolder & strDatabase)
      Stop
    Else
         If .FileExists(strFolder & "Temp" & strDatabase) Then
            Response.Write ("Something went wrong last time " _
            & "Deleting old database... Please try again")
           .DeleteFile (strFolder & "Temp" & strDatabase)
         End If
   End If
End With
With oJetEngine
.CompactDatabase SourceConn, DestConn
End With
oFSO.DeleteFile strFolder & strDatabase
oFSO.MoveFile strFolder & "Temp" _
& strDatabase, strFolder& strDatabase
Set oFSO = Nothing
Set oJetEngine = Nothing
End Sub
Private Sub dbList()
Dim oFolders
Set oFolders = Server.CreateObject("Scripting.FileSystemObject")
  Response.Write ("<Select Name=""DBFileName"">")
  For Each Item In oFolders.GetFolder(strFolder).Files
  If LCase(Right(Item, 4)) = ".mdb" Then
    Response.Write ("<Option Value=""" & Replace(Item, strFolder, "") _
    & """>" & Replace(Item, strFolder, "") & "</Option>")
  End If
Next
Response.Write ("</Select>")
Set oFolders = Nothing
End Sub
%>
<%
' Compact database and tell the user the database is optimized
Select Case Request.form("cmd")
Case "Compact"
dbCompact Request.form("DBFileName")
Response.Write ("Database " & Request.form("DBFileName") & " is optimized.")
End Select
%>
<p><font size="4">Compact and repair database</font></p>
<form method="POST" action="">
<p><%dbList%><input type="submit" value="Compact" name="cmd"></p>
</form>

