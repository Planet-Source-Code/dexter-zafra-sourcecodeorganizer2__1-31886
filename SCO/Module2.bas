Attribute VB_Name = "Module2"
Option Explicit
Public Path1 As String

Public Function BackUp()
' sub to Backup the DataBase
MsgBox "Organizer has backUp Your Old Data Base to" & vbCrLf & _
       App.Path & "\DBbackup\sourcebook.mdb"
'Just in case Dir already exist
On Error Resume Next
MkDir Path1 & "DBbackup"

'Incase DataBase File does not exist
On Error GoTo Erro
'Copy the DataBase to the BackUp dir
FileCopy Path1 & "sourcebook.mdb", Path1 & "DBbackup\sourcebook.mdb"

'let user know DataBase has been Backed Up
MsgBox "DataBase Has been Backed Up"

Exit Function

Erro:
'Message incase the DataBase could not be found
MsgBox "Could Not Find DataBase"

End Function





