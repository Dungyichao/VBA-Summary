# VBA-Summary

## 1. SQL Connection

Reference: 'https://www.access-programmers.co.uk/forums/threads/how-to-make-an-ado-connection-public.167811/
```VBA
'Tools > References > Check the checkbox in front of "Microsoft ActiveX Data Objects 2.5 Library"
Dim Conn1 As ADODB.Connection
Dim Cmd1 As ADODB.Command
Dim Param1 As ADODB.Parameter
Dim Rs1 As ADODB.Recordset


Private mcnn As ADODB.Connection
'https://www.access-programmers.co.uk/forums/threads/how-to-make-an-ado-connection-public.167811/
Function fGetConn() As ADODB.Connection
On Error Resume Next
 
    Dim Server_Name As String
    Dim Database_Name As String
    Dim User_ID As String
    Dim Password As String
    Dim fConnectionStr As String

    Server_Name = "172.16.254.99" ' Enter your server name here
    Database_Name = "DCSPOY1" ' Enter your database name here
    User_ID = "sa" ' enter your user ID here
    Password = "" ' Enter your password here

    fConnectionStr = "Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
    ";Uid=" & User_ID & ";Pwd=" & Password & ";"
    
    
    
    If mcnn Is Nothing Or mcnn.Status = 0 Then
        Set mcnn = New ADODB.Connection
        mcnn.Open fConnectionStr 'Which returns whatever connection string you use
    End If
    Set fGetConn = mcnn
 
End Function
 
Sub CloseConn()
 
    mcnn.Close
    Set mcnn = Nothing
 
End Sub
```

How to call the function
```VBA
Sub Button1_Click()
    Dim result_count As Integer
    Set Conn1 = fGetConn
    result_count = Record_Exist("2A", "01", "1", "2V34Q")
    MsgBox "Record: " & result_count
    Call CloseConn
    
End Sub

Function Record_Exist(LINE_ID As String, WINDER As String, END_NO As String, LOT_NO As String) As Integer
 Set Rs1 = Conn1.Execute("SELECT * FROM [DCSPOY1].[dbo].[Dynafil]") 
 If Rs1.RecordCount < 0 Then ' Evaluate argument.
  Exit Function ' Exit to calling procedure.
 Else
  Record_Exist = Rs1.RecordCount
 End If
End Function
```
