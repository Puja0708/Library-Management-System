' Global variable declaration
Private cmd As New ADODB.Command
Dim con As Connection
Dim rs As ADODB.Recordset   




' Setting a connection when the form loads.
Private Sub Form1_Load()
    ' Open a connection using OLE DB tags.
    Set con = New ADODB.Connection
    With con
        .ConnectionTimeout = 3
        .CursorLocation = adUseClient
        .Provider = "MSDAORA"
    End With
    ' Opens a connection to a data source.
    con.Open "user id = scott; password = tiger;"
    ' ADODB.Recordset create a Recordset object.
    Set rs = New ADODB.Recordset
    ' Opens a cursor.
    rs.Open "select * from Book", con, adOpenDynamic, adLockOptimistic
End Sub

'update
Private Sub Command1_Click()
 Dim Bno As String
    Dim cmdString As String
    
    Set rs = New ADODB.Recordset
    Bno = Text1.Text
    ' Create a query for modification record.
    cmdString = "SELECT * from book WHERE B_no = " & "'" & Bno & "'"
    rs.Open cmdString, con, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        MsgBox "Book no " & Bno & " not found "
    Else
        ' Transferring the field contents to text boxes
        Text1.Text = rs!B_no
        Text2.Text = rs!ISBN
        Text3.Text = rs!Subject
            Text4.Text = rs!Name
        Text5.Text = rs!Author
        Text6.Text = rs!Publisher
         Text7.Text = rs!Editor
        Text8.Text = rs!Copies
         Text9.Text = rs!Cost
    End If

End Sub
'deleting the searched record.
Private Sub Command2_Click()
Dim strMessage As String
  strMessage = "Deletion in progress:" & vbCr & _
        "Data in buffer = " & rs!B_no & ", " & _
        rs!ISBN & " " & vbCr & vbCr
        If MsgBox(strMessage, vbYesNoCancel) = vbYes Then
       rs.Delete
         Text1.Text = " "
            Text2.Text = " "
            Text3.Text = " "
            Text4.Text = " "
            Text5.Text = " "
            Text6.Text = " "
            Text7.Text = " "
            Text8.Text = " "
            Text9.Text = " "
          MsgBox "Data in recordset Deleted"
        Else
            MsgBox "No deletion in Book table."
        End If
 End Sub
'add a new record
Private Sub Command3_Click()
 Dim strMessage As String
    ' Adding a new row to the recordset
    rs.AddNew
    If rs.EditMode = adEditAdd Then
         rs("B_no") = Text1.Text
         rs("ISBN") = Text2.Text
         rs("Subject") = Text3.Text
         rs("Name") = Text4.Text
         rs("Author") = Text5.Text
         rs("Publisher") = Text6.Text
         rs("Editor") = Text7.Text
         rs("Copies") = Text8.Text
         rs("Cost") = Text9.Text
    End If
    strMessage = "AddNew in progress:" & vbCr & _
            "Data in buffer = " & rs!B_no & ", " & _
            rs!ISBN & " " & vbCr & vbCr & _
            "Use Update to save buffer to recordset?"
        If MsgBox(strMessage, vbYesNoCancel) = vbYes Then
           rs.Update
              MsgBox "Data in recordset = " & rs!B_no & ", " & Trim(rs!ISBN) & " Added"
        Else
            
            rs.CancelUpdate
            MsgBox "No new record added in BOOK table."
        End If


End Sub
'Search
Private Sub Command4_Click()
 Dim Bno As String
    Dim cmdString As String
    Set rs = New ADODB.Recordset
    Bno = txtTno.Text
    cmdString = "SELECT * from Book WHERE B_no = " & "'" & Bno & "'"
  rs.Open cmdString, con, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        MsgBox "Book no " & Bno & " not found "
          Else
      Text1.Text = rs!B_no
        Text2.Text = rs!ISBN
        Text3.Text = rs!Subject
            Text4.Text = rs!Name
        Text5.Text = rs!Author
        Text6.Text = rs!Publisher
         Text7.Text = rs!Editor
        Text8.Text = rs!Copies
         Text9.Text = rs!Cost
    End If
End Sub
'refresh
Private Sub Command5_Click()
            Text1.Text = " "
            Text2.Text = " "
            Text3.Text = " "
            Text4.Text = " "
            Text5.Text = " "
            Text6.Text = ""
            Text7.Text = " "
            Text8.Text = " "
            Text9.Text = ""
           End Sub
'exit
Private Sub Command6_Click()
rs.Close
con.Close
End
End Sub
'home
Private Sub Command7_Click()
Form5.Show
End Sub














' Setting a connection when the form loads.
Private Sub Form2_Load()
    ' Open a connection using OLE DB tags.
    Set con = New ADODB.Connection
    With con
        .ConnectionTimeout = 3
        .CursorLocation = adUseClient
        .Provider = "MSDAORA"
    End With
    ' Opens a connection to a data source.
    con.Open "user id = scott; password = tiger;"
    ' ADODB.Recordset create a Recordset object.
    Set rs = New ADODB.Recordset
    ' Opens a cursor.
    rs.Open "select * from User", con, adOpenDynamic, adLockOptimistic
End Sub
‘add
Private Sub Cmdadd_Click()
Dim strMessage As String
    ' Adding a new row to the recordset
    rs.AddNew
    If rs.EditMode = adEditAdd Then
         rs("id") = Text1.Text
         rs("Roll_no") = Text2.Text
         rs("Name") = Text3.Text
      rs("Branch") = Text4.Text
    End If
        ' Show contents of buffer and get user input.
        strMessage = "AddNew in progress:" & vbCr & _
            "Data in buffer = " & rs!id & ", " & _
            rs!Roll_no & " " & vbCr & vbCr & _
            "Use Update to save buffer to recordset?"
        If MsgBox(strMessage, vbYesNoCancel) = vbYes Then
            ' Saves the contents of the copy buffer row.
            rs.Update
            ' Go to the new record and show the resulting data.
            MsgBox "Data in recordset = " & rs!id & ", " & Trim(rs!Roll_no) & " Added"
        Else
            ' Cancels any changes made to the current record.
            rs.CancelUpdate
            MsgBox "No new record added in User table."
        End If
End Sub
‘refresh
Private Sub cmdrefresh_Click()
            Text1.Text = " "
            Text2.Text = " "
            Text3.Text = " "
            Text4.Text = " "
      End Sub
‘delete
Private Sub Cmddelete_Click()
Dim strMessage As String
    ' Show contents of buffer and get user input.
    strMessage = "Deletion in progress:" & vbCr & _
        "Data in buffer = " & rs!id & ", " & _
        rs!Roll_no & " " & vbCr & vbCr
        If MsgBox(strMessage, vbYesNoCancel) = vbYes Then
            ' Delete the record.
            rs.Delete
             Text1.Text = " "
            Text2.Text = " "
            Text3.Text = " "
            Text4.Text = " "
           ' Go to the new record and show the resulting data.
            MsgBox "Data in recordset Deleted"
        Else
            MsgBox "No deletion in User table."
        End If
End Sub
‘exit
Private Sub cmdexit_Click()
rs.Close
con.Close
End
End Sub
‘search
Private Sub Cmdsearch_Click()
 Dim i_d As String
    Dim cmdString As String
    ' ADODB.Recordset create a Recordset object.
    Set rs = New ADODB.Recordset
    i_d = txtTno.Text ' Enter the book no. to delete.
    ' Create a query for deleting a record.
    cmdString = "SELECT * from User WHERE id = " & "'" & i_d & "'"
    ' Opens a cursor for deletion record
    rs.Open cmdString, con, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        MsgBox "Identity no " & i_d & " not found "
        Text1.Text = " "
    Else
        ' Transferring the field contents to text boxes
         Text1.Text = rs!id
        Text2.Text = rs!Roll_no
        Text3.Text = rs!Name
            Text4.Text = rs!Branch
    End If
End Sub

‘update
Private Sub cmdupdate_Click()
Dim i_d As String
    Dim cmdString As String
    ' ADODB.Recordset create a Recordset object.
    Set rs = New ADODB.Recordset
    i_d = Text1.Text  ' Enter the identity no. to modify.
    ' Create a query for modification record.
    cmdString = "SELECT * from user WHERE id = " & "'" & i_d & "'"
    rs.Open cmdString, con, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        MsgBox "identity " & i_d & " not found "
    Else
        ' Transferring the field contents to text boxes
        Text1.Text = rs!id
        Text2.Text = rs!Roll_no
        Text3.Text = rs!Name
            Text4.Text = rs!Branch
        End If
End Sub

Private Sub Command1_Click()
Form5.Show
End Sub





' Setting a connection when the form loads.
Private Sub Form1_Load()
    ' Open a connection using OLE DB tags.
    Set con = New ADODB.Connection
    With con
        .ConnectionTimeout = 3
        .CursorLocation = adUseClient
        .Provider = "MSDAORA"
    End With
    ' Opens a connection to a data source.
    con.Open "user id = scott; password = tiger;"
    ' ADODB.Recordset create a Recordset object.
    Set rs = New ADODB.Recordset
    ' Opens a cursor.
    rs.Open "select * from Issue", con, adOpenDynamic, adLockOptimistic
End Sub

‘add
Private Sub Cmdadd_Click()
 Dim strMessage As String
    ' Adding a new row to the recordset
    rs.AddNew
    If rs.EditMode = adEditAdd Then
         rs("Bno") = Text1.Text
         rs("id") = Text2.Text
         rs("Issue_date") = Text3.Text
         rs("Due_date") = Text4.Text
         rs("Copies_available") = Text5.Text
    End If
        ' Show contents of buffer and get user input.
        strMessage = "AddNew in progress:" & vbCr & _
            "Data in buffer = " & rs!Bno & ", " & _
            rs!id & " " & vbCr & vbCr & _
            "Use Update to save buffer to recordset?"
        If MsgBox(strMessage, vbYesNoCancel) = vbYes Then
            ' Saves the contents of the copy buffer row.
            rs.Update
            ' Go to the new record and show the resulting data.
            MsgBox "Data in recordset = " & rs!Bno & ", " & Trim(rs!id) & " Added"
        Else
            ' Cancels any changes made to the current record.
            rs.CancelUpdate
            MsgBox "No new record added in Issue table."
        End If
End Sub
‘delete
Private Sub cmddel_Click()
Dim strMessage As String
    ' Show contents of buffer and get user input.
    strMessage = "Deletion in progress:" & vbCr & _
        "Data in buffer = " & rs!Bno & ", " & _
        rs!id & " " & vbCr & vbCr
        If MsgBox(strMessage, vbYesNoCancel) = vbYes Then
            ' Delete the record.
            rs.Delete
            Text1.Text = " "
            Text2.Text = " "
            Text3.Text = " "
            Combo1.Text = " "
            Combo2.Text = " "
           Combo3.Text = " "
            Combo4.Text = " "
          Combo5.Text = " "
           Combo6.Text = " "
            ' Go to the new record and show the resulting data.
            MsgBox "Data in recordset Deleted"
        Else
            MsgBox "No deletion in Issue table."
        End If
End Sub

‘exit
Private Sub cmdexit_Click()
rs.Close
con.Close
End
End Sub
‘refresh
Private Sub cmdrefresh_Click()
           Text1.Text = " "
            Text2.Text = " "
            Text3.Text = " "
            Combo1.Text = " "
            Combo2.Text = " " 
          Combo3.Text = " "
            Combo4.Text = " " 
           Combo5.Text = " "
            Combo6.Text = " "
           End Sub

‘search
Private Sub Cmdsearch_Click()
 Dim B_no As String
    Dim cmdString As String
    ' ADODB.Recordset create a Recordset object.
    Set rs = New ADODB.Recordset
    B_no = txtTno.Text ' Enter the book no. to delete.
    ' Create a query for deleting a record.
    cmdString = "SELECT * from Issue WHERE Bno = " & "'" & B_no & "'"
    ' Opens a cursor for deletion record
    rs.Open cmdString, con, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        MsgBox "Book no " & B_no & " not found "
        Text1.Text = " "
    Else
        ' Transferring the field contents to text boxes
         Text1.Text = rs!Bno
        Text2.Text = rs!id
        Text3.Text = rs!Issue_date
            Text4.Text = rs!Due_date
        Text3.Text = rs!Copies_available
        Combo1.Text = rs!Issue_date
        Combo2.Text = rs!issue_month
        Combo4.Text = rs!Due_date
        Combo5.Text = rs!Due_month
            End If
End Sub
‘update
Private Sub cmdupdate_Click()
Dim B_no As String
    Dim cmdString As String
    ' ADODB.Recordset create a Recordset object.
    Set rs = New ADODB.Recordset
    B_no = Text1.Text  ' Enter the book no. to modify.
    ' Create a query for modification record.
    cmdString = "SELECT * from issue WHERE Bno = " & "'" & B_no & "'"
    rs.Open cmdString, con, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        MsgBox "Book no " & B_no & " not found "
    Else
        ' Transferring the field contents to text boxes
        Text1.Text = rs!Bno
        Text2.Text = rs!id
        Combo1.Text = rs!Issue_date
       Combo2.Text = rs!Issue_month
       Combo3.Text = rs!Issue_year
       Combo4.Text = rs!Due_date
      Combo5.Text = rs!Due_month
     Combo6.Text = rs!Due_year
        Text3.Text = rs!Copies_available
        End If
End Sub

Private Sub Comboddate_dropdown()
Dim j As Integer
For j = 1 To 31
    Comboddate.AddItem j
    Next j
End Sub

Private Sub Combodyear_dropdown()
Dim i As Integer
For i = 2000 To 2050
    Combodyear.AddItem i
    Next i
End Sub

Private Sub Comboidate_dropdown()
Dim i As Integer
For i = 1 To 31
    Comboidate.AddItem 
   Next i
End Sub

Private Sub Comboiyear_dropdown()
Dim i As Integer
For i = 2000 To 2050
    Comboiyear.AddItem i
    Next i
End Sub

Private Sub Command1_Click()
Form5.Show
End Sub


' Setting a connection when the form loads.
Private Sub Form4_Load()
    ' Open a connection using OLE DB tags.
    Set con = New ADODB.Connection
    With con
        .ConnectionTimeout = 3
        .CursorLocation = adUseClient
        .Provider = "MSDAORA"
    End With
    ' Opens a connection to a data source.
    con.Open "user id = scott; password = tiger;"
    ' ADODB.Recordset create a Recordset object.
    Set rs = New ADODB.Recordset
    ' Opens a cursor.
    rs.Open "select * from Issue_return", con, adOpenDynamic, adLockOptimistic
End Sub
Private Sub Combo1_dropdown()
Dim i As Integer
For i = 1 To 31
    Combo1.AddItem i
    Next i
End Sub
    
Private Sub Combo3_dropdown()
Dim i As Integer
For i = 2000 To 2050
    Combo3.AddItem i
    Next i
End Sub

Private Sub Combo4_dropdown()
Dim i As Integer
For i = 1 To 31
    Combo4.AddItem i
    Next i
End Sub

Private Sub Combo6_dropdown()
Dim i As Integer
For i = 2000 To 2050
    Combo6.AddItem i
    Next i
End Sub

Private Sub Combo7_dropdown()
Dim i As Integer
For i = 1 To 31
    Combo7.AddItem i
    Next i
End Sub

Private Sub Combo9_dropdown()
Dim i As Integer
For i = 2000 To 2050
    Combo9.AddItem i
    Next i
End Sub

‘update
Private Sub Command1_Click()
Dim Bno As String
    Dim cmdString As String
    ' ADODB.Recordset create a Recordset object.
    Set rs = New ADODB.Recordset
    Bno = Text1.Text  ' Enter the teacher no. to modify.
    ' Create a query for modification record.
    cmdString = "SELECT * from Issue_return WHERE B_no = " & "'" & Bno & "'"
    rs.Open cmdString, con, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        MsgBox "Book no " & Bno & " not found "
    Else
        ' Transferring the field contents to text boxes
        Text1.Text = rs!B_no
        Text2.Text = rs!id
        Text3.Text = rs!Issue_date
            Text4.Text = rs!Due_date
        Text5.Text = rs!Return_date
        Text6.Text = rs!Fine
         Text7.Text = rs!Copies_available
        End If
End Sub

‘delete
Private Sub Command2_Click()
Dim strMessage As String
    ' Show contents of buffer and get user input.
    strMessage = "Deletion in progress:" & vbCr & _
        "Data in buffer = " & rs!B_no & ", " & _
        rs!id & " " & vbCr & vbCr
        If MsgBox(strMessage, vbYesNoCancel) = vbYes Then
            ' Delete the record.
            rs.Delete
            Text1.Text = " "
            Text2.Text = " "
            Text3.Text = " "
            Text4.Text = " "
            Combo1.Text = " "
            Combo2.Text = " "
            Combo3.Text = " "
            Combo4.Text = " "
            Combo5.Text = " "
             Combo6.Text = " "
            Combo7.Text = " "
            Combo8.Text = " "
            Combo9.Text = " "
            ' Go to the new record and show the resulting data.
            MsgBox "Data in recordset Deleted"
        Else
            MsgBox "No deletion in Issue_return table."
        End If
End Sub
‘add
Private Sub Command3_Click()
 Dim strMessage As String
    ' Adding a new row to the recordset
    rs.AddNew
    If rs.EditMode = adEditAdd Then
         rs("B_no") = Text1.Text
         rs("id") = Text2.Text
         rs("Issue_date") = Combo1.Text
         rs("Issue_month") = Combo2.Text
         rs("Due_date") = Combo4.Text
         rs("Due_month") = Combo5.Text
         rs("Return_date") = Combo7.Text
         rs("Return_month") = Combo8.Text
         rs("Copies_available") = Text3.Text
         rs("Fine") = Text4.Text
          End If
        ' Show contents of buffer and get user input.
        strMessage = "AddNew in progress:" & vbCr & _
            "Data in buffer = " & rs!B_no & ", " & _
            rs!id & " " & vbCr & vbCr & _
            "Use Update to save buffer to recordset?"
        If MsgBox(strMessage, vbYesNoCancel) = vbYes Then
            ' Saves the contents of the copy buffer row.
            rs.Update
            ' Go to the new record and show the resulting data.
            MsgBox "Data in recordset = " & rs!B_no & ", " & Trim(rs!id) & " Added"
        Else
            ' Cancels any changes made to the current record.
            rs.CancelUpdate
            MsgBox "No new record added in Issue_return table."
        End If
End Sub

‘search
Private Sub Command4_Click()
 Dim Bno As String
    Dim cmdString As String
    ' ADODB.Recordset create a Recordset object.
    Set rs = New ADODB.Recordset
    Bno = txtTno.Text ' Enter the book no. to delete.
    ' Create a query for deleting a record.
    cmdString = "SELECT * from Book WHERE B_no = " & "'" & Bno & "'"
    ' Opens a cursor for deletion record
    rs.Open cmdString, con, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        MsgBox "Book no " & Bno & " not found "
        Text1.Text = " "
    Else
        ' Transferring the field contents to text boxes
         Text1.Text = rs!B_no
        Text2.Text = rs!id
        Combo1.Text = rs!Issue_date
        Combo2.Text = rs!issue_month
        Combo4.Text = rs!Due_date
        Combo5.Text = rs!Due_month
        Combo7.Text = rs!Return_date
        Combo8.Text = rs!Return_month
        Text4.Text = rs!Fine
         Text3.Text = rs!Copies_available
    End If
End Sub
‘refresh
Private Sub Command5_Click()
            Text1.Text = " "
            Text2.Text = " "
            Text3.Text = " "
            Text4.Text = " "
            Combo1.Text = " "
            Combo2.Text = " "
            Combo3.Text = " "
            Combo4.Text = " "
            Combo5.Text = " "
             Combo6.Text = " "
            Combo7.Text = " "
            Combo8.Text = " "
            Combo9.Text = " "
         End Sub
‘exit
Private Sub Command6_Click()
rs.Close
con.Close
End
End Sub

Private Sub Command7_Click()
Form5.Show
End Sub


Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub
Private Sub Command3_Click()
Form3.Show
End Sub
Private Sub Command4_Click()
Form4.Show
End Sub




