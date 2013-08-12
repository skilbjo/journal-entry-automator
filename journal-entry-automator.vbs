


Private Sub OpenAllFiles_Click()
' MsgBox "Into Command Button"
    

   Dim wbResults As Workbook
    Dim wbCodeBook As Workbook
    Dim strPath As String
    Dim strFile As String
        Dim EmpName As String
        Dim DueDate As String
    Dim i As Integer
      Dim x As Integer
      x = 1
      Dim a As Integer
      Dim b As Integer
     Dim StartofData As Integer
    Dim EndofData As Integer
     
     Dim AcctNo As String
     Dim AcctVal As Currency
     Dim SumAcctVal As Currency
     Dim Rfrnce As String
     Dim expensedate As Date
     
     Dim journentry As String
     Dim batchid As String
     
     
     Dim firstname As String
     Dim lastname As String
     
     StartofData = 3
     EndofData = 32

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
     
    Set wbCodeBook = ThisWorkbook
    
    'Define the path to your folder
    ' Change this Path
    
     Sheets("Main Page").Select
     Range("J3").Select
     
     strPath = ActiveCell.Value
     'MsgBox "This is where the files will be picked from " & strPath
     
          Sheets("Main Page").Select
     Range("J4").Select
     
     journentry = ActiveCell.Value
     
          Sheets("Main Page").Select
     Range("J5").Select
     
     batchid = ActiveCell.Value
   
    'Ensure that the path ends in a backslash
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    'Call the first file within the folder (change the file extension, accordingly)
    strFile = Dir(strPath & "*.xls*")
    
    'Loop through each file within the folder
    Do While Len(strFile) > 0
        'Open the current file
        Set wbResults = Workbooks.Open(strPath & strFile)
            

'            Windows(strPath & strFile).Worksheets("Claim Form").Activate

             Windows(strFile).Activate
             
             Workbooks(strFile).Worksheets("Claim Form").Select
             EmpName = Workbooks(strFile).Worksheets("Claim Form").Range("d3").Value
             Rfrnce = Workbooks(strFile).Worksheets("Claim Form").Range("d9").Value 'add if blank, say MiscExpenses
            
            
            
            If Rfrnce = "" Then
                Rfrnce = "Misc Expense"
                End If
             expensedate = Workbooks(strFile).Worksheets("Claim Form").Range("d7").Value
             If expensedate = "12:00:00 AM" Then
             expensedate = Date
             End If
            
            ' MsgBox EmpName
           '  MsgBox InStr(1, EmpName, " ")
             lastname = Right(EmpName, Len(EmpName) - InStr(1, EmpName, " "))
             firstname = Left(EmpName, (InStr(1, EmpName, " ") - 1))
             
             EmpName = lastname & "," & firstname
             
             vendorid = UCase(Left(lastname, 3)) & "001"
             
            ' MsgBox EmpName
             
             'MsgBox CONCATENATE(Right(EmpName, Len(EmpName) - InStr(1, EmpName, " ")), ", ", Left(EmpName, (InStr(1, EmpName, " ") - 1)))
                         
            'DueDate = Workbooks(strFile).Worksheets("Claim Form").Range("c45").Value
            
      
            For i = StartofData To EndofData
                    
             'Workbooks(strFile).Worksheets("Accounting").Select
             
                ColDest = "E" & i
                ColtoMove = "J" & i
    
                'MsgBox ColDest
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Range(ColDest).Select
                'MsgBox Range(ColDest).Value
                
                AcctVal = Workbooks(strFile).Worksheets("Accounting").Range(ColDest).Value
               'Note if D2B, then
                If (AcctVal > 0) Then
                             
                        x = x + 1
                        
                        a = a + 1
                        b = a + 1
                        AcctNo = Workbooks(strFile).Worksheets("Accounting").Range("D" & i).Value
                        
                        
                        Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Activate
                        
                        'Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("a" & x).Value = journentry
                        'Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("c" & x).Value = DueDate
                        
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("a" & x).Value = batchid
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("b" & x).Value = vendorid  'NEED TO DEFINE AS VARIABLE
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("c" & x).Value = EmpName
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("d" & x).Value = "Invoice"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("E" & x).Value = Rfrnce
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("F" & x).Value = expensedate  'NEED TO DEFINE AS VARIABLE
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("G" & x).Value = "" 'currency id
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("H" & x).Value = "" 'rate type id
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("I" & x).Value = "" 'exchange table
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("J" & x).Value = "" 'exchange rate
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("k" & x).Value = "EXP:" & expensedate
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("L" & x).Value = AcctVal 'purchases
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("M" & x).Value = "" '1099
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("N" & x).Value = AcctNo & "-000"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("O" & x).Value = "" 'Description
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("P" & x).Value = "Purch" 'Type
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("Q" & x).Value = AcctVal
                  SumAcctVal = SumAcctVal + AcctVal
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("S" & x).Value = "1" 'doctypecode
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("t" & x).Value = "6"
                        
                 End If
                 
                 AcctVal = Workbooks(strFile).Worksheets("Accounting").Range(ColtoMove).Value
                
                
                'Note if ACC, then
                If (AcctVal > 0) Then
                             
                        x = x + 1
                    
                        
                        a = a + 1
                        b = a + 1
                        AcctNo = Workbooks(strFile).Worksheets("Accounting").Range("I" & i).Value
                        
                        
                        Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Activate
                        
                        'Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("a" & x).Value = journentry
                        'Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("c" & x).Value = DueDate
                        
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("a" & x).Value = batchid
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("b" & x).Value = vendorid  'NEED TO DEFINE AS VARIABLE
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("c" & x).Value = EmpName
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("d" & x).Value = "Invoice"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("E" & x).Value = Rfrnce
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("F" & x).Value = expensedate  'NEED TO DEFINE AS VARIABLE
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("G" & x).Value = "" 'currency id
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("H" & x).Value = "" 'rate type id
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("I" & x).Value = "" 'exchange table
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("J" & x).Value = "" 'exchange rate
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("k" & x).Value = "EXP:" & expensedate
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("L" & x).Value = AcctVal 'purchases
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("M" & x).Value = "" '1099
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("N" & x).Value = AcctNo & "-000"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("O" & x).Value = "" 'Description
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("P" & x).Value = "Purch" 'Type
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("Q" & x).Value = AcctVal
                  SumAcctVal = SumAcctVal + AcctVal
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("S" & x).Value = "1" 'doctypecode
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("T" & x).Value = "6"
                        
                 End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                 
        
             'MsgBox (a)
             
             
             
            Next i
            
            'Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("L" & a).Value = SumAcctVal
            
            'Do
            x = x + 1
                    
           For i = 1 To a
            On Error Resume Next
          '  MsgBox x - i & "abc" & SumAcctVal
                       
            Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("L" & x - i).Value = SumAcctVal
            Next i
           
           a = 0
           
            'Loop
                
            
            
            'Sum what it just did, and make it a credit as an account payable
            
    
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("a" & x).Value = batchid
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("b" & x).Value = vendorid  'NEED TO DEFINE AS VARIABLE
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("c" & x).Value = EmpName
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("d" & x).Value = "Invoice"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("E" & x).Value = Rfrnce
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("F" & x).Value = expensedate  'NEED TO DEFINE AS VARIABLE
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("G" & x).Value = "" 'currency id
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("H" & x).Value = "" 'rate type id
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("I" & x).Value = "" 'exchange table
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("J" & x).Value = "" 'exchange rate
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("k" & x).Value = "EXP:" & expensedate
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("L" & x).Value = SumAcctVal 'purchases
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("M" & x).Value = "" '1099
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("N" & x).Value = "00-2000-000-000"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("O" & x).Value = "" 'Description
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("P" & x).Value = "Pay" 'Type
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("R" & x).Value = SumAcctVal
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("S" & x).Value = "1" 'doctypecode
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("T" & x).Value = "2"
                  
                  SumAcctVal = 0
            
            
            
            
            
           'Close the current file, without saving it
                Windows(strFile).Activate
             
             Workbooks(strFile).Worksheets("Claim Form").Select
            wbResults.Close savechanges:=False
            
            'Sheets("Claim Form").Select
        'Call the next file within the folder
        strFile = Dir
 
    Loop
    
    
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("a" & 1).Value = "BatchID"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("b" & 1).Value = "VendorID"  'NEED TO DEFINE AS VARIABLE
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("c" & 1).Value = "Name"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("d" & 1).Value = "Document Type"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("E" & 1).Value = "Reference"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("F" & 1).Value = "DocDate" 'NEED TO DEFINE AS VARIABLE
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("G" & 1).Value = "Currency ID" 'currency id
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("H" & 1).Value = "Rate Type ID" 'rate type id
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("I" & 1).Value = "Exchange Table ID" 'exchange table
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("J" & 1).Value = "Exchange Rate" 'exchange rate
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("k" & 1).Value = "Document Number"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("L" & 1).Value = "Purchases"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("M" & 1).Value = "1099 Amount"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("N" & 1).Value = "Account"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("O" & 1).Value = "Description"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("P" & 1).Value = "Type"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("Q" & 1).Value = "Debit"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("R" & 1).Value = "Credit"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("S" & 1).Value = "DocTypeCode" 'doctypecode
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("T" & 1).Value = "DistType"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("u" & 1).Value = "Country"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("v" & 1).Value = "Processing Month"
                  Workbooks("Open All files.xlsm").Worksheets("Journal Entry").Range("w" & 1).Value = "Processing Year"
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With
    
    'MsgBox " Done !!!!!!!!!!!! Please check"
End Sub
