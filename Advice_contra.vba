Dim advice_contra As String
Dim vCusAcc As String
Dim IndicativeratesFile As String
Dim wksheetName As String
Dim vStr1 As String
Dim vStr2 As String
Dim advicecontra_folder As String
Dim vPathSaveFile As String
Sub fill_advice_contra()
advice_contra = ThisWorkbook.Sheets("Setup").Range("C6")
REGISTERfile = ThisWorkbook.Sheets("Setup").Range("E4")
IndicativeratesFile = ThisWorkbook.Sheets("Setup").Range("E5")
vCommissionFile = ThisWorkbook.Sheets("Setup").Range("E7")
advicecontra_folder = ThisWorkbook.Sheets("Setup").Range("C8")

Dim Filename As String
Filename = advice_contra

'Open excel file based on path
'If path is not found display the message "File not found"
If Dir(Filename) = "" Then

    MsgBox "File not found " & Filename

Exit Sub

End If

'Open file advice_contra
Set advice_contra_FILE = Workbooks.Open(Filename, UpdateLinks:=False) 'Remove update popin window

'Create folder for current month
'If Len(Dir(MMYYYYfolder, vbDirectory)) = 0 Then
'   MkDir MMYYYYfolder
'End If

'Get today's date for Advice Entries - Sheet Setup
ThisWorkbook.Sheets("Setup").Activate
Dim vDateEntries As String
vDateEntries = ThisWorkbook.Sheets("Setup").Range("H3").Value
    
'Activate register file
Workbooks(REGISTERfile).Activate

WS_Count = ActiveWorkbook.Worksheets.Count

' Begin the loop.
For xCount = 1 To WS_Count
    
    '''''''''''''''''''Insert Buying rate (Used for Contra entry) - GBP,USD,EUR(if MUR leave it blank)'''''''''''''''''''
    Workbooks(REGISTERfile).Activate
    Worksheets(xCount).Select

    If Worksheets(xCount).Name <> "MUR" Then
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''Get data from incentive rates file and set in BILL DISCOUNTED REGISTER file'''''''''''''''''''''
        
        
        ThisWorkbook.Sheets("Setup").Activate
        Dim vRateLastrow1 As Long
        vRateLastrow1 = Range("Q" & Rows.Count).End(xlUp).Row
        
        For av = 2 To vRateLastrow1
            If Cells(av, 18).Value = Workbooks(REGISTERfile).Worksheets(xCount).Name Then
            ThisWorkbook.Sheets("Setup").Activate
            Dim vStr1 As String
            Dim vStr2 As String
            Dim vBuyingTT1 As String
            Dim vSheetCurr1 As String
            Dim TT1 As Long
        
            'From sheet Setup in column Q get Buying TT currency
            vBuyingTT1 = Cells(av, 17).Value
            
            
            Workbooks(IndicativeratesFile).Activate
            Sheets("RATE0104").Select
            
        
            LastRowTT1 = Cells(Rows.Count, 2).End(xlUp).Row
            
            'Loop to find currency to get T.T amount
            For TT1 = 1 To LastRowTT1
                If Cells(TT1, 2).Value = vBuyingTT1 And Not IsEmpty(Cells(TT1, 2).Value) Then
                    
                   vStr1 = Cells(TT1, 5).Value
                   vStr2 = Cells(TT1, 8).Value
                    
                    Exit For
                    
                End If
                
            Next TT1
            
            'Paste in Setup sheet
            ThisWorkbook.Sheets("Setup").Activate
            Cells(av, 19).Value = vStr1
            Cells(av, 20).Value = vStr2
            
        End If
        Next av
        
        'ThisWorkbook.Sheets("Setup").Save
    
    
        'Call buying_selling_rate
    End If
        
   ' Insert your code here.
   ' The following line shows how to reference a sheet within
   ' the loop by displaying the worksheet name in a dialog box.
    Workbooks(REGISTERfile).Activate
    Worksheets(xCount).Select
    
    Dim LastrowDate As Long
    LastrowDate = Range("A" & Rows.Count).End(xlUp).Row
    
    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
    
    'Add filter
    Rows("2:2").AutoFilter
    
    ActiveSheet.Range("A3:A" & LastrowDate).AutoFilter Field:=1, Criteria1:=vDateEntries _
        , Operator:=xlAnd
    
    Workbooks(REGISTERfile).Activate
    Worksheets(xCount).Select
    'Lastrow count for filtered date
    Dim r As Range, N As Long
    N = Cells(Rows.Count, 1).End(xlUp).Row
    Set r = Range("A1:A" & N).Cells.SpecialCells(xlCellTypeVisible)
    vVisble_row = r.Count - 2
    'MsgBox onshoreVisble_row
    
    If vVisble_row > 1 Then
    
    
        Range("A3:R" & N).SpecialCells(xlCellTypeVisible).Copy
    
        
        'Add sheet in thisworkbook to paste all filtered date
        ThisWorkbook.Sheets("Setup").Activate
        Sheets.Add
        ActiveSheet.Name = "filtered_data"
        
        Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        Sheets("filtered_data").Select
        
        'Remove duplicate to get unique Customer
        ActiveSheet.UsedRange.RemoveDuplicates Columns:=3, Header:=xlNo
        
        Dim LastrowCustomer As Long
        LastrowCustomer = Range("A" & Rows.Count).End(xlUp).Row
        
        
        'Get Customer number in column C
        For vCustomer = 1 To LastrowCustomer

            ThisWorkbook.Sheets("Setup").Activate
            Sheets("filtered_data").Select
            vCusAcc = Cells(vCustomer, 3).Value
            vCusName = Cells(vCustomer, 4).Value

            Workbooks(REGISTERfile).Activate
            Worksheets(xCount).Select

            'Filter column C
            ActiveSheet.UsedRange.AutoFilter Field:=3, Criteria1:=vCusAcc _
            , Operator:=xlAnd

            'Fill data in contra entries file
            advice_contra_FILE.Activate
            Worksheets(1).Select
            Range("B3").Value = vCusName
            Range("B5").Value = vDateEntries

'''''''''''''''''''Insert Buying rate and selling rate (Used for Contra entry) - GBP,USD,EUR(if MUR leave it blank)'''''''''''''''''''
            Workbooks(REGISTERfile).Activate
            Worksheets(xCount).Select

            If Worksheets(xCount).Name <> "MUR" Then
                advice_contra_FILE.Activate
                Worksheets(1).Activate
                Range("B7").Value = vStr1
                Range("C7").Value = vStr2
            End If
''''''''''''''''''''Insert Bill referencein row 12'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Workbooks(REGISTERfile).Activate
                Worksheets(xCount).Select
                
        
                'Lastrow count for filtered date
                Dim r1 As Range, n1 As Long
                n1 = Cells(Rows.Count, 1).End(xlUp).Row
                Set r1 = Range("A1:A" & n1).Cells.SpecialCells(xlCellTypeVisible)
                vVisble_row1 = r1.Count - 2
            
                'Insert row in advice contra - Discount excel(rows 12 based on the number of invoice for same customer)
                If vVisble_row1 > 2 Then
                           
                    For i = 2 To vVisble_row1
                        advice_contra_FILE.Activate
                        Worksheets(1).Activate
                        Rows("12:12").EntireRow.Insert
                    Next i
                    
                    'Insert Bill reference in row 12
                    Workbooks(REGISTERfile).Activate
                    Worksheets(xCount).Select
                    Range("B3:B" & n1).SpecialCells(xlCellTypeVisible).Copy
                    
                    advice_contra_FILE.Activate
                    Worksheets(1).Activate
                    Range("A11").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    
                    'Bill amount FCY - foreign amt
                    Workbooks(REGISTERfile).Activate
                    Worksheets(xCount).Select
                    Range("L3:L" & n1).SpecialCells(xlCellTypeVisible).Copy
                    
                    advice_contra_FILE.Activate
                    Worksheets(1).Activate
                    Range("B11").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    
                    'Bills amount in Local Currency (MUR) - no conversion, pickup from register
                    Workbooks(REGISTERfile).Activate
                    Worksheets(xCount).Select
                    Range("M3:M" & n1).SpecialCells(xlCellTypeVisible).Copy
                    
                    advice_contra_FILE.Activate
                    Worksheets(1).Activate
                    Range("C11").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    
                    
                    
'                    'Open Commission file to get the currency - USD or MUR
'                    Dim Filename1 As String
'
'                    'Open excel file based on path
'                    'If path is not found display the message "File not found"
'                    If Dir(Filename1) = "" Then
'
'                        MsgBox "File not found " & Filename1
'
'                    Exit Sub
'
'                    End If
'
'                    'Open file advice_contra
'                    Set commission_file = Workbooks.Open(Filename1, UpdateLinks:=False) 'Remove update popin window

'''''''''''''The commission is calculated USD 50 for FCY bills or MUR 500 for MUR bills and inserted in the sheet'''''''
                    Call commission_file
                                                          
                    
                End If
                
'''''''''''''''''The commission is converted to the local currency using the selling rate of the day'''''''''''''''''''''
                Workbooks(REGISTERfile).Activate
                Worksheets(xCount).Select
                Dim lastrowCom As Long
                lastrowCom = Range("C" & Rows.Count).End(xlUp).Row
                
                If Worksheets(xCount).Name <> "MUR" Then
                    Range("D11").Formula = "=C7+D11"
                    
                    Range("D11:D" & lastrowCom).Select
                    Selection.FillDown
                   
                End If
                   
                 'Create folder for current month
                Workbooks(REGISTERfile).Activate
                Worksheets(xCount).Select
                If Len(Dir(advicecontra_folder & Worksheets(xCount).Name, vbDirectory)) = 0 Then
                   MkDir advicecontra_folder & Worksheets(xCount).Name
                End If
                
                vPathSaveFile = advicecontra_folder & Worksheets(xCount).Name & "\" & "Commission_" & Worksheets(xCount).Name & ".xlsx"
                Debug.Print vPathSaveFile
                
                'Close file commission
                Workbooks(vCommissionFile).Activate
                
                Application.DisplayAlerts = False
                Application.EnableEvents = False
    
                Workbooks(vCommissionFile).SaveAs vPathSaveFile
                Workbooks(vCommissionFile).Close
                
                Application.DisplayAlerts = True
                Application.EnableEvents = True
                
        Next vCustomer
        
        ThisWorkbook.Sheets("Setup").Activate
        
        Application.DisplayAlerts = False
        Application.EnableEvents = False
    
        Sheets("filtered_data").Delete
    
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        
      
    End If

Next xCount

End Sub
