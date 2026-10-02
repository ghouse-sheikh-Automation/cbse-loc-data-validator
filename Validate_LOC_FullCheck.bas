Attribute VB_Name = "Module1"
Option Explicit

Sub Validate_LOC_FullCheck()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim ws As Worksheet, newsheet As Worksheet
    Dim lastrow As Long, lastcol As Long, x As Long, y As Long
    Dim CnameCol As Long, Mnamecol As Long, Fnamecol As Long
    Dim Gendercol As Long, Castcol As Long, Handicapcol As Long
    Dim Annualcol As Long, Onlychildcol As Long, Rollcol As Long, Preyearcol As Long
    Dim Boardcol As Long, Admcol As Long, Admdatecol As Long, Mincol As Long
    Dim schcol As Long, serialcol As Long, irecol As Long
    Dim admVal As String
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("CLASSXI")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "CLASSXI Sheet not exists!", vbCritical
        GoTo cleanexit
    End If
    
    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastcol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Header Column Mapping
    For x = 1 To lastcol
        Select Case UCase(Trim(ws.Cells(1, x).Value))
            Case "CNAME": CnameCol = x
            Case "MNAME": Mnamecol = x
            Case "FNAME": Fnamecol = x
            Case "SEX", "GENDER": Gendercol = x
            Case "CAST", "CATEGORY", "CAT": Castcol = x
            Case "HANDICAP", "PH": Handicapcol = x
            Case "ANNUAL_INC", "ANNUAL INCOME", "INCOME": Annualcol = x
            Case "ONLY_CHILD", "SINGLE CHILD": Onlychildcol = x
            Case "X_ROLL", "ROLL": Rollcol = x
            Case "X_YEAR", "PASS YEAR", "YEAR": Preyearcol = x
            Case "X_BOARD", "BOARD": Boardcol = x
            Case "ADM_SRL", "ADM NO", "ADMN NO": Admcol = x
            Case "ADM_DATE", "ADMN DATE": Admdatecol = x
            Case "MINORITY", "MIN": Mincol = x
            Case "SCH_NO": schcol = x
            Case "SERIAL": serialcol = x
            Case "IREGDNO": irecol = x
        End Select
    Next x
    
    ' Error Sheet Setup
    On Error Resume Next
    Set newsheet = Worksheets("Error")
    On Error GoTo 0
    
    If newsheet Is Nothing Then
        Set newsheet = Worksheets.Add(after:=Worksheets(Worksheets.Count))
        newsheet.Name = "Error"
    Else
        newsheet.Cells.Clear
    End If
    
    newsheet.Cells(1, 1).Value = "Row Number"
    newsheet.Cells(1, 2).Value = "Column"
    newsheet.Cells(1, 3).Value = "Reason for Error"
    newsheet.Range("A1:C1").Font.Bold = True
    
    Dim nextrow As Long
    nextrow = 2
    
    For x = 2 To lastrow
        
        ' ==========================================
        ' STEP 1: CLEANING & DEFAULTS FIRST (Data Standardization)
        ' ==========================================
        
        ' Names to Uppercase
        If Trim(ws.Cells(x, CnameCol).Value) <> "" Then ws.Cells(x, CnameCol).Value = UCase(Trim(ws.Cells(x, CnameCol).Value))
        If Trim(ws.Cells(x, Mnamecol).Value) <> "" Then ws.Cells(x, Mnamecol).Value = UCase(Trim(ws.Cells(x, Mnamecol).Value))
        If Trim(ws.Cells(x, Fnamecol).Value) <> "" Then ws.Cells(x, Fnamecol).Value = UCase(Trim(ws.Cells(x, Fnamecol).Value))
        
        ' Gender Standardization
        If Trim(ws.Cells(x, Gendercol).Value) <> "" Then
            Select Case UCase(Trim(ws.Cells(x, Gendercol).Value))
                Case "F", "FEMALE": ws.Cells(x, Gendercol).Value = "F"
                Case "M", "MALE": ws.Cells(x, Gendercol).Value = "M"
            End Select
        End If
        
        ' Category Default
            If Trim(ws.Cells(x, Castcol).Value) = "" Then
                ws.Cells(x, Castcol).Value = "G"
            Else
                Select Case UCase(Trim(ws.Cells(x, Castcol).Value))
                    Case "ST", "T": ws.Cells(x, Castcol).Value = "T"
                    Case "SC", "C": ws.Cells(x, Castcol).Value = "C"
                    Case "OBC", "O": ws.Cells(x, Castcol).Value = "O"
                    Case Else: ws.Cells(x, Castcol).Value = "G"
                End Select
            End If
        
        ' Handicap Default (Fills BEFORE empty check loop)
        If Trim(ws.Cells(x, Handicapcol).Value) = "" Then ws.Cells(x, Handicapcol).Value = "NA"
        
        ' Income Cleaning
        If Trim(ws.Cells(x, Annualcol).Value) <> "" Then
            If InStr(ws.Cells(x, Annualcol).Value, ",") > 0 Then ws.Cells(x, Annualcol).Value = Replace(ws.Cells(x, Annualcol).Value, ",", "")
        End If
        
        ' Single Child Default
            If Trim(ws.Cells(x, Onlychildcol).Value) = "" Or UCase(Trim(ws.Cells(x, Onlychildcol).Value)) = "NO" Then
                ws.Cells(x, Onlychildcol).Value = "N"
            ElseIf UCase(Trim(ws.Cells(x, Onlychildcol).Value)) = "YES" Or UCase(Trim(ws.Cells(x, Onlychildcol).Value)) = "Y" Then
                ws.Cells(x, Onlychildcol).Value = "Y"
            End If
        
        ' Admission Date Formatting Fix (Handles 7/8 digit numeric dates like 9012014 or 29012014)
        If Trim(ws.Cells(x, Admdatecol).Value) <> "" Then
            admVal = CStr(Trim(ws.Cells(x, Admdatecol).Value))
            If Len(admVal) = 7 Then
                ws.Cells(x, Admdatecol).NumberFormat = "@"
                ws.Cells(x, Admdatecol).Value = "0" & admVal
            End If
        End If
        
        ' Minority Default
            If UCase(Trim(ws.Cells(x, Mincol).Value)) = "YES" Or UCase(Trim(ws.Cells(x, Mincol).Value)) = "Y" Then
                ws.Cells(x, Mincol).Value = "Y"
            Else
                ws.Cells(x, Mincol).Value = "N"
            End If
        
        ' ==========================================
        ' STEP 2: VALIDATION & ERROR REPORTING
        ' ==========================================
        
        ' Core IDs
        If Trim(ws.Cells(x, schcol).Value) = "" Then
            newsheet.Cells(nextrow, 1).Value = x
            newsheet.Cells(nextrow, 2).Value = "SCH_NO"
            newsheet.Cells(nextrow, 3).Value = "Empty SCH_NO"
            nextrow = nextrow + 1
        End If
        If Trim(ws.Cells(x, serialcol).Value) = "" Then
            newsheet.Cells(nextrow, 1).Value = x
            newsheet.Cells(nextrow, 2).Value = "SERIAL"
            newsheet.Cells(nextrow, 3).Value = "Empty SERIAL"
            nextrow = nextrow + 1
        End If
        
        ' Columns 10 to 19 Mandatory Empty Check (Excludes HANDICAP if auto-filled)
        For y = 10 To 19
            If y <= lastcol And y <> Handicapcol Then
                If Trim(ws.Cells(x, y).Value) = "" Then
                    newsheet.Cells(nextrow, 1).Value = x
                    newsheet.Cells(nextrow, 2).Value = ws.Cells(1, y).Value
                    newsheet.Cells(nextrow, 3).Value = "Empty cell found in Subject/Data Column"
                    nextrow = nextrow + 1
                End If
            End If
        Next y
        
        ' Student Details Validation
        If Trim(ws.Cells(x, CnameCol).Value) = "" Then
            newsheet.Cells(nextrow, 1).Value = x
            newsheet.Cells(nextrow, 2).Value = "CNAME"
            newsheet.Cells(nextrow, 3).Value = "Student Name Empty"
            nextrow = nextrow + 1
        End If
        If Trim(ws.Cells(x, Mnamecol).Value) = "" Then
            newsheet.Cells(nextrow, 1).Value = x
            newsheet.Cells(nextrow, 2).Value = "MNAME"
            newsheet.Cells(nextrow, 3).Value = "Mother Name Empty"
            nextrow = nextrow + 1
        End If
        If Trim(ws.Cells(x, Fnamecol).Value) = "" Then
            newsheet.Cells(nextrow, 1).Value = x
            newsheet.Cells(nextrow, 2).Value = "FNAME"
            newsheet.Cells(nextrow, 3).Value = "Father Name Empty"
            nextrow = nextrow + 1
        End If
        
        If Trim(ws.Cells(x, Gendercol).Value) = "" Then
            newsheet.Cells(nextrow, 1).Value = x
            newsheet.Cells(nextrow, 2).Value = "SEX"
            newsheet.Cells(nextrow, 3).Value = "Gender Empty"
            nextrow = nextrow + 1
        End If
    
        If Trim(ws.Cells(x, Annualcol).Value) = "" Then
            newsheet.Cells(nextrow, 1).Value = x
            newsheet.Cells(nextrow, 2).Value = "ANNUAL_INC"
            newsheet.Cells(nextrow, 3).Value = "Income Empty"
            nextrow = nextrow + 1
        ElseIf InStr(UCase(ws.Cells(x, Annualcol).Value), "L") > 0 Then
                newsheet.Cells(nextrow, 1).Value = x
                newsheet.Cells(nextrow, 2).Value = "ANNUAL_INC"
                newsheet.Cells(nextrow, 3).Value = "Invalid (LAKH text format): " & ws.Cells(x, Annualcol).Value
                nextrow = nextrow + 1
        End If
        
        If Trim(ws.Cells(x, Rollcol).Value) = "" Then
            newsheet.Cells(nextrow, 1).Value = x
            newsheet.Cells(nextrow, 2).Value = "X_ROLL"
            newsheet.Cells(nextrow, 3).Value = "Empty Roll"
            nextrow = nextrow + 1
        End If
        If Trim(ws.Cells(x, Preyearcol).Value) <> "" Then
            If Val(ws.Cells(x, Preyearcol).Value) <> 2024 And Val(ws.Cells(x, Preyearcol).Value) <> 2025 Then
                newsheet.Cells(nextrow, 1).Value = x
                newsheet.Cells(nextrow, 2).Value = "X_YEAR"
                newsheet.Cells(nextrow, 3).Value = "Invalid Year: " & ws.Cells(x, Preyearcol).Value
                nextrow = nextrow + 1
            End If
        End If
        If Trim(ws.Cells(x, Boardcol).Value) = "" Then
            newsheet.Cells(nextrow, 1).Value = x
            newsheet.Cells(nextrow, 2).Value = "X_BOARD"
            newsheet.Cells(nextrow, 3).Value = "Empty Board"
            nextrow = nextrow + 1
        End If
        If Trim(ws.Cells(x, Admcol).Value) = "" Then
            newsheet.Cells(nextrow, 1).Value = x
            newsheet.Cells(nextrow, 2).Value = "ADM_SRL"
            newsheet.Cells(nextrow, 3).Value = "Empty Adm No"
            nextrow = nextrow + 1
        End If
        If Trim(ws.Cells(x, Admdatecol).Value) = "" Then
            newsheet.Cells(nextrow, 1).Value = x
            newsheet.Cells(nextrow, 2).Value = "ADM_DATE"
            newsheet.Cells(nextrow, 3).Value = "Empty Adm Date"
            nextrow = nextrow + 1
        End If
        
    Next x
    
    newsheet.Columns("A:C").AutoFit
    If nextrow = 2 Then
        newsheet.Cells(2, 1).Value = "All Good!"
        MsgBox "LOC Validation Done - Clean Data Verified!", vbInformation
    Else
        MsgBox "Validation Done! Total Errors Found: " & (nextrow - 2) & " - Check Error Sheet.", vbExclamation
    End If
    
cleanexit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
