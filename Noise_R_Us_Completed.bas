Attribute VB_Name = "NoiseRUsCompleted"
Option Explicit

' Replacement standard module for the original Sheet1 layout.
' All prices are supplied by the operator; no assignment rates are assumed.
' Condition codes follow the sheet: 1 = good, 2 = fair, 3 = bad.

Public Sub GetInput()
    Dim ws As Worksheet
    Dim groupCode As String, delivery As String, description As String
    Dim days As Long, beforeCondition As Long
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    If ws.ProtectContents Then
        MsgBox "Unprotect Sheet1 before running this macro.", vbExclamation
        Exit Sub
    End If
    If Len(Trim$(CStr(ws.Range("B2").Value))) > 0 Then
        If MsgBox("Starting a new rental replaces the displayed rental and invoice. Continue?", _
                  vbYesNo + vbQuestion + vbDefaultButton2, "New rental") <> vbYes Then Exit Sub
    End If

    ' Collect everything first: Cancel leaves the current sheet unchanged.
    If Not AskChoice("Sound system group: A, B or C", "Group", "GROUP", groupCode) Then Exit Sub
    If Not AskDays(days) Then Exit Sub
    If Not AskChoice("Delivery required? Enter Y or N.", "Delivery", "DELIVERY", delivery) Then Exit Sub
    If Not AskCondition("Condition before rental: 1 = good, 2 = fair, 3 = bad.", beforeCondition) Then Exit Sub
    Select Case groupCode
        Case "A": description = CStr(ws.Range("G10").Value)
        Case "B": description = CStr(ws.Range("G11").Value)
        Case "C": description = CStr(ws.Range("G12").Value)
    End Select

    ws.Range("B2").Value = groupCode
    ws.Range("B3").Value = description
    ws.Range("E2").Value = days
    ws.Range("E3").Value = delivery
    ws.Range("H2").Value = beforeCondition
    ws.Range("H3").ClearContents
    ws.Range("B10:B13").ClearContents
    MsgBox "Rental captured. Run ReturnSoundSystem when the equipment is returned.", vbInformation
    Exit Sub
Failed:
    MsgBox "Could not complete GetInput: " & Err.Description & _
           vbCrLf & "Check the sheet before trying again; changes may be partial.", vbExclamation
End Sub

Public Sub ReturnSoundSystem()
    Dim ws As Worksheet
    Dim groupCode As String, delivery As String
    Dim days As Long, beforeCondition As Long, afterCondition As Long
    Dim dailyRate As Currency, deliveryFee As Currency, damageFee As Currency
    Dim rentalFee As Currency
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    If ws.ProtectContents Then
        MsgBox "Unprotect Sheet1 before running this macro.", vbExclamation
        Exit Sub
    End If
    ' Validate the existing rental even if cells were manually edited.
    If IsError(ws.Range("B2").Value) Or IsError(ws.Range("E2").Value) Or _
       IsError(ws.Range("E3").Value) Or IsError(ws.Range("H2").Value) Then GoTo BadRental
    groupCode = UCase$(Trim$(CStr(ws.Range("B2").Value)))
    delivery = UCase$(Trim$(CStr(ws.Range("E3").Value)))
    If Not IsChoice(groupCode, "GROUP") Then GoTo BadRental
    If Not IsChoice(delivery, "DELIVERY") Then GoTo BadRental
    If Not ReadWholeNumber(CStr(ws.Range("E2").Value), 1, 10000, days) Then GoTo BadRental
    If Not ReadWholeNumber(CStr(ws.Range("H2").Value), 1, 3, beforeCondition) Then GoTo BadRental

    ' Gather every value before updating any sheet cells.
    If Not AskCondition("Condition after return: 1 = good, 2 = fair, 3 = bad.", afterCondition) Then Exit Sub
    If Not AskMoney("Enter the agreed daily rate for group " & groupCode & ".", dailyRate) Then Exit Sub
    If delivery = "Y" Then
        If Not AskMoney("Enter the agreed delivery fee for this rental.", deliveryFee) Then Exit Sub
    Else
        deliveryFee = 0
    End If
    If afterCondition > beforeCondition Then
        If Not AskMoney("The condition code worsened. Enter the agreed condition/damage charge " & _
                        "for this rental (enter 0 if none). No policy is assumed.", damageFee) Then Exit Sub
    Else
        damageFee = 0
    End If
    rentalFee = dailyRate * days

    ws.Range("H3").Value = afterCondition
    ws.Range("B10").Value = rentalFee
    ws.Range("B11").Value = deliveryFee
    ws.Range("B12").Value = damageFee
    ws.Range("B10:B13").NumberFormat = "0.00"
    ws.Range("B13").Formula = "=SUM(B10:B12)"
    ws.Range("B13").Calculate
    MsgBox "Invoice calculated. Please check the entered rates and charges.", vbInformation
    Exit Sub
BadRental:
    MsgBox "Rental details are missing or invalid. Run GetInput first, or correct B2, E2, E3 and H2.", vbExclamation
    Exit Sub
Failed:
    MsgBox "Could not complete ReturnSoundSystem: " & Err.Description & _
           vbCrLf & "Check the sheet before trying again; changes may be partial.", vbExclamation
End Sub

Private Function AskChoice(ByVal prompt As String, ByVal title As String, _
                           ByVal kind As String, ByRef answer As String) As Boolean
    Dim response As Variant, text As String
    Do
        response = Application.InputBox(prompt, title, Type:=2)
        If VarType(response) = vbBoolean Then Exit Function ' Cancel
        text = UCase$(Trim$(CStr(response)))
        If IsChoice(text, kind) Then
            answer = text
            AskChoice = True
            Exit Function
        End If
        MsgBox "Invalid entry. Please use one of the listed choices, or Cancel.", vbExclamation
    Loop
End Function

Private Function IsChoice(ByVal text As String, ByVal kind As String) As Boolean
    Select Case kind
        Case "GROUP": IsChoice = (text = "A" Or text = "B" Or text = "C")
        Case "DELIVERY": IsChoice = (text = "Y" Or text = "N")
    End Select
End Function

Private Function AskDays(ByRef days As Long) As Boolean
    Dim response As Variant
    Do
        response = Application.InputBox("Number of rental days (whole number from 1 to 10000):", "Days", Type:=2)
        If VarType(response) = vbBoolean Then Exit Function
        If ReadWholeNumber(CStr(response), 1, 10000, days) Then
            AskDays = True
            Exit Function
        End If
        MsgBox "Enter a whole number from 1 to 10000, or Cancel.", vbExclamation
    Loop
End Function

Private Function AskCondition(ByVal prompt As String, ByRef result As Long) As Boolean
    Dim response As Variant
    Do
        response = Application.InputBox(prompt, "Condition", Type:=2)
        If VarType(response) = vbBoolean Then Exit Function
        If ReadWholeNumber(CStr(response), 1, 3, result) Then
            AskCondition = True
            Exit Function
        End If
        MsgBox "Enter 1, 2 or 3, or Cancel.", vbExclamation
    Loop
End Function

Private Function ReadWholeNumber(ByVal text As String, ByVal minimum As Long, _
                                 ByVal maximum As Long, ByRef result As Long) As Boolean
    Dim number As Double
    On Error GoTo Invalid
    If Not IsNumeric(text) Then Exit Function
    number = CDbl(text)
    If number < minimum Or number > maximum Then Exit Function
    If number <> Fix(number) Then Exit Function
    result = CLng(number)
    ReadWholeNumber = True
Invalid:
End Function

Private Function AskMoney(ByVal prompt As String, ByRef amount As Currency) As Boolean
    Dim response As Variant
    Do
        response = Application.InputBox(prompt & vbCrLf & _
                   "Enter 0 to 1000000 with at most two decimal places.", "Amount", Type:=2)
        If VarType(response) = vbBoolean Then Exit Function
        If ReadMoney(CStr(response), amount) Then
            AskMoney = True
            Exit Function
        End If
        MsgBox "Invalid amount. Use a non-negative number up to 1000000 with at most two decimal places, or Cancel.", vbExclamation
    Loop
End Function

Private Function ReadMoney(ByVal text As String, ByRef amount As Currency) As Boolean
    ' Decimal Variant avoids binary floating-point tolerance in the cents check.
    Dim number As Variant
    On Error GoTo Invalid
    If Not IsNumeric(text) Then Exit Function
    number = CDec(text)
    If number < 0 Or number > 1000000 Then Exit Function
    If number * 100 <> Fix(number * 100) Then Exit Function
    amount = CCur(number)
    ReadMoney = True
Invalid:
End Function
