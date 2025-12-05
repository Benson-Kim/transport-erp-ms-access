Attribute VB_Name = "modValidation"
Option Compare Database
Option Explicit

Public Function IsPasswordStrong(ByVal strPassword As String) As Boolean
    If Len(strPassword) < 8 Then GoTo Fail
    If Not strPassword Like "*[a-z]*" Then GoTo Fail
    If Not strPassword Like "*[A-Z]*" Then GoTo Fail
    If Not strPassword Like "*[0-9]*" Then GoTo Fail
    If Not strPassword Like "*[!@#$%^&*()_+-=[]{}|;':,.<>?/]*" Then GoTo Fail
    
    IsPasswordStrong = True
    Exit Function
Fail:
    IsPasswordStrong = False
End Function

Public Function IsValidEmail(strEmail As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,})+$"
    IsValidEmail = regEx.Test(strEmail)
End Function

Public Function FormatIBAN(strIBAN As String) As String
    Dim strClean As String
    strClean = Replace(Replace(Replace(strIBAN, " ", ""), "-", ""), ".", "")
    Dim i As Integer
    For i = 1 To Len(strClean) Step 4
        FormatIBAN = FormatIBAN & Mid(strClean, i, 4) & IIf(i + 4 <= Len(strClean), " ", "")
    Next
    FormatIBAN = Trim(FormatIBAN)
End Function

' ========================================================
' COMPLETE SPANISH VAT (NIF-IVA) VALIDATOR FOR ACCESS (2025)
' Handles: Individuals, Companies, Non-residents, etc.
' Works with or without "ES" prefix
' Correctly allows digit OR letter as control character for companies
' ========================================================

Public Function IsSpanishVATValid(VAT As String) As Boolean
    Dim cleaned As String
    cleaned = UCase(Trim(Nz(VAT, "")))
    
    If cleaned = "" Then Exit Function
    
    ' Remove ES prefix if present
    If Left(cleaned, 2) = "ES" Then cleaned = Mid(cleaned, 3)
    
    ' Remove spaces, dots, hyphens
    cleaned = Replace(Replace(Replace(cleaned, "-", ""), " ", ""), ".", "")
    
    ' Must be 9 characters
    If Len(cleaned) <> 9 Then Exit Function
    
    Dim first As String
    first = Left(cleaned, 1)
    
    ' PERSONS (NIF / NIE / KLM)
    If IsNumeric(first) Or InStr("XYZKLM", first) > 0 Then
        IsSpanishVATValid = ValidatePerson(cleaned)
        
    ' COMPANIES & SPECIAL ENTITIES (CIF)
    ElseIf InStr("ABCDEFGHIJNPQRSUVW", first) > 0 Then
        IsSpanishVATValid = ValidateCompany(cleaned)
    End If
End Function

' ====================== PERSON VALIDATION ======================
Private Function ValidatePerson(code As String) As Boolean
    Dim num8 As String
    num8 = code
    
    Select Case Left(code, 1)
        Case "X": num8 = "0" & Mid(code, 2)
        Case "Y": num8 = "1" & Mid(code, 2)
        Case "Z": num8 = "2" & Mid(code, 2)
        ' K, L, M and numeric stay unchanged
    End Select
    
    ' First 8 chars must be numeric
    If Not IsNumeric(Left(num8, 8)) Then Exit Function
    
    Dim letters As String
    letters = "TRWAGMYFPDXBNJZSQVHLCKE"
    
    Dim pos As Long
    pos = CLng(Left(num8, 8)) Mod 23 + 1
    
    ValidatePerson = (Right(code, 1) = Mid(letters, pos, 1))
End Function

' ====================== COMPANY VALIDATION =====================
Private Function ValidateCompany(cif As String) As Boolean
    Dim entityType As String
    entityType = Left(cif, 1)
    
    Dim digits As String
    digits = Mid(cif, 2, 7)
    
    Dim control As String
    control = UCase(Right(cif, 1))
    
    If Not IsNumeric(digits) Then Exit Function
    
    Dim sum As Long, temp As Long, i As Integer
    
    ' Even positions (2,4,6)
    For i = 2 To 6 Step 2
        sum = sum + CLng(Mid(digits, i, 1))
    Next i
    
    ' Odd positions (1,3,5,7)
    For i = 1 To 7 Step 2
        temp = 2 * CLng(Mid(digits, i, 1))
        sum = sum + (temp \ 10) + (temp Mod 10)
    Next i
    
    Dim check As Integer
    check = (10 - (sum Mod 10)) Mod 10
    
    Dim letters As String
    letters = "JABCDEFGHI"
    
    Select Case entityType
        Case "A", "B", "C", "D", "E", "F", "G", "H", "S"
            ' Must be numeric
            ValidateCompany = (control = CStr(check))
        Case "J", "P", "Q", "R", "U", "V", "W"
            ' Must be a letter
            ValidateCompany = (control = Mid(letters, check + 1, 1))
        Case "N"
            ' Can be digit or letter
            ValidateCompany = (control = CStr(check)) Or (control = Mid(letters, check + 1, 1))
        Case Else
            ValidateCompany = False
    End Select
End Function

Sub TestSpanishVATs()
    Dim vats(1 To 12) As String
    Dim i As Integer
    Dim result As Boolean
    
    ' === 12 example VATs covering all types ===
    vats(1) = "ES12345678Z"    ' DNI
    vats(2) = "ESX1234567L"    ' NIE X
    vats(3) = "ESY7654321R"    ' NIE Y
    vats(4) = "ESZ1234567T"    ' NIE Z
    vats(5) = "ESK1234567J"    ' Special NIF K
    vats(6) = "ESL2345678S"    ' Special NIF L
    vats(7) = "ESM3456789P"    ' Special NIF M
    vats(8) = "ESA58818537"    ' CIF Company A (numeric control)
    vats(9) = "ESJ1234567B"    ' CIF Company J (letter control)
    vats(10) = "ESN1234567J"   ' CIF Company N (digit or letter)
    vats(11) = "ESV2345678E"   ' CIF Company V (letter control)
    vats(12) = "ESH63275145"   ' CIF Company W (letter control)
    
    ' Loop through all VATs and test
    For i = 1 To 12
        result = IsSpanishVATValid(vats(i))
        Debug.Print vats(i) & " ? " & IIf(result, "Valid", "Invalid")
    Next i
End Sub

