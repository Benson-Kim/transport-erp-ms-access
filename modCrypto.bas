Attribute VB_Name = "modCrypto"

' MODULE: modCrypto
' PURPOSE: Secure password hashing using Windows CNG API (BCrypt)
'          Pure Windows API - no .NET required

Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modCrypto"

' BCrypt API declarations
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
    ByRef phAlgorithm As LongPtr, _
    ByVal pszAlgId As LongPtr, _
    ByVal pszImplementation As LongPtr, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByRef phHash As LongPtr, _
    ByVal pbHashObject As LongPtr, _
    ByVal cbHashObject As Long, _
    ByVal pbSecret As LongPtr, _
    ByVal cbSecret As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" ( _
    ByVal hHash As LongPtr, _
    ByVal pbInput As LongPtr, _
    ByVal cbInput As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" ( _
    ByVal hHash As LongPtr, _
    ByVal pbOutput As LongPtr, _
    ByVal cbOutput As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" ( _
    ByVal hHash As LongPtr) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByVal pszProperty As LongPtr, _
    ByRef pbOutput As Any, _
    ByVal cbOutput As Long, _
    ByRef pcbResult As Long, _
    ByVal dwFlags As Long) As Long

Private Const BCRYPT_OBJECT_LENGTH As String = "ObjectLength"

Private Const BCRYPT_SHA256_ALGORITHM As String = "SHA256"

Private EmptyByteArray() As Byte


' PUBLIC: HashPassword – Returns Base64(salt + hash)

Public Function HashPassword(ByVal strPassword As String) As String
    On Error GoTo ErrorHandler

    Dim arrSalt() As Byte
    Dim arrPassword() As Byte
    Dim arrHash() As Byte
    Dim arrCombined() As Byte
    Dim arrSalted() As Byte
    Dim strResult As String
    Dim i As Long

    ' Generate 16-byte random salt
    Randomize Timer
    ReDim arrSalt(0 To 15)
    For i = 0 To 15
        arrSalt(i) = Int(Rnd * 256)
    Next i

    ' Convert password to UTF-8 bytes
    arrPassword = StringToUTF8(strPassword)

    ' Combine salt + password
    ReDim arrSalted(0 To UBound(arrSalt) + UBound(arrPassword) + 1)
    Call CopyMemory(arrSalted(0), arrSalt(0), 16)
    Call CopyMemory(arrSalted(16), arrPassword(0), UBound(arrPassword) + 1)

    ' Compute SHA-256 hash using Windows BCrypt
    arrHash = ComputeSHA256(arrSalted)

    If UBound(arrHash) < 0 Then
        HashPassword = ""
        Exit Function
    End If

    ' Combine salt + hash for storage
    ReDim arrCombined(0 To 47) ' 16 bytes salt + 32 bytes hash
    Call CopyMemory(arrCombined(0), arrSalt(0), 16)
    Call CopyMemory(arrCombined(16), arrHash(0), 32)

    strResult = EncodeBase64(arrCombined)
    HashPassword = strResult

CleanExit:
    Exit Function

ErrorHandler:
    On Error Resume Next
    modUtilities.LogError "HashPassword", Err.Number, Err.Description
    HashPassword = ""
    Resume CleanExit
End Function


' PUBLIC: VerifyPassword – Compare input vs stored hash

Public Function VerifyPassword(ByVal strPassword As String, ByVal strStoredHash As String) As Boolean
    On Error GoTo ErrorHandler

    Dim arrStored() As Byte
    Dim arrSalt(0 To 15) As Byte
    Dim arrPassword() As Byte
    Dim arrSalted() As Byte
    Dim arrComputedHash() As Byte
    Dim i As Long

    ' Decode stored Base64
    arrStored = DecodeBase64(strStoredHash)
    If UBound(arrStored) < 47 Then GoTo Fail ' Need 48 bytes

    ' Extract salt (first 16 bytes)
    Call CopyMemory(arrSalt(0), arrStored(0), 16)

    ' Convert password to UTF-8
    arrPassword = StringToUTF8(strPassword)

    ' Combine salt + password
    ReDim arrSalted(0 To 15 + UBound(arrPassword) + 1)
    Call CopyMemory(arrSalted(0), arrSalt(0), 16)
    Call CopyMemory(arrSalted(16), arrPassword(0), UBound(arrPassword) + 1)

    ' Compute hash
    arrComputedHash = ComputeSHA256(arrSalted)

    If UBound(arrComputedHash) < 31 Then GoTo Fail

    ' Compare computed hash with stored hash (constant-time comparison)
    VerifyPassword = True
    For i = 0 To 31
        If arrComputedHash(i) <> arrStored(16 + i) Then
            VerifyPassword = False
        End If
    Next i

    GoTo CleanExit

Fail:
    VerifyPassword = False

CleanExit:
    Exit Function

ErrorHandler:
    VerifyPassword = False
    Resume CleanExit
End Function


' PRIVATE: ComputeSHA256 using Windows BCrypt API

Private Function ComputeSHA256(arrData() As Byte) As Byte()
    Dim hAlg As LongPtr
    Dim hHash As LongPtr
    Dim arrResult() As Byte
    Dim hashObj() As Byte
    Dim lResult As Long
    Dim cbHashObj As Long
    Dim cbReturned As Long
    Dim strAlg As String
    Dim pbData As LongPtr
    Dim cbData As Long

    On Error GoTo ErrorHandler
    
    ReDim arrResult(0 To 31)

    strAlg = BCRYPT_SHA256_ALGORITHM

    ' 1. Open algorithm provider
    lResult = BCryptOpenAlgorithmProvider(hAlg, StrPtr(strAlg), 0, 0)
    If lResult <> 0 Then GoTo ErrorHandler

    ' 2. Get required hash object size
    lResult = BCryptGetProperty(hAlg, StrPtr(BCRYPT_OBJECT_LENGTH), cbHashObj, 4, cbReturned, 0)
    If lResult <> 0 Then GoTo ErrorHandler
    
    ReDim hashObj(0 To cbHashObj - 1)

    ' 3. Create hash object — NOW with proper buffer
    lResult = BCryptCreateHash(hAlg, hHash, VarPtr(hashObj(0)), cbHashObj, 0, 0, 0)
    If lResult <> 0 Then GoTo ErrorHandler

    ' 4. Hash the data (64-bit safe)
    pbData = VarPtr(arrData(0))
    cbData = UBound(arrData) + 1
    lResult = BCryptHashData(hHash, pbData, cbData, 0)
    If lResult <> 0 Then GoTo ErrorHandler

    ' 5. Finish hash
    lResult = BCryptFinishHash(hHash, VarPtr(arrResult(0)), 32, 0)
    If lResult <> 0 Then GoTo ErrorHandler

    ComputeSHA256 = arrResult
    GoTo Cleanup

ErrorHandler:
    ComputeSHA256 = EmptyByteArray

Cleanup:
    If hHash <> 0 Then BCryptDestroyHash hHash
    If hAlg <> 0 Then BCryptCloseAlgorithmProvider hAlg, 0
End Function


' PRIVATE: Convert String to UTF-8 Byte Array

Private Function StringToUTF8(ByVal strText As String) As Byte()
    Dim objStream As Object
    Dim arrBytes() As Byte

    On Error GoTo ErrorHandler

    ' Use ADODB.Stream for UTF-8 conversion
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.Open
    objStream.WriteText strText
    objStream.Position = 0
    objStream.Type = 1 ' adTypeBinary
    objStream.Position = 3 ' Skip UTF-8 BOM
    arrBytes = objStream.Read
    objStream.Close

    StringToUTF8 = arrBytes
    Set objStream = Nothing
    Exit Function

ErrorHandler:
    ReDim arrBytes(0 To -1)
    StringToUTF8 = arrBytes
    If Not objStream Is Nothing Then objStream.Close
    Set objStream = Nothing
End Function


' PRIVATE: Base64 Encode/Decode using MSXML

Private Function EncodeBase64(arrData() As Byte) As String
    Dim objXML As Object
    Dim objNode As Object

    On Error GoTo ErrorHandler

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = Replace(objNode.Text, vbLf, "")

    Set objNode = Nothing
    Set objXML = Nothing
    Exit Function

ErrorHandler:
    EncodeBase64 = ""
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Private Function DecodeBase64(ByVal strData As String) As Byte()
    Dim objXML As Object
    Dim objNode As Object
    Dim arrEmpty() As Byte

    On Error GoTo ErrorHandler

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = objNode.nodeTypedValue

    Set objNode = Nothing
    Set objXML = Nothing
    Exit Function

ErrorHandler:
    ReDim arrEmpty(0 To -1)
    DecodeBase64 = arrEmpty
    Set objNode = Nothing
    Set objXML = Nothing
End Function
