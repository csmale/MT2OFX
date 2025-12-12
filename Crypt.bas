Attribute VB_Name = "Crypt"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : Crypt
'    Project    : MT2OFX
'
'    Description: Functions related to encryption, hashes etc
'
'    Modified   : $Author: Colin $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/Crypt.bas 4     8/11/08 18:53 Colin $"
' $History: Crypt.bas $
' 
' *****************  Version 4  *****************
' User: Colin        Date: 8/11/08    Time: 18:53
' Updated in $/MT2OFX
' Initial working version
' NB: base64 encoding won't work on Windows 2000; need custom
' implementation!

'</CSCC>
Private Declare Function CryptAcquireContext _
                Lib "advapi32.dll" _
                Alias "CryptAcquireContextA" (ByRef phProv As Long, _
                                              ByVal pszContainer As String, _
                                              ByVal pszProvider As String, _
                                              ByVal dwProvType As Long, _
                                              ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext _
                Lib "advapi32.dll" (ByVal hProv As Long, _
                                    ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash _
                Lib "advapi32.dll" (ByVal hProv As Long, _
                                    ByVal Algid As Long, _
                                    ByVal hKey As Long, _
                                    ByVal dwFlags As Long, _
                                    ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash _
                Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptHashData _
                Lib "advapi32.dll" (ByVal hHash As Long, _
                                    ByRef pbData As Any, _
                                    ByVal dwDataLen As Long, _
                                    ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam _
                Lib "advapi32.dll" (ByVal hHash As Long, _
                                    ByVal dwParam As Long, _
                                    ByRef pbData As Any, _
                                    ByRef pdwDataLen As Long, _
                                    ByVal dwFlags As Long) As Long
Private Declare Function CryptBinaryToString _
                Lib "crypt32.dll" _
                Alias "CryptBinaryToStringA" _
                                    (ByRef pbBinary As Byte, _
                                    ByVal cbBinary As Long, _
                                    ByVal dwFlags As Long, _
                                    ByVal pszString As String, _
                                    ByRef pcchString As Long) As Long


Private Const MS_DEF_PROV As String = "Microsoft Base Cryptographic Provider v1.0"
Private Const CRYPT_NEWKEYSET As Long = &H8
Private Const PROV_RSA_FULL As Long = &H1
Private Const ALG_CLASS_HASH As Long = &H8000&
Private Const ALG_TYPE_ANY As Long = &H0
Private Const ALG_SID_MD5 As Long = &H3
Private Const CALG_MD5 As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)
Private Const HP_HASHVAL As Long = &H2 ' Hash value
Private Const CRYPT_STRING_BASE64 As Long = &H1
Private Const CRYPT_STRING_HEX As Long = &H4

Private Declare Function GetLastError Lib "kernel32.dll" () As Long

Private hCryptProv As Long
Private hCryptHash As Long
Private lHashErr As Long


'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       MD5HashStart
' Description:       Initialise an MD5 hashing operation
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       5/22/2008-11:29:30
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function MD5HashStart() As Boolean
    lHashErr = 0
    If Not GetContext() Then
        MD5HashStart = False
        Exit Function
    End If
    ' Create new MD5 hash
    MD5HashStart = (CryptCreateHash(hCryptProv, CALG_MD5, 0&, 0&, hCryptHash) <> 0)
    If Not MD5HashStart Then
        lHashErr = GetLastError
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       MD5HashDataString
' Description:       Add a string into the MD5 hash
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       5/22/2008-11:29:30
'
' Parameters :       sData (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function MD5HashDataString(sData As String) As Boolean
    Dim bData() As Byte
    Dim lBufLen As Long
    Dim lBufPtr As Long
    lHashErr = 0
    lBufPtr = StrPtr(sData)
    lBufLen = LenB(sData)
    MD5HashDataString = (CryptHashData(hCryptHash, ByVal lBufPtr, lBufLen, 0&) <> 0)
    If Not MD5HashDataString Then
        lHashErr = GetLastError
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       MD5HashEnd
' Description:       Conclude an MD5 hashing operation and return the result
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       5/22/2008-11:29:30
'
' Parameters :       bHashOut() (Byte)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function MD5HashEnd(bHashOut() As Byte) As Boolean
    Dim bHash() As Byte
    Dim lBufLen As Long
    lHashErr = 0
    
    MD5HashEnd = False
    ' Get buffer length for hash
    If (CryptGetHashParam(hCryptHash, HP_HASHVAL, ByVal 0&, lBufLen, 0&)) Then
        ReDim bHash(0 To lBufLen - 1) As Byte

        If (CryptGetHashParam(hCryptHash, HP_HASHVAL, bHash(0), lBufLen, 0&)) Then
            MD5HashEnd = (lBufLen > 0) ' Return final hash buffer
            bHashOut = bHash
        Else
            lHashErr = GetLastError
        End If
    Else
        lHashErr = GetLastError
    End If
    ' Done with hash object
    Call CryptDestroyHash(hCryptHash)
    hCryptHash = 0
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       HashHexString
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       5/22/2008-11:29:30
'
' Parameters :       bData() (Byte)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function HashHexString(bData() As Byte) As String
    Dim i As Long
    Dim sTmp As String
    sTmp = ""
    For i = LBound(bData) To UBound(bData)
        sTmp = sTmp & Right$("0" & Hex(bData(i)), 2)
    Next
    HashHexString = sTmp
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       BinaryToString
' Description:       Convert binary hash value to a string representation
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       5/22/2008-11:29:30
'
' Parameters :       bytBuf() (Byte)
'                    lFlags (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function BinaryToString(bytBuf() As Byte, lFlags As Long) As String
    Dim lngOutLen As Long
    Dim strBase64 As String
    Dim lTmp As Long
    Dim sTmp As String
    
    BinaryToString = ""

    'Determine Base64 output String length required.
    lngOutLen = 0
    lTmp = CryptBinaryToString(bytBuf(0), _
        UBound(bytBuf) + 1, _
        lFlags, _
        vbNullString, _
        lngOutLen)
        
    If lTmp = 0 Then
        LogMessage True, True, "Error " & Err.LastDllError & " from CryptBinaryToString(len)"
        Exit Function
    End If
        
    'Convert binary to Base64.
    strBase64 = String(lngOutLen, vbNullChar)
    
    If CryptBinaryToString(bytBuf(0), _
        UBound(bytBuf) + 1, _
        lFlags, _
        strBase64, _
        lngOutLen) <> 0 Then
        'Use the Base64 output. trim off the CR+LF+Null which the API appends.
        sTmp = Replace$(Left(strBase64, lngOutLen), vbTab, "")
        sTmp = Replace$(sTmp, vbLf, "")
        sTmp = Replace$(sTmp, vbCr, "")
        sTmp = Replace$(sTmp, " ", "")
        BinaryToString = sTmp
        
    Else
        LogMessage True, True, "Error " & Err.LastDllError & " from CryptBinaryToString(data)"
        Exit Function
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       Base64Encode
' Description:       Convert byte array to Base64-encoded string
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       5/22/2008-11:29:30
'
' Parameters :       b() (Byte)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Base64Encode(b() As Byte) As String
    Base64Encode = BinaryToString(b, CRYPT_STRING_BASE64)
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetContext
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       5/22/2008-11:29:30
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Function GetContext() As Boolean
    GetContext = True
    If hCryptProv = 0 Then
        If (CryptAcquireContext(hCryptProv, vbNullString, MS_DEF_PROV, PROV_RSA_FULL, 0&) = 0) Then
            If (CryptAcquireContext(hCryptProv, vbNullString, MS_DEF_PROV, PROV_RSA_FULL, CRYPT_NEWKEYSET) = 0) Then
                lHashErr = GetLastError
                GetContext = False ' Failed to acquire cryptographic context
            End If
        End If
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetHashError
' Description:       Returns most recent error code from the hashing functions
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       5/22/2008-11:29:29
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetHashError() As Long
    GetHashError = lHashErr
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       HashErrorString
' Description:       [type_description_here]
' Created by :       Colin
' Machine    :       IBM-GN893WUKICU
' Date-Time  :       5/22/2008-11:29:29
'
' Parameters :       sText (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function HashErrorString(sText As String) As String
    Dim sTmp As String
    sTmp = sText & ": 0x" & Hex(GetHashError())
    HashErrorString = sTmp
End Function

#If False Then
Private Function HashFile(ByRef inFile As String, _
                          ByRef outHash() As Byte) As Long
    Dim FNum As Integer
    Dim FBuf() As Byte, BufLen As Long
    Dim BytesLeft As Long
    Dim HashError As Boolean
    ' Default chunk size to use when hashing file data
    Const ChunkSize As Long = 1024 ' 1Kb
    FNum = FreeFile() ' Get free file handle and open file
    Open inFile For Binary As #FNum
    ReDim FBuf(0 To ChunkSize - 1) As Byte

    ' Attempt to open default cryptographic context
    If (CryptAcquireContext(hCryptProv, vbNullString, MS_DEF_PROV, PROV_RSA_FULL, 0&) = 0) Then

        If (CryptAcquireContext(hCryptProv, vbNullString, MS_DEF_PROV, PROV_RSA_FULL, CRYPT_NEWKEYSET) = 0) Then
            Exit Function ' Failed to acquire cryptographic context
            Close #FNum
        End If
    End If

    ' Create new MD5 hash
    If (CryptCreateHash(hCryptProv, CALG_MD5, 0&, 0&, hCryptHash)) Then
        BytesLeft = LOF(FNum)
        BufLen = ChunkSize

        Do

            If (BytesLeft < ChunkSize) Then ' Last chunk
                ReDim FBuf(0 To BytesLeft - 1) As Byte
                BufLen = BytesLeft
            End If

            ' Read chunk from file
            Get #FNum, , FBuf()

            ' Add this data to the hash
            If (CryptHashData(hCryptHash, FBuf(0), BufLen, 0&) = 0) Then
                HashError = True
                Exit Do
            End If

            ' Decrement read count
            BytesLeft = BytesLeft - BufLen
        Loop While BytesLeft > 0

        If (Not HashError) Then
            BufLen = 0

            ' Get buffer length for hash
            If (CryptGetHashParam(hCryptHash, HP_HASHVAL, ByVal 0&, BufLen, 0&)) Then
                ReDim FBuf(0 To BufLen - 1) As Byte

                If (CryptGetHashParam(hCryptHash, HP_HASHVAL, FBuf(0), BufLen, 0&)) Then
                    HashFile = BufLen ' Return final hash buffer
                    outHash = FBuf
                End If
            End If
        End If

        ' Done with hash object
        Call CryptDestroyHash(hCryptHash)
    End If

    ' Done with provider
    Call CryptReleaseContext(hCryptProv, 0&)
    Close #FNum
End Function

Private Function GetFileHash(ByRef inFile As String) As String
    Dim Hash() As Byte, HashLen As Long, LoopHash As Long
    ' Perform hash on file
    HashLen = HashFile(inFile, Hash())

    If (HashLen > 0) Then
        ' Allocate return buffer
        GetFileHash = String$(HashLen * 2, "0")

        For LoopHash = 0 To HashLen - 1

            If (Hash(LoopHash) < &H10) Then ' Single digit
                Mid$(GetFileHash, (LoopHash * 2) + 2, 1) = Hex$(Hash(LoopHash))
            Else ' Double digit
                Mid$(GetFileHash, (LoopHash * 2) + 1, 2) = Hex$(Hash(LoopHash))
            End If

        Next LoopHash

    End If

End Function
#End If
