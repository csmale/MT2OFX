Attribute VB_Name = "Config"
Option Explicit

' $Header: /MT2OFX/Config.bas 23    24/11/09 22:04 Colin $

Public Const AppName As String = "MT2OFX"
Public Const AppTitle As String = "MT940 to OFX Translator"

#If False Then
Public Const UseBookDate = 0
Public Const UseValueDate = 1
Public Const UseTxnDate = 2
#End If

Private IniFile As String
Const IniFileName = "MT2OFX.INI"

Public Cfg As New ProgramConfig
Public Bcfg As New BankConfig

Const IniSectionGeneral = "General"

'Const IniStructured86 = "Structured86"
'Private xStructured86 As Boolean
Public Const IniBankName = "BankName"


'Const IniSkipEmptyMemoFields = "SkipEmptyMemoFields"
'Const IniFindPayeeFn = "FindPayee"
'Const IniFindTxnDateFn = "FindTxnDate"
'Const IniFindServerTimeFn = "FindServerTime"
'Const IniIntuitBankID = "IntuitBankID"
'Const IniScriptFile = "ScriptFile"

Const IniSectionAccountTypes = "AccountTypes"
Const IniDefault = "Default"
Private xDefaultAccountType As String
Const IniAccountTypeDefault = "CHECKING"

Const IniSectionTxnTypeMap = "TxnTypeMap"
Const IniMapDefault = "OTHER"

Public Const IniSectionBankRules = "BankRules"
Public Const IniBankRulePrefix = "Bank"

Public Const IniSectionText = "Text"
Public Const IniTextScriptPrefix = "Script"
Public Const IniSectionTextExtension = "TextExtensions"

Const IniCookie = "XyZZy"

Public Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpDefault As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSectionNames Lib "kernel32" _
    Alias "GetPrivateProfileSectionNamesA" _
    (ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long

Private xDecSep As String
Private Declare Function GetLocaleInfo Lib "unicows.dll" _
    Alias "GetLocaleInfoW" ( _
    ByVal Locale As Long, _
    ByVal LCType As Long, _
    ByVal lpLCData As Long, _
    ByVal cchData As Long) As Long
Private Const LOCALE_SDECIMAL = &HE
Private Const LOCALE_USER_DEFAULT = &H400
Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "User32" () As Long
Public Const SW_NORMAL = 1

Public Const SE_ERR_FNF = 2
Public Const SE_ERR_PNF = 3
Public Const SE_ERR_ACCESSDENIED = 5
Public Const SE_ERR_OOM = 8
Public Const ERROR_BAD_FORMAT = 11&
Public Const SE_ERR_SHARE = 26
Public Const SE_ERR_ASSOCINCOMPLETE = 27
Public Const SE_ERR_DDETIMEOUT = 28
Public Const SE_ERR_DDEFAIL = 29
Public Const SE_ERR_DDEBUSY = 30
Public Const SE_ERR_NOASSOC = 31
Public Const SE_ERR_DLLNOTFOUND = 32

Private Sub MakeFileName()
    If IniFile = "" Then
' 20061129 CS: take ini file from program config (assume this has been loaded already!)
        IniFile = Cfg.IniFileName
    End If
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetIniSectionList
' Description:       returns an array containing the section names
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-17:29:16
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetIniSectionList() As Variant
    Dim vTmp As Variant
    Dim sTmp As String
    Dim bSize As Long
    MakeFileName
    bSize = 4000
    sTmp = String$(bSize, vbNullChar)
    bSize = GetPrivateProfileSectionNames(sTmp, bSize, IniFile)
    If bSize <= 0 Then
        GetIniSectionList = Array()
        Debug.Assert False
        Exit Function
    End If
    sTmp = Left$(sTmp, bSize - 1)
    vTmp = Split(sTmp, vbNullChar)
    GetIniSectionList = vTmp
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetIniKeyList
' Description:       returns list of keys within a section
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-17:34:13
'
' Parameters :       sSection (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetIniKeyList(sSection As String) As Variant
    Dim vTmp As Variant
    Dim sTmp As String
    Dim bSize As Long
    MakeFileName
    bSize = 4000
    sTmp = String$(bSize, vbNullChar)
    bSize = GetPrivateProfileString(sSection, vbNullString, "", sTmp, bSize, IniFile)
    If bSize <= 0 Then
        GetIniKeyList = Array()
'        Debug.Assert False
        Exit Function
    End If
    sTmp = Left$(sTmp, bSize - 1)
    vTmp = Split(sTmp, vbNullChar)
    GetIniKeyList = vTmp
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetBankSections
' Description:       Returns a list of bank sections
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       10/12/2003-21:59:28
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetBankSections() As Variant
    Dim vTmp As Variant
    Dim sTmp As String
    Dim bSize As Long
    MakeFileName
    vTmp = GetIniSectionList()
    Dim i As Long
    bSize = 0
    For i = 0 To UBound(vTmp)
        sTmp = vTmp(i)
        sTmp = GetMyString(sTmp, IniBankName, "")
        If sTmp = "" Or vTmp(i) = "*" Then  ' ignore old default
            vTmp(i) = ""
        Else
            bSize = bSize + 1
        End If
    Next
    If bSize = 0 Then
        GetBankSections = Array()
        Exit Function
    End If
    Dim j As Long
    Dim vTmp2 As Variant
    ReDim vTmp2(1 To bSize)
    j = 1
    For i = 0 To UBound(vTmp)
        If vTmp(i) <> "" Then
            vTmp2(j) = vTmp(i)
            j = j + 1
        End If
    Next
    GetBankSections = vTmp2
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ReadIniString
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/02/2004-22:52:51
'
' Parameters :       sSection (String)
'                    sKey (String)
'                    sDefault (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ReadIniString(sFile As String, sSection As String, sKey As String, _
    Optional sDefault As String = "") As String
    Dim sTmp As String
    Dim lTmp As Long
    sTmp = String$(255, vbNullChar)
    lTmp = GetPrivateProfileString(sSection, sKey, _
        sDefault, sTmp, 255, sFile)
    If lTmp > 0 Then
        sTmp = Left$(sTmp, lTmp)
    Else
        sTmp = ""
    End If
    ReadIniString = sTmp
End Function
Public Function GetMyString(sSection As String, sKey As String, _
    sDefault As String) As String
    MakeFileName
    GetMyString = ReadIniString(IniFile, sSection, sKey, sDefault)
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       WriteIniString
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/02/2004-22:54:54
'
' Parameters :       sSection (String)
'                    sKey (String)
'                    sValue (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub WriteIniString(sFile As String, sSection As String, sKey As String, _
    sValue As String)
    WritePrivateProfileString sSection, sKey, sValue, sFile
End Sub

Public Sub PutMyString(sSection As String, sKey As String, _
    sValue As String)
    MakeFileName
    WriteIniString IniFile, sSection, sKey, sValue
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DeleteMyString
' Description:       Delete an entry from the ini file
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       17/12/2003-21:37:54
'
' Parameters :       sSection (String)
'                    sKey (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub DeleteMyString(sSection As String, sKey As String)
    MakeFileName
    WritePrivateProfileString sSection, sKey, vbNullString, IniFile
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DeleteMySection
' Description:       Delete a whole (bank) section
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/01/2004-14:53:34
'
' Parameters :       sSection (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub DeleteMySection(sSection As String)
    MakeFileName
    WritePrivateProfileString sSection, vbNullString, vbNullString, IniFile
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ClearMySection
' Description:       deletes all entries from a section but leaves the
'                    section header so new entries will go in the right place
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       11/03/2005-13:37:06
'
' Parameters :       sSection (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub ClearMySection(sSection As String)
    Dim vArr As Variant
    Dim i As Integer
    vArr = GetIniKeyList(sSection)
    For i = LBound(vArr) To UBound(vArr)
        DeleteMyString sSection, CStr(vArr(i))
    Next
End Sub
Public Function GetConfigLang() As Long
    GetConfigLang = CLng(GetMyString(IniSectionGeneral, "LCID", "0"))
End Function

Public Function LoadBankSettingsEx(bc As BankConfig, sBank As String) As Boolean
    LoadBankSettingsEx = bc.Load(sBank)
End Function
Public Function LoadBankSettings(sBank As String) As Boolean
    Dim bTmp As Boolean
    bTmp = Bcfg.Load(sBank)
    If Not Bcfg.Load(sBank) Then
        bTmp = Bcfg.Load("*")
    End If
    LoadBankSettings = bTmp
End Function
' 20041217 CS Remove blank function names from ini-file (depracated anyway)
Public Function SaveBankSettingsEx(bc As BankConfig, sSection As String) As Boolean
    SaveBankSettingsEx = bc.Save(sSection)
End Function
Public Function SaveBankSettings() As Boolean
    SaveBankSettings = SaveBankSettingsEx(Bcfg, "*" & Bcfg.IDString)
End Function
Public Function AccountType(Account As String) As String
    Dim sTmp As String
    Dim iTmp As Long
    sTmp = GetMyString(IniSectionAccountTypes, Account, IniCookie)
    If sTmp = IniCookie Then    ' no explicit type
        If xDefaultAccountType = "" Then
            sTmp = GetMyString(IniSectionAccountTypes, IniDefault, _
                IniAccountTypeDefault)
            xDefaultAccountType = sTmp
        Else
            sTmp = xDefaultAccountType
        End If
    End If
    AccountType = sTmp
End Function
' 20060806 CS: now handles stmt.accttype which overrides old system
Public Function OFCAccountType(Account As String, AcctType As String) As String
    Dim sTmp As String
    Dim iTmp As Long
    sTmp = AcctType
    If Len(sTmp) = 0 Then sTmp = AccountType(Account)
    Select Case sTmp
    Case "CHECKING"
        sTmp = "0"
    Case "SAVINGS"
        sTmp = "1"
    Case "CREDITCARD"   ' NB: Not OFX, internal value only
        sTmp = "2"
    Case "MONEYMRKT"
        sTmp = "3"
    Case "CREDITLINE"
        sTmp = "4"
    Case Else
        sTmp = "7"
    End Select
    OFCAccountType = sTmp
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       DoImport
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/02/2005-10:52:10
'
' Parameters :       sFile (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function DoImport(sFile As String, bNoGUI As Boolean) As Boolean
    Dim sCmd As String
    Dim sTmp As String
    Dim iOpt As Long
    Dim vArr As Variant
    Dim sExt As String
    DoImport = False
' 20090210 CS: make extension case-insensitive
    sExt = UCase$(GetExtension(sFile))
    Select Case sExt
    Case "OFX"
        sCmd = Cfg.OFXExportTo
    Case "OFC"
        sCmd = Cfg.OFCExportTo
    Case "QIF"
        sCmd = Cfg.QIFExportTo
    Case "QFX"
        sCmd = Cfg.QFXExportTo
    Case Else
        sCmd = ""
    End Select
    If Len(sCmd) = 0 Then
        iOpt = 1 ' do default action from windows
    Else
        vArr = Split(sCmd, ",")
        If IsNumeric(vArr(0)) Then
            iOpt = CLng(vArr(0))
            If iOpt < 0 Or iOpt > 2 Then
                iOpt = 1
            End If
            If iOpt = 2 Then
                If UBound(vArr) > 0 Then
                    sCmd = vArr(1)
                Else
                    iOpt = 1
                End If
            End If
        Else
            iOpt = 1
        End If
    End If
    
    If Not bNoGUI Then
        If iOpt > 0 Then
            If MyMsgBox(LoadResStringL(110), _
                vbInformation + vbOKCancel) <> vbOK Then
                Exit Function
            End If
        Else
            MyMsgBox LoadResStringL(111), vbInformation + vbOKOnly
        End If
    End If
    Select Case iOpt
    Case 0
        DoImport = True ' no action required
    Case 1              ' windows default action
        DoImport = DoDefaultFileAction(sFile)
    Case 2              ' user defined program
        DoImport = RunExternalProgram(sCmd, sFile)
    End Select
End Function

Public Function DoDefaultFileAction(sFile As String) As Boolean
    DoDefaultFileAction = RunExternalProgram(sFile, vbNullString)
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       RunExternalProgram
' Description:
' Created by :       Colin Smale
' Machine    :       IBM-FWQ8A7OCQYF
' Date-Time  :       12/02/2005-23:04:56
'
' Parameters :       sProg (String)
'                    sParams (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function RunExternalProgram(sProg As String, sParams As String) As Boolean
    Dim lTmp As Long
    Dim sTmp As String
' pass output file to the given program
    lTmp = ShellExecute(GetDesktopWindow(), vbNullString, _
        sProg, sParams, vbNullString, SW_NORMAL)
    If lTmp <= 32 Then
        Select Case lTmp
        Case SE_ERR_NOASSOC, SE_ERR_ASSOCINCOMPLETE
            sTmp = GetExtension(sProg)
            MyMsgBox GetString(128, sTmp), vbCritical + vbOKOnly
        Case Else
' 20050318 CS: Error(lTmp) doesn't give the right message - it is a Windows
' error and not a VB error. ReturnAPIError finds the system error string
            sTmp = GetString(112, lTmp, sProg, sParams, ReturnAPIError(lTmp))
            LogMessage True, True, sTmp
        End Select
    End If
    RunExternalProgram = (lTmp > 32)
End Function

Public Function OFXTxnType(SwiftTxnType As String, Amount As Double) As String
    Dim sTmp As String
    Dim iTmp As Long
    If SwiftTxnType = "" Then
        OFXTxnType = ""
        Exit Function
    End If
    sTmp = GetMyString(IniSectionTxnTypeMap, SwiftTxnType, IniCookie)
    If sTmp = IniCookie And Len(SwiftTxnType) > 0 Then
        WritePrivateProfileString IniSectionTxnTypeMap, SwiftTxnType, _
        IniMapDefault, IniFile
        sTmp = IniMapDefault
    End If
    iTmp = InStr(sTmp, ",")
    If iTmp > 0 Then
        If Amount < 0 Then
            sTmp = Mid$(sTmp, iTmp + 1)
        Else
            sTmp = Left$(sTmp, iTmp - 1)
        End If
    End If
    OFXTxnType = sTmp
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       OFXTxnType2OFC
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       19/02/2004-21:07:41
'
' Parameters :       sOFXTxnType (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function OFXTxnType2OFC(sOFXTxnType As String) As String
    Dim iTmp As Long
    Select Case sOFXTxnType
    Case "CREDIT": iTmp = 0
    Case "DEBIT": iTmp = 1
    Case "INT": iTmp = 2
    Case "DIV": iTmp = 3
    Case "FEE": iTmp = 4
    Case "SRVCHG": iTmp = 4
    Case "DEP": iTmp = 5
    Case "ATM": iTmp = 6
    Case "XFER": iTmp = 7
    Case "CHECK": iTmp = 8
    Case "PAYMENT": iTmp = 9
    Case "CASH": iTmp = 10
    Case "DIRECTDEP": iTmp = 5
    Case "DIRECTDEBIT": iTmp = 11
    Case "REPEATPMT": iTmp = 9
    Case "OTHER": iTmp = 12
    Case Else: iTmp = 12
    End Select

'Possible values here:
'Value Description
'0               Credit, positively affects account balance
'1               Debit, negatively affects account balance
'2               Interest
'3               Dividend
'4               Service charge
'5               Deposit
'6               ATM withdrawal
'7               Transfer
'8               CHECK
'9               Electronic PAYMENT
'10              CASH withdrawal
'11              Direct debit of paycheck
'12              OTHER
    
    OFXTxnType2OFC = CStr(iTmp)
End Function

Public Function IdentifyBank(oFile As InputFile) As String
    Dim iBank As Integer
    Dim vArr As Variant
    Dim sTmp As String
    Dim i As Integer
    Dim iPos As Long
    
    IdentifyBank = ""
    iBank = 1
    iPos = oFile.Pos
nextbank:
    sTmp = GetMyString(IniSectionBankRules, _
        IniBankRulePrefix & CStr(iBank), "")
    If sTmp = "" Then Exit Function
    vArr = Split(sTmp, ",")
    If Not IsArray(vArr) Then GoTo gonext
    If UBound(vArr) <> 2 Then GoTo gonext
    If Not IsNumeric(vArr(1)) Then GoTo gonext
    oFile.Rewind
    For i = 1 To CInt(vArr(1))
        If oFile.AtEOF Then GoTo gonext
        sTmp = oFile.ReadLine
        If sTmp Like vArr(2) Then
            IdentifyBank = vArr(0)
            oFile.Rewind
            Exit Function
        End If
    Next
    
gonext:
    iBank = iBank + 1
    GoTo nextbank
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ScriptModuleName
' Description:       Returns the module name to be used for this script
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       13/01/2004-13:24:08
'
' Parameters :       sScript (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ScriptModuleName(sScript As String) As String
    Dim iTmp As Long
    Dim sMod As String
    iTmp = InStrRev(sScript, "\")
    If iTmp > 0 Then
        sMod = Mid$(sScript, iTmp + 1)
    Else
        sMod = sScript
    End If
    iTmp = InStr(sMod, ".")
    If iTmp > 0 Then
        sMod = Left$(sMod, iTmp - 1)
    End If
    ScriptModuleName = sMod
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       ScriptForInputFile
' Description:       looks up input script for the given file
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       23/11/2003-22:33:43
'
' Parameters :       Extension (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ScriptForInputFile() As String
    Dim sExt As String
    Dim sTmp As String
    Dim bRet As Boolean
    Dim sScript As String, sMod As String
    Dim iTmp As Integer
    Dim i As Integer
    Dim vArr As Variant
    Dim bActive As Boolean
    
' see if there's an explicit map to a text processing script
    sExt = UCase$(GetExtension(Session.FileIn))
    If sExt <> "" Then
        sTmp = GetMyString(IniSectionTextExtension, sExt, "")
        If sTmp <> "" Then
            ScriptForInputFile = sTmp
            Session.BankID = ScriptModuleName(sTmp)
            Exit Function
        End If
    End If
    
' see if it's an MT940 extension - if so, bale out
    If InExtensionList(Cfg.AutoMT940, sExt) Then
        LogMessage False, True, "File extension " & sExt & " is in AutoMT940"
        ScriptForInputFile = ""
        Exit Function
    End If
    
    i = 1
    If Not InitialiseScripting() Then
        Exit Function
    End If
    sTmp = GetMyString(IniSectionText, IniTextScriptPrefix & CStr(i), "")
    Do While Len(sTmp) > 0
' 20050210 CS: added script active status
        vArr = Split(sTmp, ",")
        If UBound(vArr) > 0 Then
            If IsNumeric(vArr(0)) Then
                bActive = (CLng(vArr(0)) > 0)
            Else
                bActive = True
            End If
            sScript = vArr(1)
        Else
            sScript = sTmp
            bActive = True
        End If
        If bActive Then
            sMod = ScriptModuleName(sScript)
            sScript = FindScript(sScript)
' 20060125 CS: FindScript now checks if the file exists
            If sScript = "" Then
                LogMessage False, True, "Unable to find script file: " & sScript, ""
                GoTo trynext
            End If
            LogMessage False, True, "Loading script file: " & sScript, ""
            LogMessage False, True, "Module: " & sMod, ""
' 20070104 CS: clear abort flag so we keep on trying
            GetScriptEnv().AbortRequested = False
            If Not ScriptInit(sMod, sScript) Then
' message 118 is "Processing aborted" - we must be able to improve on this!
                LogMessage True, True, GetString(118), AppName
' 20070104 CS: If script cannot be loaded, there's no reason not to try other scripts!
                GoTo trynext
            End If
' call RecogniseTextFile in this script
            Session.InputFile.Rewind
            Session.InputFile.CodePage = CP_UNKNOWN
            If CallScriptFunctionRet(sMod, "RecogniseTextFile", bRet) Then
                If bRet Then
                    LogMessage False, True, "File recognised by " & sScript, ""
                    ' script recognises the input file
                    ScriptForInputFile = sScript
                    Session.BankID = sMod
                    Exit Do
                End If
            Else
                LogMessage False, True, "Error calling RecogniseTextFile()", ""
            End If
        End If
trynext:
        i = i + 1
        sTmp = GetMyString(IniSectionText, IniTextScriptPrefix & CStr(i), "")
    Loop
' nothing found...
    If Len(sTmp) = 0 Then
        ScriptForInputFile = ""
        Session.BankID = ""
        LogMessage False, True, "No script matched " & Session.FileIn, ""
    End If
End Function

Public Function XMLEncodingForCodepage(iCP As Long) As String
    Dim sTmp As String
    sTmp = GetRegString(HKEY_CLASSES_ROOT, "Mime\Database\Codepage\" & CStr(iCP), "BodyCharset")
    XMLEncodingForCodepage = sTmp
End Function

