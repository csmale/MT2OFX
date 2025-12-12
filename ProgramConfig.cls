VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProgramConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : ProgramConfig
'    Project    : MT2OFX
'
'    Description:
'
'    Modified   : $Author: Colin $ $Date: 15/11/10 0:16 $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/ProgramConfig.cls 31    15/11/10 0:16 Colin $"
' $History: ProgramConfig.cls $
' 
' *****************  Version 31  *****************
' User: Colin        Date: 15/11/10   Time: 0:16
' Updated in $/MT2OFX
'
' *****************  Version 30  *****************
' User: Colin        Date: 15/06/09   Time: 19:25
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 29  *****************
' User: Colin        Date: 25/11/08   Time: 22:23
' Updated in $/MT2OFX
' moving vss server!
'
' *****************  Version 27  *****************
' User: Colin        Date: 20/04/08   Time: 10:06
' Updated in $/MT2OFX
' For 3.5 beta 1
'
' *****************  Version 26  *****************
' User: Colin        Date: 7/12/06    Time: 15:07
' Updated in $/MT2OFX
' MT2OFX Version 3.5.2
'
' *****************  Version 23  *****************
' User: Colin        Date: 2/11/05    Time: 23:03
' Updated in $/MT2OFX
' V3.4 beta 1
'
' *****************  Version 22  *****************
' User: Colin        Date: 11/06/05   Time: 19:33
' Updated in $/MT2OFX
'
' *****************  Version 21  *****************
' User: Colin        Date: 6/03/05    Time: 23:42
' Updated in $/MT2OFX
'</CSCC>

Private IniFile As String
Const MyIniFileName = "MT2OFX.INI"
Const IniCookie = "XyZZy"

Const IniSectionGeneral As String = "General"

Public EnableJavaScript As Boolean
Const IniEnableJavaScript As String = "EnableJavaScript"
Const IniEnableJavaScriptDefault As Boolean = False

Public MemoDelimiter As String
Const IniMemoDelimiter As String = "MemoDelimiter"
Const IniMemoDelimDefault As String = ";"

Public OutputFileType As String
Const IniOutputFileType As String = "OutputFileType"
Const IniOutputFileTypeDefault As String = "OFX"

Public HomePage As String
Const IniHomePage As String = "ApplicationHomePage"
Const IniHomePageDefault As String = "http://www.xs4all.nl/~csmale/mt2ofx/"

Public PayeeCase As Long
Const IniPayeeCase As String = "PayeeCase"
Const IniPayeeCaseDefault As Long = 0&

Public AutoBrowse As Boolean
Const IniAutoBrowse = "AutoBrowse"
Const IniAutoBrowseDefault As Boolean = False

'Public AutoStartImport As Boolean
'Const IniAutoStartImport As String = "StartImportAfterConversion"
'Const IniAutoStartImportDefault As Boolean = False

Public PromptForOutput As OutfileEnum
Const IniPromptForOutput As String = "PromptForOutput"
Const IniPromptForOutputDefault As Integer = OUTFILE_AUTO

Public SaveClipboardOutput As Boolean
Const IniSaveClipboardOutput As String = "SaveClipOutput"
Const IniSaveClipboardOutputDefault As Boolean = True

Public BookDateMode As Long
Const IniBookDateMode As String = "BookDateMode"
Public Enum BookDateModes
    bdmBookDate = 0
    bdmValueDate = 1
    bdmTransDate = 2
End Enum
Const IniBookDateModeDefault As Long = bdmBookDate
Const IniBookDateModeMax As Long = bdmTransDate

Public CompressSpaces As Boolean
Const IniCompressSpaces As String = "CompressSpaces"
Const IniCompressSpacesDefault As Boolean = True

Public ScriptDebugLevel As Long
Const IniScriptDebugLevel As String = "ScriptDebugLevel"
Const IniScriptDebugLevelDefault As Integer = 0

' 20100202 CS: Introduce LogLevel to govern verbosity
Public Enum logLevels
    logNONE = 0&
    logNORMAL = 1&
    logVERBOSE = 2&
    logDEBUG = 3&
End Enum
Public LogLevel As logLevels
Const IniLogLevel As String = "LogLevel"
Const IniLogLevelDefault As Long = logNORMAL

Public LogFile As String
Const IniLogFile As String = "LogFile"
Const IniLogFileDefault As String = "Log.txt"

Public TxnDumpFile As String
Const IniTxnDump = "TxnDumpFile"
Const IniTxnDumpDefault = ""

Public ScriptTimeout As Long
Const IniScriptTimeout As String = "ScriptTimeout"
Const IniScriptTimeoutDefault As Long = 5000&

Const IniSectionQIF = "QIF"

Public QifDateFormat As Long
Const IniQifDateFormat As String = "DateFormat"
Const IniQifDateFormatDefault As Long = 3&

' 20050831 CS: Added QIF custom date format
Public QifCustomDateFormat As String
Const IniQifCustomDateFormat As String = "CustomDateFormat"
Const IniQifCustomDateFormatDefault As String = ""

Public QifDecimalSeparator As String
Const IniQifDecimalSeparator As String = "DecimalSeparator"
Const IniQifDecimalSeparatorDefault As String = ""

' 20050112 CS: Now a script parameter
'Public QifInputCurrency As String

Public QifNoAcctHeader As Boolean
Const IniQifNoAcctHeader As String = "NoAccountHeader"
Const IniQifNoAcctHeaderDefault As Boolean = False

' 20050126 CS: Added QifMaxMemoLength
Public QifMaxMemoLength As Long
Const IniQifMaxMemoLength As String = "MaxMemoLength"
Const IniQifMaxMemoLengthDefault As Long = 0

' 20080704 CS: Added option to output amounts as "U" as well as "T"
Public QifOutputUAmount As Boolean
Const IniQifOutputUAmount As String = "OutputUAmount"
Const IniQifOutputUAmountDefault As Boolean = False

Public ScriptEditor As String
Const IniScriptEditor As String = "ScriptEditor"
Const IniScriptEditorDefault As String = "notepad.exe"

Public AutoMT940 As String
Const IniAutoMT940 As String = "AutoMT940"
Const IniAutoMT940Default As String = "940,SWI,STA"

Public AutoText As String
Const IniAutoText As String = "AutoText"
Const IniAutoTextDefault As String = "ASC,CSV,TXT"

Public LastInputDirectory As String
Const IniLastInputDirectory As String = "LastInputDirectory"
Const IniLastInputDirectoryDefault As String = ""

Public LastInputExtension As String
Const IniLastInputExtension As String = "LastInputExtension"
Const IniLastInputExtensionDefault As String = ""

Public EnableSounds As Boolean
Const IniEnableSounds As String = "EnableSounds"
Const IniEnableSoundsDefault As Boolean = False

Public PayeeMapFile As String
Const IniPayeeMapFile As String = "PayeeMapFile"
Const IniPayeeMapFileDefault As String = "PayeeMap.dat"

Public PayeeMapIgnoreCase As Boolean
Const IniPayeeMapIgnoreCase As String = "PayeeMapIgnoreCase"
Const IniPayeeMapIgnoreCaseDefault As Boolean = False

Public TestPayee As String
Const IniTestPayee As String = "TestPayee"
Const IniTestPayeeDefault As String = ""

Public TestMemo As String
Const IniTestMemo As String = "TestMemo"
Const IniTestMemoDefault As String = ""

Public DefaultPayee As String
Const IniDefaultPayee As String = "DefaultPayee"
' CS 20050109: Allow payee name to be empty (now omitted in OFC/OFX output)
'Const IniDefaultPayeeDefault As String = "*** Unknown ***"
Const IniDefaultPayeeDefault As String = ""

Public GenerateCheckNum As Boolean
Const IniGenerateCheckNum As String = "GenerateCheckNum"
Const IniGenerateCheckNumDefault As Boolean = False

Public OFXVersion As Integer
Const IniOFXVersion As String = "OFXVersion"
Const IniOFXVersionDefault As Integer = 102

' 20080104 CS: Add configurable decimal separator for OFX
Public OFXDecimal As String
Const IniOFXDecimal As String = "OFXDecimal"
Const IniOFXDecimalDefault As String = "."

' 20081211 CS: Force real balance date
Public OFXForceRealBalanceDate As Boolean
Const IniOFXForceRealBalanceDate As String = "OFXForceRealBalanceDate"
Const IniOFXForceRealBalanceDateDefault As Boolean = False

Public NoScriptingVersionCheck As Boolean
Const IniNoScriptingVersionCheck As String = "NoScriptingVersionCheck"
Const IniNoScriptingVersionCheckDefault As Boolean = False

' 20050208 CS: support message stuff
Public SptLogFileLines As Integer
Const IniSptLogFileLines As String = "SupportLogFileLines"
Const IniSptLogFileLinesDefault As Integer = 20

' 20050211 CS: configurable output programs
Public OFXExportTo As String
Const IniOFXExportTo As String = "OFXExportTo"
Public OFCExportTo As String
Const IniOFCExportTo As String = "OFCExportTo"
Public QIFExportTo As String
Const IniQIFExportTo As String = "QIFExportTo"
Public QFXExportTo As String
Const IniQFXExportTo As String = "QFXExportTo"
Const IniExportToDefault As String = ""

' 20091203 CS: add OFX splits
Public Enum OFXSplitHandlingEnum
    ofxsMainOnly = 0
    ofxsSplitsOnly = 1
    ofxsMainAndSplits = 2
End Enum
Public OFXSplitHandling As OFXSplitHandlingEnum
Const IniOFXSplitHandling As String = "OFXSplitHandling"
Const IniOFXSplitHandlingDefault As Integer = ofxsMainOnly

' 20050212 CS: delay before deleting temp output file (in ms)
Public TempFileDelay As Long
Const IniTempFileDelay As String = "TempFileDelay"
Const IniTempFileDelayDefault As Long = 1000

' 20050305 CS: added external resource file for localisation
Public ResFile As String
Const IniResFile As String = "ResFile"
Const IniResFileDefault As String = ""

' 20050531 CS: added NoSortTransactions
Public NoSortTxns As Boolean
Const IniNoSortTxns As String = "NoSortTxns"
Const IniNoSortTxnsDefault As Boolean = False

' 20060821 CS: added DebugDBCS
Public DebugDBCS As Boolean
Const IniDebugDBCS As String = "DebugDBCS"
Const IniDebugDBCSDefault As Boolean = False

' 20060825 CS: added SuppressEmptyStatements
Public SuppressEmptyStatements As Boolean
Const IniSuppressEmpty As String = "SuppressEmptyStatements"
Const IniSuppressEmptyDefault As Boolean = False

' 20061015
Public ProgressBarDelay As Long
Const IniProgressBarDelay As String = "ProgressBarDelay"
Const IniProgressBarDelayDefault As Long = 1500

' 20070208 CS: Add QuickStart hook for first run after install
Public ShowQuickStart As Boolean
Const IniShowQuickStart As String = "ShowQuickStart"
Const IniShowQuickStartDefault As Boolean = False

' 20071017 CS: Add output code page handling
Public OutputCodePage As Long
Const IniOutputCodePage As String = "OutputCodePage"
Const IniOutputCodePageDefault As Long = CP_ACP

Public OutputBOM As Boolean
Const IniOutputBOM As String = "OutputBOM"
Const IniOutputBOMDefault As Boolean = False

Public Enum LineEndings
    leCRLF = 0
    leCR = 1
    leLF = 2
    leLFCR = 3
End Enum
Public LineEnding As LineEndings
Const IniLineEnding As String = "LineEnding"
Const IniLineEndingDefault As Long = leCRLF

Public UseOldFITID As Boolean
Const IniUseOldFITID As String = "UseOldFITID"
Const IniUseOldFITIDDefault As Boolean = False

Public QFXForceCurrency As String
Const IniQFXForceCurrency As String = "QFXForceCurrency"
Const IniQFXForceCurrencyDefault As String = ""

Public NoConfirmImport As Boolean
Const IniNoConfirmImport As String = "NoConfirmImport"
Const IniNoConfirmImportDefault As Boolean = False

Public Changed As Boolean

' 20061129 CS: Vista compatibility: save user app data path here
Private xAppDataPath As String
Private xAppPath As String
Private xMyDocumentsPath As String
Private xScriptPath As String
Private xResourcePath As String
Private xHelpPath As String
Private xTempPath As String
Private xCustomOutputPath As String

Public Property Get AppPath() As String
    AppPath = xAppPath
End Property
Public Property Get AppDataPath() As String
    AppDataPath = xAppDataPath
End Property
Public Property Get MyDocumentsPath() As String
    MyDocumentsPath = xMyDocumentsPath
End Property
Public Property Get ScriptPath() As String
    ScriptPath = xScriptPath
End Property
Public Property Get ResourcePath() As String
    ResourcePath = xResourcePath
End Property
Public Property Get HelpPath() As String
    HelpPath = xHelpPath
End Property
Public Property Get TempPath() As String
    TempPath = xTempPath
End Property
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       CustomOutputPath
' Description:
' Created by :       Colin Smale
' Machine    :       L3CCG6P-6474B84
' Date-Time  :       18/04/2010-22:47:20
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get CustomOutputPath() As String
    CustomOutputPath = xCustomOutputPath
End Property
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetIniFileName
' Description:       returns the full name of the INI file to be used
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/02/2004-22:05:43
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get IniFileName() As String
    If IniFile = "" Then
        IniFile = xAppDataPath & "\" & MyIniFileName
        If Dir(IniFile, vbNormal) = "" Then
            IniFile = xAppPath & "\" & MyIniFileName
        End If
    End If
    IniFileName = IniFile
End Property

Friend Property Let IniFileName(IniFileName As String)
    If IniFileName = "" Or Dir(IniFileName, vbNormal) = "" Then
        IniFile = ""
    Else
        IniFile = IniFileName
    End If
End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetString
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/02/2004-22:01:59
'
' Parameters :       sSection (String)
'                    sKey (String)
'                    sDefault (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function GetString(sSection As String, sKey As String, sDefault As String) As String
    Dim sTmp As String
    Dim iTmp As Long
    sTmp = ReadIniString(IniFile, sSection, sKey, IniCookie)
    If sTmp = IniCookie Then
        WriteIniString IniFile, sSection, sKey, sDefault
        GetString = sDefault
    Else
        GetString = sTmp
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetBool
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/02/2004-22:00:22
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Function GetBool(sSection As String, sKey As String, bDefault As Boolean) As Boolean
    Dim sTmp As String
    sTmp = ReadIniString(IniFile, sSection, sKey, IIf(bDefault, "1", "0"))
    If sTmp = "" Then sTmp = IIf(bDefault, "1", "0")
    If IsNumeric(sTmp) Then
        GetBool = (CLng(sTmp) <> 0)
    Else
        GetBool = bDefault
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       GetNumber
' Description:       [type_description_here]
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/02/2004-22:01:44
'
' Parameters :       sSection (String)
'                    sKey (String)
'                    iDefault (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function GetNumber(sSection As String, sKey As String, iDefault As Long) As Long
    Dim sTmp As String
    sTmp = ReadIniString(IniFile, sSection, sKey, CStr(iDefault))
    If sTmp = "" Then sTmp = CStr(iDefault)
    If IsNumeric(sTmp) Then
        GetNumber = CLng(sTmp)
    Else
        GetNumber = iDefault
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       PutString
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/02/2004-23:49:17
'
' Parameters :       sSection (String)
'                    sKey (String)
'                    sVal (String)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub PutString(sSection As String, sKey As String, sVal As String)
    WriteIniString IniFile, sSection, sKey, sVal
End Sub
'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       PutNumber
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/02/2004-23:48:42
'
' Parameters :       sSection (String)
'                    sKey (String)
'                    iVal (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub PutNumber(sSection As String, sKey As String, ByVal iVal As Long)
    PutString sSection, sKey, CStr(iVal)
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       PutBool
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/02/2004-23:50:07
'
' Parameters :       sSection (String)
'                    sKey (String)
'                    bVal (Boolean)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub PutBool(sSection As String, sKey As String, bVal As Boolean)
    PutString sSection, sKey, IIf(bVal, "1", "0")
End Sub

Public Function Load() As Boolean
    Dim sTmp As String
    sTmp = IniFileName
    EnableJavaScript = GetBool(IniSectionGeneral, IniEnableJavaScript, IniEnableJavaScriptDefault)
    MemoDelimiter = GetString(IniSectionGeneral, IniMemoDelimiter, IniMemoDelimDefault)
    OutputFileType = GetString(IniSectionGeneral, IniOutputFileType, IniOutputFileTypeDefault)
    HomePage = GetString(IniSectionGeneral, IniHomePage, IniHomePageDefault)
    PayeeCase = GetNumber(IniSectionGeneral, IniPayeeCase, IniPayeeCaseDefault)
    AutoBrowse = GetBool(IniSectionGeneral, IniAutoBrowse, IniAutoBrowseDefault)
 '   AutoStartImport = GetBool(IniSectionGeneral, IniAutoStartImport, IniAutoStartImportDefault)
    PromptForOutput = GetNumber(IniSectionGeneral, IniPromptForOutput, IniPromptForOutputDefault)
    SaveClipboardOutput = GetBool(IniSectionGeneral, IniSaveClipboardOutput, IniSaveClipboardOutputDefault)
    BookDateMode = GetNumber(IniSectionGeneral, IniBookDateMode, IniBookDateModeDefault)
    CompressSpaces = GetBool(IniSectionGeneral, IniCompressSpaces, IniCompressSpacesDefault)
    ScriptDebugLevel = GetNumber(IniSectionGeneral, IniScriptDebugLevel, IniScriptDebugLevelDefault)
    LogFile = GetString(IniSectionGeneral, IniLogFile, IniLogFileDefault)
    TxnDumpFile = GetString(IniSectionGeneral, IniTxnDump, IniTxnDumpDefault)
    ScriptTimeout = GetNumber(IniSectionGeneral, IniScriptTimeout, IniScriptTimeoutDefault)
    ScriptEditor = GetString(IniSectionGeneral, IniScriptEditor, IniScriptEditorDefault)
    AutoMT940 = GetString(IniSectionGeneral, IniAutoMT940, IniAutoMT940Default)
    AutoText = GetString(IniSectionGeneral, IniAutoText, IniAutoTextDefault)
    EnableSounds = GetBool(IniSectionGeneral, IniEnableSounds, IniEnableSoundsDefault)
    LastInputDirectory = GetString(IniSectionGeneral, IniLastInputDirectory, IniLastInputDirectoryDefault)
    LastInputExtension = GetString(IniSectionGeneral, IniLastInputExtension, IniLastInputExtensionDefault)
    PayeeMapFile = GetString(IniSectionGeneral, IniPayeeMapFile, IniPayeeMapFileDefault)
    LogLevel = GetNumber(IniSectionGeneral, IniLogLevel, IniLogLevelDefault)
    If PathIsRelative(PayeeMapFile) Then
' 20061129 CS: changed to app data for Vista compatibility
        PayeeMapFile = Cfg.AppDataPath & "\" & PayeeMapFile
    End If
    PayeeMapIgnoreCase = GetBool(IniSectionGeneral, IniPayeeMapIgnoreCase, IniPayeeMapIgnoreCaseDefault)
    TestPayee = GetString(IniSectionGeneral, IniTestPayee, IniTestPayeeDefault)
' 20050214 CS: add TestMemo to support matching on memo field
    TestMemo = GetString(IniSectionGeneral, IniTestMemo, IniTestMemoDefault)
    DefaultPayee = GetString(IniSectionGeneral, IniDefaultPayee, IniDefaultPayeeDefault)
    GenerateCheckNum = GetBool(IniSectionGeneral, IniGenerateCheckNum, IniGenerateCheckNumDefault)
    OFXVersion = GetNumber(IniSectionGeneral, IniOFXVersion, IniOFXVersionDefault)
' 20080104 cs: add OFXDecimal
    OFXDecimal = GetString(IniSectionGeneral, IniOFXDecimal, IniOFXDecimalDefault)
    If OFXDecimal <> "." And OFXDecimal <> "," Then
        OFXDecimal = IniOFXDecimalDefault
    End If
' 20081211 CS: add force real balance date
    OFXForceRealBalanceDate = GetBool(IniSectionGeneral, IniOFXForceRealBalanceDate, IniOFXForceRealBalanceDateDefault)
' 20091203 CS: add OFX split handling
    OFXSplitHandling = GetNumber(IniSectionGeneral, IniOFXSplitHandling, IniOFXSplitHandlingDefault)
    
    NoScriptingVersionCheck = GetBool(IniSectionGeneral, IniNoScriptingVersionCheck, IniNoScriptingVersionCheckDefault)
    
    QifDateFormat = GetNumber(IniSectionQIF, IniQifDateFormat, IniQifDateFormatDefault)
' 20050831 CS: Add QifCustomDateFormat
    QifCustomDateFormat = GetString(IniSectionQIF, IniQifCustomDateFormat, IniQifCustomDateFormatDefault)
    QifDecimalSeparator = GetString(IniSectionQIF, IniQifDecimalSeparator, IniQifDecimalSeparatorDefault)
    QifNoAcctHeader = GetBool(IniSectionQIF, IniQifNoAcctHeader, IniQifNoAcctHeaderDefault)
' 20050126 CS: Add QifMaxMemoLength
    QifMaxMemoLength = GetNumber(IniSectionQIF, IniQifMaxMemoLength, IniQifMaxMemoLengthDefault)
' 20080704 CS: Add UseUAmount
    QifOutputUAmount = GetBool(IniSectionQIF, IniQifOutputUAmount, IniQifOutputUAmountDefault)
    
' 20050208 CS: support key stuff
    SptLogFileLines = GetNumber(IniSectionGeneral, IniSptLogFileLines, IniSptLogFileLinesDefault)

' 20050211 CS: configurable output programs
    OFXExportTo = GetString(IniSectionGeneral, IniOFXExportTo, IniExportToDefault)
    OFCExportTo = GetString(IniSectionGeneral, IniOFCExportTo, IniExportToDefault)
    QIFExportTo = GetString(IniSectionGeneral, IniQIFExportTo, IniExportToDefault)
    QFXExportTo = GetString(IniSectionGeneral, IniQFXExportTo, IniExportToDefault)

' 20050212 CS: delay before deleting temp output file
    TempFileDelay = GetNumber(IniSectionGeneral, IniTempFileDelay, IniTempFileDelayDefault)

' 20050305 CS: added external resource file
    ResFile = GetString(IniSectionGeneral, IniResFile, IniResFileDefault)

' 20050531 CS: added NoSortTxns
    NoSortTxns = GetBool(IniSectionGeneral, IniNoSortTxns, IniNoSortTxnsDefault)
    
' 20060821 CS: added DebugDBCS
    DebugDBCS = GetBool(IniSectionGeneral, IniDebugDBCS, IniDebugDBCSDefault)

' 20060825 CS: added SuppressEmptyStatements
    SuppressEmptyStatements = GetBool(IniSectionGeneral, IniSuppressEmpty, IniSuppressEmptyDefault)

' 20061015 CS: Added ProgressBarDelay
    ProgressBarDelay = GetNumber(IniSectionGeneral, IniProgressBarDelay, IniProgressBarDelayDefault)

' 20070208 CS: addedShowQuickStart
    ShowQuickStart = GetBool(IniSectionGeneral, IniShowQuickStart, IniShowQuickStartDefault)
    
' 20071017 CS: added output code page
    OutputCodePage = GetNumber(IniSectionGeneral, IniOutputCodePage, IniOutputCodePageDefault)
    OutputBOM = GetBool(IniSectionGeneral, IniOutputBOM, IniOutputBOMDefault)
    LineEnding = GetNumber(IniSectionGeneral, IniLineEnding, IniLineEndingDefault)

' 20080511 CS: added UseOldFITID
    UseOldFITID = GetBool(IniSectionGeneral, IniUseOldFITID, IniUseOldFITIDDefault)

' 20080610 CS: added QFXForceCurrency
    QFXForceCurrency = GetString(IniSectionGeneral, IniQFXForceCurrency, IniQFXForceCurrencyDefault)

' 20080729 CS: force old FITID generation prior to XP/2003: older systems don't have the MSCrypt hash functions
    If Not IsXpOrLater() Then
        UseOldFITID = True
    End If

' 20080729 CS: allow prompt for import to be skipped over
    NoConfirmImport = GetBool(IniSectionGeneral, IniNoConfirmImport, IniNoConfirmImportDefault)

    Load = True
End Function

Public Function Save() As Boolean
    Dim sTmp As String
    
    PutString IniSectionGeneral, IniMemoDelimiter, MemoDelimiter
    PutString IniSectionGeneral, IniOutputFileType, OutputFileType
    PutString IniSectionGeneral, IniHomePage, HomePage
    PutNumber IniSectionGeneral, IniPayeeCase, PayeeCase
    PutBool IniSectionGeneral, IniAutoBrowse, AutoBrowse
'    PutBool IniSectionGeneral, IniAutoStartImport, AutoStartImport
    PutNumber IniSectionGeneral, IniPromptForOutput, PromptForOutput
    PutBool IniSectionGeneral, IniSaveClipboardOutput, SaveClipboardOutput
    PutNumber IniSectionGeneral, IniBookDateMode, BookDateMode
    PutBool IniSectionGeneral, IniCompressSpaces, CompressSpaces
    PutNumber IniSectionGeneral, IniScriptDebugLevel, ScriptDebugLevel
    PutString IniSectionGeneral, IniLogFile, LogFile
    PutString IniSectionGeneral, IniTxnDump, TxnDumpFile
    PutNumber IniSectionGeneral, IniScriptTimeout, ScriptTimeout
    PutString IniSectionGeneral, IniScriptEditor, ScriptEditor
    PutString IniSectionGeneral, IniAutoMT940, AutoMT940
    PutString IniSectionGeneral, IniAutoText, AutoText
    PutBool IniSectionGeneral, IniEnableSounds, EnableSounds
' 20061129 CS: changed to app data for Vista compatibility
    If Left$(PayeeMapFile, Len(Cfg.AppDataPath) + 1) = Cfg.AppDataPath & "\" Then
        sTmp = Mid$(PayeeMapFile, Len(Cfg.AppDataPath) + 2)
    Else
        sTmp = PayeeMapFile
    End If
    PutString IniSectionGeneral, IniPayeeMapFile, sTmp
    PutBool IniSectionGeneral, IniPayeeMapIgnoreCase, PayeeMapIgnoreCase
    PutString IniSectionGeneral, IniTestPayee, TestPayee
' 20050214 CS: added TestMemo to support matching on memo field
    PutString IniSectionGeneral, IniTestMemo, TestMemo
    PutString IniSectionGeneral, IniDefaultPayee, IniDefaultPayeeDefault
    PutBool IniSectionGeneral, IniGenerateCheckNum, GenerateCheckNum
    PutNumber IniSectionGeneral, IniOFXVersion, OFXVersion
' 20081211 CS: add force real balance date
    PutBool IniSectionGeneral, IniOFXForceRealBalanceDate, OFXForceRealBalanceDate
' 20080104 cs: add OFXDecimal
    PutString IniSectionGeneral, IniOFXDecimal, OFXDecimal
'    PutBool IniSectionGeneral, IniNoScriptingVersionCheck, NoScriptingVersionCheck

    PutNumber IniSectionQIF, IniQifDateFormat, QifDateFormat
' 20050831 CS: Add QifCustomDateFormat
    PutString IniSectionQIF, IniQifCustomDateFormat, QifCustomDateFormat
    PutString IniSectionQIF, IniQifDecimalSeparator, QifDecimalSeparator
    PutBool IniSectionQIF, IniQifNoAcctHeader, QifNoAcctHeader
' 20050126 CS: Add QifMaxMemoLength
    PutNumber IniSectionQIF, IniQifMaxMemoLength, QifMaxMemoLength
' 20080704 CS: Add QifOutputUAmount
    PutBool IniSectionQIF, IniQifOutputUAmount, QifOutputUAmount
' 20050208 CS: support key stuff
    PutNumber IniSectionGeneral, IniSptLogFileLines, SptLogFileLines
' 20050211 CS: configurable output programs
    PutString IniSectionGeneral, IniOFXExportTo, OFXExportTo
    PutString IniSectionGeneral, IniOFCExportTo, OFCExportTo
    PutString IniSectionGeneral, IniQIFExportTo, QIFExportTo
    PutString IniSectionGeneral, IniQFXExportTo, QFXExportTo
' 20050305 CS: added external resource file
    PutString IniSectionGeneral, IniResFile, ResFile
' 20050531 CS: added NoSortTxns
    PutBool IniSectionGeneral, IniNoSortTxns, NoSortTxns
' 20060825 CS: added SuppressEmptyStatements
    PutBool IniSectionGeneral, IniSuppressEmpty, SuppressEmptyStatements
' 20070208 CS: added ShowQuickStart
    PutBool IniSectionGeneral, IniShowQuickStart, ShowQuickStart
' 20071017 CS: added output code page
    PutNumber IniSectionGeneral, IniOutputCodePage, OutputCodePage
    PutBool IniSectionGeneral, IniOutputBOM, OutputBOM
    PutNumber IniSectionGeneral, IniLineEnding, LineEnding
' 20080511 CS: added UseOldFITID
    PutBool IniSectionGeneral, IniUseOldFITID, UseOldFITID
' 20080610 CS: added QFXForceCurrency
    PutString IniSectionGeneral, IniQFXForceCurrency, QFXForceCurrency
' 20080729 CS: added NoConfirmImport
    PutBool IniSectionGeneral, IniNoConfirmImport, NoConfirmImport

    Save = True
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       SetBookDateMode
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       02/02/2004-22:15:28
'
' Parameters :       iMode (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub SetBookDateMode(iMode As Integer)
    BookDateMode = iMode
    Changed = True
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       SetLastInputDirectory
' Description:       Update setting for Download directory
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       15/02/2004-23:28:04
'
' Parameters :       sDir (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub SetLastInputDirectory(sDir As String)
    LastInputDirectory = sDir
    PutString IniSectionGeneral, IniLastInputDirectory, LastInputDirectory
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       MT2OFX
' Procedure  :       SetLastInputExtension
' Description:
' Created by :       Colin Smale
' Machine    :       55662M8
' Date-Time  :       16/02/2004-00:10:59
'
' Parameters :       sExt (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub SetLastInputExtension(sExt As String)
    LastInputExtension = UCase$(sExt)
    PutString IniSectionGeneral, IniLastInputExtension, LastInputExtension
End Sub

Private Sub Class_Initialize()
    On Error Resume Next
    Dim fso As New Scripting.FileSystemObject
    
    xMyDocumentsPath = GetSpecialFolder(CSIDL_PERSONAL)
    xAppDataPath = GetSpecialFolder(CSIDL_APPDATA) & "\MT2OFX"
    xAppPath = App.Path
    xResourcePath = xAppPath & "\Res"
    If Not fso.FolderExists(xResourcePath) Then
        xResourcePath = xAppPath
    End If
    xScriptPath = xAppPath & "\Scripts"
    If Not fso.FolderExists(xScriptPath) Then
        xScriptPath = xAppPath
    End If
    xHelpPath = xAppPath
    xCustomOutputPath = xAppPath & "\CustomOutput"
End Sub
