VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'<CSCC>
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : SecurityInfo
'    Project    : MT2OFX
'
'    Description: Information about a Security (share etc)
'
'    Modified   : $Author: Colin $
'--------------------------------------------------------------------------------
Private Const ModuleHeader As String = "$Header: /MT2OFX/SecurityInfo.cls 3     15/06/09 19:25 Colin $"
' $History: SecurityInfo.cls $
' 
' *****************  Version 3  *****************
' User: Colin        Date: 15/06/09   Time: 19:25
' Updated in $/MT2OFX
' For transfer to new laptop
'
' *****************  Version 1  *****************
' User: Colin        Date: 16/08/06   Time: 1:14
' Created in $/MT2OFX

'</CSCC>

Public Enum SecurityTypeConstants
    sctyOther = 0
    sctyMutualFund = 1
    sctyStock = 2
    sctyOption = 3
    sctyDebt = 4
End Enum
Public SecurityType As SecurityTypeConstants

Public UniqueID As String
Public UniqueIDType As String

Public Name As String
Public Ticker As String
Public FIID As String
Public Rating As String
Public UnitPrice As Double
Public UnitPriceDate As Date
Public Ccy As String
Public ExchangeRate As Double
Public Memo As String
Public FIAssetClass As String

Public Enum StockTypeConstants
    sttyOther = 0
    sttyCommon = 1
    sttyPreferred = 2
    sttyConvertible = 3
End Enum
Public StockType As StockTypeConstants

Public Yield As Double
Public YieldDate As Date

Public Enum AssetClassContstants
    asclOther = 0
    asclDomesticBond = 1
    asclInternationalBond = 2
    asclLargeStock = 3
    asclSmallStock = 4
    asclInternationalStock = 5
    asclMoneyMarket = 6
End Enum
Public AssetClass As AssetClassContstants

Public Function AssetClassToString(ascl As AssetClassContstants) As String
    Dim sTmp As String
    Select Case ascl
    Case asclOther: sTmp = "OTHER"
    Case asclDomesticBond: sTmp = "DOMESTICBOND"
    Case asclInternationalBond: sTmp = "INTLBOND"
    Case asclLargeStock: sTmp = "LARGESTOCK"
    Case asclSmallStock: sTmp = "SMALLSTOCK"
    Case asclInternationalStock: sTmp = "INTLSTOCK"
    Case asclMoneyMarket: sTmp = "MONEYMRKT"
    Case Else
        Debug.Print "AssetClassToString: unknown value " & CStr(ascl)
        Debug.Assert False
        sTmp = "OTHER"
    End Select
    AssetClassToString = sTmp
End Function

Public Property Get AssetClassString() As String
    AssetClassString = AssetClassToString(AssetClass)
End Property

Public Function StockTypeToString(stty As StockTypeConstants)
    Dim sTmp As String
    Select Case stty
    Case sttyOther: sTmp = "OTHER"
    Case sttyCommon: sTmp = "COMMON"
    Case sttyPreferred: sTmp = "PREFERRED"
    Case sttyConvertible: sTmp = "CONVERTIBLE"
    Case Else
        Debug.Print "StockClassToString: unknown value " & CStr(stty)
        Debug.Assert False
        sTmp = "OTHER"
    End Select
End Function

Public Property Get StockTypeString() As String
    StockTypeString = StockTypeToString(StockType)
End Property

Private Sub Class_Initialize()
    ExchangeRate = 1#
    UnitPriceDate = NODATE
End Sub
