VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pActDate As String
Private pKindData As String
Private pAutomType As String
Private pRegNo As String
Private pProducer As String
Private pVIN As String
Private pEngine As String
Private pMadeYear As String

Property Get MadeYear() As String
    MadeYear = pMadeYear
End Property

Property Let MadeYear(s As String)
    pMadeYear = s
End Property

Property Get Engine() As String
    Engine = pEngine
End Property

Property Let Engine(s As String)
    pEngine = s
End Property

Property Get VIN() As String
    VIN = pVIN
End Property

Property Let VIN(s As String)
    pVIN = s
End Property

Property Get Producer() As String
    Producer = pProducer
End Property

Property Let Producer(s As String)
    pProducer = s
End Property

Property Get RegNo() As String
    RegNo = pRegNo
End Property

Property Let RegNo(s As String)
    pRegNo = s
End Property

Property Get AutomType() As String
    AutomType = pAutomType
End Property

Property Let AutomType(s As String)
    pAutomType = s
End Property

Property Get KindData() As String
    KindData = pKindData
End Property

Property Let KindData(s As String)
    pKindData = s
End Property

Property Get ActDate() As String
    ActDate = pActDate
End Property

Property Let ActDate(s As String)
    pActDate = s
End Property

