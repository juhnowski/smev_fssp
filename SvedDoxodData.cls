VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SvedDoxodData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pActDate As String
Private pKindData As String
Private pDataDox As String
Private pSumDox As String
Private pGround As String

Property Get Ground() As String
    Ground = pGround
End Property

Property Let Ground(s As String)
    pGround = s
End Property

Property Get SumDox() As String
    SumDox = pSumDox
End Property

Property Let SumDox(s As String)
    pSumDox = s
End Property

Property Get DataDox() As String
    DataDox = pDataDox
End Property

Property Let DataDox(s As String)
    pDataDox = s
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

