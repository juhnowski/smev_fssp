VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SvedNedvData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pActDate As String
Private pKindData As String
Private pNaimNedv As String
Private pKadastrN As String
Private pAdresNedv As String
Private pSNedv As String
Private pRegisterDate As String


Property Get RegisterDate() As String
    RegisterDate = pRegisterDate
End Property

Property Let RegisterDate(s As String)
    pRegisterDate = s
End Property

Property Get SNedv() As String
    SNedv = pSNedv
End Property

Property Let SNedv(s As String)
    pSNedv = s
End Property

Property Get AdresNedv() As String
    AdresNedv = pAdresNedv
End Property

Property Let AdresNedv(s As String)
    pAdresNedv = s
End Property

Property Get KadastrN() As String
    KadastrN = pKadastrN
End Property

Property Let KadastrN(s As String)
    pKadastrN = s
End Property

Property Get NaimNedv() As String
    NaimNedv = pNaimNedv
End Property

Property Let NaimNedv(s As String)
    pNaimNedv = s
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

