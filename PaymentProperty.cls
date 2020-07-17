VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PaymentProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pRecpName As String
Private pRecpBank As String
Private pRecpBIK As String
Private pRecpCnt As String
Private pRecpINN As String
Private pRecpKPP As String
Private pOkato As String
Private pOktmo As String
Private pKbk As String
Private pPokPl As String
Private pRekvSum As String

Property Get RekvSum() As String
    RekvSum = pRekvSum
End Property

Property Let RekvSum(s As String)
    pRekvSum = s
End Property

Property Get PokPl() As String
    PokPl = pPokPl
End Property

Property Let PokPl(s As String)
    pPokPl = s
End Property

Property Get Kbk() As String
    Kbk = pKbk
End Property

Property Let Kbk(s As String)
    pKbk = s
End Property

Property Get Oktmo() As String
    Oktmo = pOktmo
End Property

Property Let Oktmo(s As String)
    pOktmo = s
End Property

Property Get Okato() As String
    Okato = pOkato
End Property

Property Let Okato(s As String)
    pOkato = s
End Property

Property Get RecpKPP() As String
    RecpKPP = pRecpKPP
End Property

Property Let RecpKPP(s As String)
    pRecpKPP = s
End Property

Property Get RecpINN() As String
    RecpINN = pRecpINN
End Property

Property Let RecpINN(s As String)
    pRecpINN = s
End Property

Property Get RecpCnt() As String
    RecpCnt = pRecpCnt
End Property

Property Let RecpCnt(s As String)
    pRecpCnt = s
End Property

Property Get RecpBIK() As String
    RecpBIK = pRecpBIK
End Property

Property Let RecpBIK(s As String)
    pRecpBIK = s
End Property

Property Get RecpBank() As String
    RecpBank = pRecpBank
End Property

Property Let RecpBank(s As String)
    pRecpBank = s
End Property


Property Get RecpName() As String
    RecpName = pRecpName
End Property

Property Let RecpName(s As String)
    pRecpName = s
End Property

