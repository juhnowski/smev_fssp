VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddressData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pActDate As String
Private pStrAddr As String
Private pcountryCode As String
Private pOkato As String
Private pOktmo As String
Private pzipCode As String
Private parea As String
Private pcity As String
Private pstreet As String
Private phouseNumber As String
Private pflatNumber As String
Private paddressText As String
Private pRegDate As String

Property Get regDate() As String
    regDate = pRegDate
End Property

Property Let regDate(s As String)
    pRegDate = s
End Property

Property Get addressText() As String
    addressText = paddressText
End Property

Property Let addressText(s As String)
    paddressText = s
End Property

Property Get flatNumber() As String
    flatNumber = pflatNumber
End Property

Property Let flatNumber(s As String)
    pflatNumber = s
End Property

Property Get houseNumber() As String
    houseNumber = phouseNumber
End Property

Property Let houseNumber(s As String)
    phouseNumber = s
End Property

Property Get street() As String
    street = pstreet
End Property

Property Let street(s As String)
    pstreet = s
End Property

Property Get city() As String
    city = pcity
End Property

Property Let city(s As String)
    pcity = s
End Property

Property Get area() As String
    area = parea
End Property

Property Let area(s As String)
    parea = s
End Property

Property Get zipCode() As String
    zipCode = pzipCode
End Property

Property Let zipCode(s As String)
    pzipCode = s
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

Property Get countryCode() As String
    countryCode = pcountryCode
End Property

Property Let countryCode(s As String)
    pcountryCode = s
End Property

Property Get StrAddr() As String
    StrAddr = pStrAddr
End Property

Property Let StrAddr(s As String)
    pStrAddr = s
End Property

Property Get ActDate() As String
    ActDate = pActDate
End Property

Property Let ActDate(s As String)
    pActDate = s
End Property

