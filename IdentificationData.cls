VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IdentificationData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pActDate As String
Private pKindData As String
Private pTypeDoc As String
Private pSerDoc As String
Private pNumDoc As String
Private pIssuedDoc As String
Private pDateDoc As String
Private pSurname As String
Private pFirstName As String
Private pPatronymic As String

Property Get Patronymic() As String
    Patronymic = pPatronymic
End Property

Property Let Patronymic(s As String)
    pPatronymic = s
End Property

Property Get FirstName() As String
    FirstName = pFirstName
End Property

Property Let FirstName(s As String)
    pFirstName = s
End Property

Property Get Surname() As String
    Surname = pSurname
End Property

Property Let Surname(s As String)
    pSurname = s
End Property

Property Get DateDoc() As String
    DateDoc = pDateDoc
End Property

Property Let DateDoc(s As String)
    pDateDoc = s
End Property

Property Get IssuedDoc() As String
    IssuedDoc = pIssuedDoc
End Property

Property Let IssuedDoc(s As String)
    pIssuedDoc = s
End Property

Property Get NumDoc() As String
    NumDoc = pNumDoc
End Property

Property Let NumDoc(s As String)
    pNumDoc = s
End Property

Property Get SerDoc() As String
    SerDoc = pSerDoc
End Property

Property Let SerDoc(s As String)
    pSerDoc = s
End Property

Property Get TypeDoc() As String
    TypeDoc = pTypeDoc
End Property

Property Let TypeDoc(s As String)
    pTypeDoc = s
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

