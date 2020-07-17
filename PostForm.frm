VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PostForm 
   Caption         =   "Постановление"
   ClientHeight    =   7575
   ClientLeft      =   4365
   ClientTop       =   2505
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   12495
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11400
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Открыть"
      Height          =   375
      Left            =   9480
      TabIndex        =   172
      Top             =   120
      Width           =   2895
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   6
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Общие"
      TabPicture(0)   =   "PostForm.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "AktDate"
      Tab(0).Control(1)=   "OrganSignFIO"
      Tab(0).Control(2)=   "OrganSignPost"
      Tab(0).Control(3)=   "OrganAdr"
      Tab(0).Control(4)=   "Organ"
      Tab(0).Control(5)=   "OrganCode"
      Tab(0).Control(6)=   "DeloDate"
      Tab(0).Control(7)=   "DeloNum"
      Tab(0).Control(8)=   "IDDate"
      Tab(0).Control(9)=   "IDNum"
      Tab(0).Control(10)=   "IDType"
      Tab(0).Control(11)=   "DocDate"
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(13)=   "Label12"
      Tab(0).Control(14)=   "Label11"
      Tab(0).Control(15)=   "Label10"
      Tab(0).Control(16)=   "Label9"
      Tab(0).Control(17)=   "Label8"
      Tab(0).Control(18)=   "Label7"
      Tab(0).Control(19)=   "Label6"
      Tab(0).Control(20)=   "Label5"
      Tab(0).Control(21)=   "Label4"
      Tab(0).Control(22)=   "Label3"
      Tab(0).Control(23)=   "Label2"
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Должник"
      TabPicture(1)   =   "PostForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(2)=   "Label16"
      Tab(1).Control(3)=   "Label17"
      Tab(1).Control(4)=   "Label18"
      Tab(1).Control(5)=   "Label19"
      Tab(1).Control(6)=   "Label20"
      Tab(1).Control(7)=   "Label21"
      Tab(1).Control(8)=   "Label22"
      Tab(1).Control(9)=   "Label23"
      Tab(1).Control(10)=   "DebtorType"
      Tab(1).Control(11)=   "DebtorName"
      Tab(1).Control(12)=   "Surname"
      Tab(1).Control(13)=   "FirstName"
      Tab(1).Control(14)=   "Patronymic"
      Tab(1).Control(15)=   "DebtorAdr"
      Tab(1).Control(16)=   "DebtorBirthDate"
      Tab(1).Control(17)=   "DebtorBirthPlace"
      Tab(1).Control(18)=   "DebtorINN"
      Tab(1).Control(19)=   "DebtorRegDate"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "Взыскатель"
      TabPicture(2)   =   "PostForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "IDSum"
      Tab(2).Control(1)=   "IDSubjName"
      Tab(2).Control(2)=   "IDSubj"
      Tab(2).Control(3)=   "ClaimerAdr"
      Tab(2).Control(4)=   "ClaimerName"
      Tab(2).Control(5)=   "ClaimerType"
      Tab(2).Control(6)=   "Label29"
      Tab(2).Control(7)=   "Label28"
      Tab(2).Control(8)=   "Label27"
      Tab(2).Control(9)=   "Label26"
      Tab(2).Control(10)=   "Label25"
      Tab(2).Control(11)=   "Label24"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Данные"
      TabPicture(3)   =   "PostForm.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "SSTab2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Платеж"
      TabPicture(4)   =   "PostForm.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "RekvSum"
      Tab(4).Control(1)=   "PokPl"
      Tab(4).Control(2)=   "Kbk"
      Tab(4).Control(3)=   "Oktmo"
      Tab(4).Control(4)=   "Okato"
      Tab(4).Control(5)=   "RecpKPP"
      Tab(4).Control(6)=   "RecpINN"
      Tab(4).Control(7)=   "RecpCnt"
      Tab(4).Control(8)=   "RecpBIK"
      Tab(4).Control(9)=   "RecpBank"
      Tab(4).Control(10)=   "RecpName"
      Tab(4).Control(11)=   "PaymentProperties"
      Tab(4).Control(12)=   "Label81"
      Tab(4).Control(13)=   "Label80"
      Tab(4).Control(14)=   "Label79"
      Tab(4).Control(15)=   "Label78"
      Tab(4).Control(16)=   "Label77"
      Tab(4).Control(17)=   "Label76"
      Tab(4).Control(18)=   "Label75"
      Tab(4).Control(19)=   "Label74"
      Tab(4).Control(20)=   "Label73"
      Tab(4).Control(21)=   "Label72"
      Tab(4).Control(22)=   "Label71"
      Tab(4).Control(23)=   "Label70"
      Tab(4).ControlCount=   24
      TabCaption(5)   =   "Подпись"
      TabPicture(5)   =   "PostForm.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "SignedInfo"
      Tab(5).Control(1)=   "X509Certificate"
      Tab(5).Control(2)=   "SignatureValue"
      Tab(5).Control(3)=   "Label84"
      Tab(5).Control(4)=   "Label83"
      Tab(5).Control(5)=   "Label82"
      Tab(5).ControlCount=   6
      Begin VB.TextBox SignedInfo 
         Height          =   375
         Left            =   -74880
         TabIndex        =   174
         Text            =   "Text1"
         Top             =   720
         Width           =   12015
      End
      Begin VB.TextBox X509Certificate 
         Height          =   4455
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   171
         Text            =   "PostForm.frx":00A8
         Top             =   2160
         Width           =   11895
      End
      Begin VB.TextBox SignatureValue 
         Height          =   405
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   169
         Text            =   "PostForm.frx":00AE
         Top             =   1440
         Width           =   12015
      End
      Begin VB.TextBox RekvSum 
         Height          =   375
         Left            =   -68520
         TabIndex        =   166
         Text            =   "Text1"
         Top             =   6120
         Width           =   2415
      End
      Begin VB.TextBox PokPl 
         Height          =   375
         Left            =   -68520
         TabIndex        =   164
         Text            =   "Text1"
         Top             =   5640
         Width           =   2415
      End
      Begin VB.TextBox Kbk 
         Height          =   375
         Left            =   -68520
         TabIndex        =   162
         Text            =   "Text1"
         Top             =   5160
         Width           =   4695
      End
      Begin VB.TextBox Oktmo 
         Height          =   375
         Left            =   -67800
         TabIndex        =   160
         Text            =   "Text1"
         Top             =   4680
         Width           =   2055
      End
      Begin VB.TextBox Okato 
         Height          =   405
         Left            =   -70560
         TabIndex        =   158
         Text            =   "Text1"
         Top             =   4680
         Width           =   1695
      End
      Begin VB.TextBox RecpKPP 
         Height          =   375
         Left            =   -66720
         TabIndex        =   156
         Text            =   "Text1"
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox RecpINN 
         Height          =   375
         Left            =   -66720
         TabIndex        =   154
         Text            =   "Text1"
         Top             =   3720
         Width           =   2895
      End
      Begin VB.TextBox RecpCnt 
         Height          =   375
         Left            =   -66720
         TabIndex        =   152
         Text            =   "Text1"
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox RecpBIK 
         Height          =   375
         Left            =   -66720
         TabIndex        =   150
         Text            =   "Text1"
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox RecpBank 
         Height          =   375
         Left            =   -71400
         TabIndex        =   148
         Text            =   "Text1"
         Top             =   2280
         Width           =   7575
      End
      Begin VB.TextBox RecpName 
         Height          =   1095
         Left            =   -71400
         MultiLine       =   -1  'True
         TabIndex        =   146
         Text            =   "PostForm.frx":00B4
         Top             =   720
         Width           =   7575
      End
      Begin VB.ListBox PaymentProperties 
         Height          =   6105
         Left            =   -74880
         TabIndex        =   144
         Top             =   480
         Width           =   3375
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   6135
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   10821
         _Version        =   393216
         Tabs            =   5
         Tab             =   2
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Документ"
         TabPicture(0)   =   "PostForm.frx":00BA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Doc_Patronymic"
         Tab(0).Control(1)=   "Doc_FirstName"
         Tab(0).Control(2)=   "Doc_Surname"
         Tab(0).Control(3)=   "Doc_DateDoc"
         Tab(0).Control(4)=   "Doc_IssuedDoc"
         Tab(0).Control(5)=   "Doc_NumDoc"
         Tab(0).Control(6)=   "Doc_SerDoc"
         Tab(0).Control(7)=   "TypeDoc"
         Tab(0).Control(8)=   "Doc_ActDate"
         Tab(0).Control(9)=   "IdentificationData"
         Tab(0).Control(10)=   "Label38"
         Tab(0).Control(11)=   "Label37"
         Tab(0).Control(12)=   "Label36"
         Tab(0).Control(13)=   "Label35"
         Tab(0).Control(14)=   "Label34"
         Tab(0).Control(15)=   "Label33"
         Tab(0).Control(16)=   "Label32"
         Tab(0).Control(17)=   "Label31"
         Tab(0).Control(18)=   "Label30"
         Tab(0).ControlCount=   19
         TabCaption(1)   =   "Адрес"
         TabPicture(1)   =   "PostForm.frx":00D6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "AD_addressText"
         Tab(1).Control(1)=   "AD_flatNumber"
         Tab(1).Control(2)=   "AD_houseNumber"
         Tab(1).Control(3)=   "AD_street"
         Tab(1).Control(4)=   "AD_city"
         Tab(1).Control(5)=   "AD_area"
         Tab(1).Control(6)=   "AD_zipCode"
         Tab(1).Control(7)=   "AD_OKTMO"
         Tab(1).Control(8)=   "AD_OKATO"
         Tab(1).Control(9)=   "AD_countryCode"
         Tab(1).Control(10)=   "AD_StrAddr"
         Tab(1).Control(11)=   "AD_RegDate"
         Tab(1).Control(12)=   "AD_ActDate"
         Tab(1).Control(13)=   "AddressData"
         Tab(1).Control(14)=   "Label51"
         Tab(1).Control(15)=   "Label50"
         Tab(1).Control(16)=   "Label49"
         Tab(1).Control(17)=   "Label48"
         Tab(1).Control(18)=   "Label47"
         Tab(1).Control(19)=   "Label46"
         Tab(1).Control(20)=   "Label45"
         Tab(1).Control(21)=   "Label44"
         Tab(1).Control(22)=   "Label43"
         Tab(1).Control(23)=   "Label42"
         Tab(1).Control(24)=   "Label41"
         Tab(1).Control(25)=   "Label40"
         Tab(1).Control(26)=   "Label39"
         Tab(1).ControlCount=   27
         TabCaption(2)   =   "Транспорт"
         TabPicture(2)   =   "PostForm.frx":00F2
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Label52"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Label53"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label54"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Label55"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Label56"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "Label57"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "Label58"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "TransportData"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "Tr_ActDate"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "AutomType"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "RegNo"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "Producer"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).Control(12)=   "vin"
         Tab(2).Control(12).Enabled=   0   'False
         Tab(2).Control(13)=   "Engine"
         Tab(2).Control(13).Enabled=   0   'False
         Tab(2).Control(14)=   "MadeYear"
         Tab(2).Control(14).Enabled=   0   'False
         Tab(2).ControlCount=   15
         TabCaption(3)   =   "Доход"
         TabPicture(3)   =   "PostForm.frx":010E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Ground"
         Tab(3).Control(1)=   "SumDox"
         Tab(3).Control(2)=   "DataDox"
         Tab(3).Control(3)=   "D_ActDate"
         Tab(3).Control(4)=   "SvedDoxodData"
         Tab(3).Control(5)=   "Label62"
         Tab(3).Control(6)=   "Label61"
         Tab(3).Control(7)=   "Label60"
         Tab(3).Control(8)=   "Label59"
         Tab(3).ControlCount=   9
         TabCaption(4)   =   "Недвижимость"
         TabPicture(4)   =   "PostForm.frx":012A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "RegisterDate"
         Tab(4).Control(1)=   "SNedv"
         Tab(4).Control(2)=   "AdresNedv"
         Tab(4).Control(3)=   "KadastrN"
         Tab(4).Control(4)=   "NaimNedv"
         Tab(4).Control(5)=   "SND_ActDate"
         Tab(4).Control(6)=   "SvedNedvData"
         Tab(4).Control(7)=   "Label69"
         Tab(4).Control(8)=   "Label68"
         Tab(4).Control(9)=   "Label67"
         Tab(4).Control(10)=   "Label66"
         Tab(4).Control(11)=   "Label65"
         Tab(4).Control(12)=   "Label64"
         Tab(4).Control(13)=   "Label63"
         Tab(4).ControlCount=   14
         Begin VB.TextBox RegisterDate 
            Height          =   375
            Left            =   -68160
            TabIndex        =   143
            Text            =   "Text1"
            Top             =   3120
            Width           =   1815
         End
         Begin VB.TextBox SNedv 
            Height          =   375
            Left            =   -68160
            TabIndex        =   140
            Text            =   "Text1"
            Top             =   2640
            Width           =   1815
         End
         Begin VB.TextBox AdresNedv 
            Height          =   615
            Left            =   -68760
            TabIndex        =   138
            Text            =   "Text1"
            Top             =   1920
            Width           =   4575
         End
         Begin VB.TextBox KadastrN 
            Height          =   375
            Left            =   -68160
            TabIndex        =   136
            Text            =   "Text1"
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox NaimNedv 
            Height          =   375
            Left            =   -68160
            TabIndex        =   134
            Text            =   "Text1"
            Top             =   960
            Width           =   3975
         End
         Begin VB.TextBox SND_ActDate 
            Height          =   375
            Left            =   -69240
            TabIndex        =   132
            Text            =   "Text1"
            Top             =   480
            Width           =   1695
         End
         Begin VB.ListBox SvedNedvData 
            Height          =   5325
            Left            =   -74880
            TabIndex        =   130
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox Ground 
            Height          =   3975
            Left            =   -69960
            MultiLine       =   -1  'True
            TabIndex        =   129
            Text            =   "PostForm.frx":0146
            Top             =   1920
            Width           =   5895
         End
         Begin VB.TextBox SumDox 
            Height          =   375
            Left            =   -69960
            TabIndex        =   127
            Text            =   "Text1"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox DataDox 
            Height          =   375
            Left            =   -69960
            TabIndex        =   125
            Text            =   "Text1"
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox D_ActDate 
            Height          =   375
            Left            =   -69120
            TabIndex        =   123
            Text            =   "Text1"
            Top             =   480
            Width           =   1935
         End
         Begin VB.ListBox SvedDoxodData 
            Height          =   5520
            Left            =   -74880
            TabIndex        =   121
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox MadeYear 
            Height          =   375
            Left            =   5640
            TabIndex        =   120
            Text            =   "Text1"
            Top             =   3360
            Width           =   1815
         End
         Begin VB.TextBox Engine 
            Height          =   375
            Left            =   5640
            TabIndex        =   118
            Text            =   "Text1"
            Top             =   2880
            Width           =   4455
         End
         Begin VB.TextBox vin 
            Height          =   375
            Left            =   5640
            TabIndex        =   116
            Text            =   "Text1"
            Top             =   2400
            Width           =   4455
         End
         Begin VB.TextBox Producer 
            Height          =   375
            Left            =   5640
            TabIndex        =   114
            Text            =   "Text1"
            Top             =   1920
            Width           =   5295
         End
         Begin VB.TextBox RegNo 
            Height          =   405
            Left            =   5640
            TabIndex        =   112
            Text            =   "Text1"
            Top             =   1440
            Width           =   2415
         End
         Begin VB.TextBox AutomType 
            Height          =   375
            Left            =   5640
            TabIndex        =   110
            Text            =   "Text1"
            Top             =   960
            Width           =   5295
         End
         Begin VB.TextBox Tr_ActDate 
            Height          =   375
            Left            =   5640
            TabIndex        =   108
            Text            =   "Text1"
            Top             =   480
            Width           =   3015
         End
         Begin VB.ListBox TransportData 
            Height          =   5325
            Left            =   120
            TabIndex        =   106
            Top             =   480
            Width           =   3255
         End
         Begin VB.TextBox AD_addressText 
            Height          =   735
            Left            =   -70080
            TabIndex        =   105
            Text            =   "Text1"
            Top             =   5160
            Width           =   6015
         End
         Begin VB.TextBox AD_flatNumber 
            Height          =   375
            Left            =   -67680
            TabIndex        =   103
            Text            =   "Text1"
            Top             =   4680
            Width           =   1335
         End
         Begin VB.TextBox AD_houseNumber 
            Height          =   375
            Left            =   -70080
            TabIndex        =   101
            Text            =   "Text1"
            Top             =   4680
            Width           =   1455
         End
         Begin VB.TextBox AD_street 
            Height          =   375
            Left            =   -70080
            TabIndex        =   99
            Text            =   "Text1"
            Top             =   4200
            Width           =   6015
         End
         Begin VB.TextBox AD_city 
            Height          =   375
            Left            =   -70080
            TabIndex        =   97
            Text            =   "Text1"
            Top             =   3720
            Width           =   6015
         End
         Begin VB.TextBox AD_area 
            Height          =   375
            Left            =   -70080
            TabIndex        =   95
            Text            =   "Text1"
            Top             =   3240
            Width           =   6015
         End
         Begin VB.TextBox AD_zipCode 
            Height          =   375
            Left            =   -70080
            TabIndex        =   93
            Text            =   "Text1"
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox AD_OKTMO 
            Height          =   405
            Left            =   -67200
            TabIndex        =   91
            Text            =   "Text1"
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox AD_OKATO 
            Height          =   375
            Left            =   -70080
            TabIndex        =   89
            Text            =   "Text1"
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox AD_countryCode 
            Height          =   375
            Left            =   -65760
            TabIndex        =   87
            Text            =   "Text1"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox AD_StrAddr 
            Height          =   615
            Left            =   -69120
            MultiLine       =   -1  'True
            TabIndex        =   85
            Text            =   "PostForm.frx":014C
            Top             =   960
            Width           =   5055
         End
         Begin VB.TextBox AD_RegDate 
            Height          =   375
            Left            =   -65520
            TabIndex        =   83
            Text            =   "Text1"
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox AD_ActDate 
            Height          =   375
            Left            =   -69120
            TabIndex        =   81
            Text            =   "Text1"
            Top             =   480
            Width           =   1815
         End
         Begin VB.ListBox AddressData 
            Height          =   5130
            Left            =   -74880
            TabIndex        =   79
            Top             =   480
            Width           =   3735
         End
         Begin VB.TextBox Doc_Patronymic 
            Height          =   375
            Left            =   -69360
            TabIndex        =   78
            Text            =   "Text1"
            Top             =   4320
            Width           =   5295
         End
         Begin VB.TextBox Doc_FirstName 
            Height          =   375
            Left            =   -69360
            TabIndex        =   77
            Text            =   "Text1"
            Top             =   3840
            Width           =   5295
         End
         Begin VB.TextBox Doc_Surname 
            Height          =   375
            Left            =   -69360
            TabIndex        =   76
            Text            =   "Text1"
            Top             =   3360
            Width           =   5295
         End
         Begin VB.TextBox Doc_DateDoc 
            Height          =   375
            Left            =   -69360
            TabIndex        =   72
            Text            =   "Text1"
            Top             =   2880
            Width           =   2415
         End
         Begin VB.TextBox Doc_IssuedDoc 
            Height          =   615
            Left            =   -69360
            MultiLine       =   -1  'True
            TabIndex        =   70
            Text            =   "PostForm.frx":0152
            Top             =   2160
            Width           =   5295
         End
         Begin VB.TextBox Doc_NumDoc 
            Height          =   285
            Left            =   -69360
            TabIndex        =   68
            Text            =   "Text1"
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox Doc_SerDoc 
            Height          =   285
            Left            =   -69360
            TabIndex        =   66
            Text            =   "Text1"
            Top             =   1440
            Width           =   1815
         End
         Begin VB.ComboBox TypeDoc 
            Height          =   315
            Left            =   -69360
            TabIndex        =   64
            Text            =   "Combo1"
            Top             =   960
            Width           =   5295
         End
         Begin VB.TextBox Doc_ActDate 
            Height          =   375
            Left            =   -69360
            TabIndex        =   63
            Text            =   "Text1"
            Top             =   480
            Width           =   1815
         End
         Begin VB.ListBox IdentificationData 
            Height          =   5325
            Left            =   -74880
            TabIndex        =   60
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label Label69 
            Caption         =   "Дата регистрации"
            Height          =   375
            Left            =   -69960
            TabIndex        =   142
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label68 
            Caption         =   "м2"
            Height          =   255
            Left            =   -66240
            TabIndex        =   141
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label67 
            Caption         =   "площадь"
            Height          =   255
            Left            =   -69240
            TabIndex        =   139
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label66 
            Caption         =   "точный адрес (местоположение)"
            Height          =   255
            Left            =   -71400
            TabIndex        =   137
            Top             =   1920
            Width           =   3255
         End
         Begin VB.Label Label65 
            Caption         =   "кадастровый (условный) номер"
            Height          =   255
            Left            =   -70320
            TabIndex        =   135
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label Label64 
            Caption         =   "наименование объекта недвижимости"
            Height          =   255
            Left            =   -71280
            TabIndex        =   133
            Top             =   960
            Width           =   3495
         End
         Begin VB.Label Label63 
            Caption         =   "Актуальность сведений"
            Height          =   375
            Left            =   -71280
            TabIndex        =   131
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label62 
            Caption         =   "Основание"
            Height          =   255
            Left            =   -71160
            TabIndex        =   128
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label61 
            Caption         =   "Сумма дохода"
            Height          =   375
            Left            =   -71280
            TabIndex        =   126
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label60 
            Caption         =   "Дата дохода"
            Height          =   375
            Left            =   -71280
            TabIndex        =   124
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label59 
            Caption         =   "Актуальность сведений"
            Height          =   375
            Left            =   -71280
            TabIndex        =   122
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label58 
            Caption         =   "Год выпуска"
            Height          =   255
            Left            =   3480
            TabIndex        =   119
            Top             =   3360
            Width           =   1935
         End
         Begin VB.Label Label57 
            Caption         =   "Номер двигателя"
            Height          =   375
            Left            =   3480
            TabIndex        =   117
            Top             =   2880
            Width           =   2055
         End
         Begin VB.Label Label56 
            Caption         =   "VIN"
            Height          =   375
            Left            =   3480
            TabIndex        =   115
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label55 
            Caption         =   "марка транспортного средства"
            Height          =   375
            Left            =   3480
            TabIndex        =   113
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label Label54 
            Caption         =   "государственный регистрационный знак"
            Height          =   495
            Left            =   3480
            TabIndex        =   111
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label53 
            Caption         =   "категория транспортного средства"
            Height          =   495
            Left            =   3480
            TabIndex        =   109
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label52 
            Caption         =   "Актуальность сведений"
            Height          =   375
            Left            =   3600
            TabIndex        =   107
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label51 
            Caption         =   "Адрес текстом"
            Height          =   495
            Left            =   -70920
            TabIndex        =   104
            Top             =   5160
            Width           =   855
         End
         Begin VB.Label Label50 
            Caption         =   "Квартира"
            Height          =   375
            Left            =   -68520
            TabIndex        =   102
            Top             =   4680
            Width           =   855
         End
         Begin VB.Label Label49 
            Caption         =   "Дом"
            Height          =   255
            Left            =   -70800
            TabIndex        =   100
            Top             =   4680
            Width           =   735
         End
         Begin VB.Label Label48 
            Caption         =   "Улица"
            Height          =   255
            Left            =   -70920
            TabIndex        =   98
            Top             =   4200
            Width           =   735
         End
         Begin VB.Label Label47 
            Caption         =   "Город"
            Height          =   255
            Left            =   -70920
            TabIndex        =   96
            Top             =   3720
            Width           =   855
         End
         Begin VB.Label Label46 
            Caption         =   "Область"
            Height          =   255
            Left            =   -70920
            TabIndex        =   94
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label45 
            Caption         =   "Индекс"
            Height          =   255
            Left            =   -70920
            TabIndex        =   92
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label44 
            Caption         =   "OKTMO"
            Height          =   255
            Left            =   -67920
            TabIndex        =   90
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label43 
            Caption         =   "OKATO"
            Height          =   255
            Left            =   -70920
            TabIndex        =   88
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label42 
            Caption         =   "код страны принадлежности должника по Общероссийскому классификатору стран мира"
            Height          =   495
            Left            =   -70920
            TabIndex        =   86
            Top             =   1680
            Width           =   5055
         End
         Begin VB.Label Label41 
            Caption         =   "Адрес должника"
            Height          =   255
            Left            =   -71040
            TabIndex        =   84
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label40 
            Caption         =   "Дата регистрации"
            Height          =   255
            Left            =   -67080
            TabIndex        =   82
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label39 
            Caption         =   "Актуальность сведений"
            Height          =   375
            Left            =   -71040
            TabIndex        =   80
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label38 
            Caption         =   "Отчество"
            Height          =   375
            Left            =   -70320
            TabIndex        =   75
            Top             =   4320
            Width           =   735
         End
         Begin VB.Label Label37 
            Caption         =   "Имя"
            Height          =   375
            Left            =   -70200
            TabIndex        =   74
            Top             =   3840
            Width           =   495
         End
         Begin VB.Label Label36 
            Caption         =   "Фамилия"
            Height          =   255
            Left            =   -70440
            TabIndex        =   73
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label35 
            Caption         =   "дата выдачи документа"
            Height          =   375
            Left            =   -71520
            TabIndex        =   71
            Top             =   3000
            Width           =   2055
         End
         Begin VB.Label Label34 
            Caption         =   "кем выдан"
            Height          =   255
            Left            =   -70440
            TabIndex        =   69
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label33 
            Caption         =   "номер документа"
            Height          =   375
            Left            =   -70920
            TabIndex        =   67
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label32 
            Caption         =   "серия документа"
            Height          =   375
            Left            =   -70920
            TabIndex        =   65
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label31 
            Caption         =   "код вида документа"
            Height          =   375
            Left            =   -71520
            TabIndex        =   62
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label30 
            Caption         =   "Актуальность сведений"
            Height          =   375
            Left            =   -71520
            TabIndex        =   61
            Top             =   480
            Width           =   2535
         End
      End
      Begin VB.TextBox IDSum 
         Height          =   375
         Left            =   -68280
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox IDSubjName 
         Height          =   375
         Left            =   -72600
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   2640
         Width           =   8415
      End
      Begin VB.ComboBox IDSubj 
         Height          =   315
         Left            =   -72600
         TabIndex        =   54
         Top             =   2280
         Width           =   8415
      End
      Begin VB.TextBox ClaimerAdr 
         Height          =   615
         Left            =   -69960
         MultiLine       =   -1  'True
         TabIndex        =   52
         Text            =   "PostForm.frx":0158
         Top             =   1560
         Width           =   5775
      End
      Begin VB.TextBox ClaimerName 
         Height          =   495
         Left            =   -72480
         MultiLine       =   -1  'True
         TabIndex        =   50
         Text            =   "PostForm.frx":015E
         Top             =   960
         Width           =   8295
      End
      Begin VB.ComboBox ClaimerType 
         Height          =   315
         Left            =   -73320
         TabIndex        =   48
         Top             =   480
         Width           =   5895
      End
      Begin VB.TextBox DebtorRegDate 
         Height          =   375
         Left            =   -72240
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   4440
         Width           =   2295
      End
      Begin VB.TextBox DebtorINN 
         Height          =   375
         Left            =   -72240
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   3960
         Width           =   8175
      End
      Begin VB.TextBox DebtorBirthPlace 
         Height          =   375
         Left            =   -72240
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   3480
         Width           =   8175
      End
      Begin VB.TextBox DebtorBirthDate 
         Height          =   375
         Left            =   -72240
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox DebtorAdr 
         Height          =   495
         Left            =   -69720
         MultiLine       =   -1  'True
         TabIndex        =   38
         Text            =   "PostForm.frx":0164
         Top             =   2400
         Width           =   5895
      End
      Begin VB.TextBox Patronymic 
         Height          =   285
         Left            =   -73920
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   2040
         Width           =   10095
      End
      Begin VB.TextBox FirstName 
         Height          =   285
         Left            =   -73920
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   1680
         Width           =   10095
      End
      Begin VB.TextBox Surname 
         Height          =   285
         Left            =   -73920
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   1320
         Width           =   10095
      End
      Begin VB.TextBox DebtorName 
         Height          =   375
         Left            =   -70440
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   840
         Width           =   6735
      End
      Begin VB.ComboBox DebtorType 
         Height          =   315
         ItemData        =   "PostForm.frx":016A
         Left            =   -73560
         List            =   "PostForm.frx":016C
         TabIndex        =   28
         Top             =   480
         Width           =   4695
      End
      Begin VB.TextBox AktDate 
         Height          =   285
         Left            =   -70440
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   5280
         Width           =   6615
      End
      Begin VB.TextBox OrganSignFIO 
         Height          =   285
         Left            =   -70440
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   4800
         Width           =   6615
      End
      Begin VB.TextBox OrganSignPost 
         Height          =   285
         Left            =   -70440
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   4440
         Width           =   6615
      End
      Begin VB.TextBox OrganAdr 
         Height          =   615
         Left            =   -71280
         MultiLine       =   -1  'True
         TabIndex        =   20
         Text            =   "PostForm.frx":016E
         Top             =   3720
         Width           =   7455
      End
      Begin VB.TextBox Organ 
         Height          =   525
         Left            =   -71280
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "PostForm.frx":0174
         Top             =   3120
         Width           =   7455
      End
      Begin VB.TextBox OrganCode 
         Height          =   285
         Left            =   -69600
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox DeloDate 
         Height          =   285
         Left            =   -71520
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox DeloNum 
         Height          =   285
         Left            =   -71520
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox IDDate 
         Height          =   285
         Left            =   -71520
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox IDNum 
         Height          =   285
         Left            =   -71520
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox IDType 
         Height          =   315
         ItemData        =   "PostForm.frx":017A
         Left            =   -71520
         List            =   "PostForm.frx":017C
         TabIndex        =   6
         Top             =   960
         Width           =   7695
      End
      Begin VB.TextBox DocDate 
         Height          =   375
         Left            =   -71520
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label84 
         Caption         =   "SignedInfo"
         Height          =   255
         Left            =   -74880
         TabIndex        =   173
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label83 
         Caption         =   "KeyInfo"
         Height          =   255
         Left            =   -74880
         TabIndex        =   170
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label82 
         Caption         =   "SignatureValue"
         Height          =   255
         Left            =   -74880
         TabIndex        =   168
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label81 
         Caption         =   "руб."
         Height          =   255
         Left            =   -65880
         TabIndex        =   167
         Top             =   6120
         Width           =   735
      End
      Begin VB.Label Label80 
         Caption         =   "перечисляемая сумма"
         Height          =   255
         Left            =   -70440
         TabIndex        =   165
         Top             =   6120
         Width           =   1815
      End
      Begin VB.Label Label79 
         Caption         =   "показатель типа платежа"
         Height          =   255
         Left            =   -70680
         TabIndex        =   163
         Top             =   5640
         Width           =   2055
      End
      Begin VB.Label Label78 
         Caption         =   "код бюджетной классификации"
         Height          =   375
         Left            =   -71160
         TabIndex        =   161
         Top             =   5160
         Width           =   2535
      End
      Begin VB.Label Label77 
         Caption         =   "ОКТМО"
         Height          =   375
         Left            =   -68640
         TabIndex        =   159
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Label Label76 
         Caption         =   "ОКАТО"
         Height          =   255
         Left            =   -71280
         TabIndex        =   157
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label75 
         Caption         =   "код причины постановки на учет получателя"
         Height          =   375
         Left            =   -71160
         TabIndex        =   155
         Top             =   4200
         Width           =   4335
      End
      Begin VB.Label Label74 
         Caption         =   "индивидуальный номер налогоплательщика получателя"
         Height          =   375
         Left            =   -71280
         TabIndex        =   153
         Top             =   3720
         Width           =   4455
      End
      Begin VB.Label Label73 
         Caption         =   "счет получателя"
         Height          =   255
         Left            =   -68160
         TabIndex        =   151
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label72 
         Caption         =   "банковский идентификационный код банка получателя"
         Height          =   375
         Left            =   -71160
         TabIndex        =   149
         Top             =   2760
         Width           =   4335
      End
      Begin VB.Label Label71 
         Caption         =   "наименование банка получателя"
         Height          =   255
         Left            =   -71400
         TabIndex        =   147
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label70 
         Caption         =   "наименование получателя"
         Height          =   495
         Left            =   -71400
         TabIndex        =   145
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label29 
         Caption         =   "общая сумма требований, подлежащих взысканию по исполнительному документу"
         Height          =   375
         Left            =   -74880
         TabIndex        =   57
         Top             =   3120
         Width           =   7095
      End
      Begin VB.Label Label28 
         Caption         =   "предмет исполнения"
         Height          =   255
         Left            =   -74880
         TabIndex        =   55
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label27 
         Caption         =   "код предмета исполнения"
         Height          =   255
         Left            =   -74880
         TabIndex        =   53
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label26 
         Caption         =   "место жительства или место пребывания - для физического лица; место нахождения - для юридического лица"
         Height          =   495
         Left            =   -74880
         TabIndex        =   51
         Top             =   1560
         Width           =   5055
      End
      Begin VB.Label Label25 
         Caption         =   "наименование или фамилия, имя, отчество взыскателя"
         Height          =   495
         Left            =   -74880
         TabIndex        =   49
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label24 
         Caption         =   "тип взыскателя"
         Height          =   375
         Left            =   -74880
         TabIndex        =   47
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "дата регистрации должника"
         Height          =   375
         Left            =   -74520
         TabIndex        =   45
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label22 
         Caption         =   "ИНН должника"
         Height          =   375
         Left            =   -73680
         TabIndex        =   43
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label21 
         Caption         =   "место рождения должника"
         Height          =   255
         Left            =   -74520
         TabIndex        =   41
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label20 
         Caption         =   "дата рождения должника"
         Height          =   255
         Left            =   -74760
         TabIndex        =   39
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label19 
         Caption         =   "место жительства или место пребывания - для физического лица, место нахождения - для юридического лица"
         Height          =   495
         Left            =   -74880
         TabIndex        =   37
         Top             =   2400
         Width           =   5415
      End
      Begin VB.Label Label18 
         Caption         =   "Отчество"
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Имя"
         Height          =   255
         Left            =   -74640
         TabIndex        =   33
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Фамилия"
         Height          =   255
         Left            =   -74760
         TabIndex        =   31
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "наименование или фамилия, имя, отчество должника"
         Height          =   255
         Left            =   -74760
         TabIndex        =   29
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label14 
         Caption         =   "тип должника"
         Height          =   255
         Left            =   -74760
         TabIndex        =   27
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "дата вступления решения в законную силу"
         Height          =   375
         Left            =   -73800
         TabIndex        =   25
         Top             =   5280
         Width           =   3375
      End
      Begin VB.Label Label12 
         Caption         =   "фамилия, имя, отчество должностного лица, выдавшего исполнительный документ"
         Height          =   495
         Left            =   -74880
         TabIndex        =   23
         Top             =   4800
         Width           =   4335
      End
      Begin VB.Label Label11 
         Caption         =   "должность лица, вынесшего исполнительный документ"
         Height          =   375
         Left            =   -74760
         TabIndex        =   21
         Top             =   4440
         Width           =   4335
      End
      Begin VB.Label Label10 
         Caption         =   "адрес органа, выдавшего исполнительный документ"
         Height          =   615
         Left            =   -74880
         TabIndex        =   19
         Top             =   3600
         Width           =   3495
      End
      Begin VB.Label Label9 
         Caption         =   "орган, выдавший исполнительный документ"
         Height          =   375
         Left            =   -74880
         TabIndex        =   17
         Top             =   3120
         Width           =   3495
      End
      Begin VB.Label Label8 
         Caption         =   "код подразделения органа, выдавшего исполнительный документ"
         Height          =   375
         Left            =   -74880
         TabIndex        =   15
         Top             =   2760
         Width           =   5535
      End
      Begin VB.Label Label7 
         Caption         =   "Дата дела"
         Height          =   255
         Left            =   -72720
         TabIndex        =   13
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Номер дела"
         Height          =   255
         Left            =   -72840
         TabIndex        =   11
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Дата исполнительного документа"
         Height          =   375
         Left            =   -74640
         TabIndex        =   9
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Номер исполнительного документа"
         Height          =   255
         Left            =   -74640
         TabIndex        =   7
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Код вида исполнительного документа"
         Height          =   375
         Left            =   -74760
         TabIndex        =   5
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Дата выдачи исполнительного документа"
         Height          =   255
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   3375
      End
   End
   Begin VB.TextBox ExternalKey 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Идентификатор документа"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "PostForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim dictPP As New Collection
Dim flgPP As Boolean
Dim pp As PaymentProperty
Dim idxPP As Integer

Dim dictAD As New Collection
Dim flgAD As Boolean
Dim ad As AddressData
Dim idxAD As Integer

Dim dictID As New Collection
Dim flgID As Boolean
Dim id As IdentificationData
Dim idxID As Integer

Dim dictSDD As New Collection
Dim flgSDD As Boolean
Dim sdd As SvedDoxodData
Dim idxSDD As Integer

Dim dictSND As New Collection
Dim flgSND As Boolean
Dim snd As SvedNedvData
Dim idxSND As Integer

Dim dictTD As New Collection
Dim flgTD As Boolean
Dim td As TransportData
Dim idxTD As Integer

Private Sub AddressData_Click()
    AddressData_Change
End Sub

Private Sub AddressData_KeyPress(KeyAscii As Integer)
    AddressData_Change
End Sub

Private Sub Command1_Click()

CommonDialog1.Filter = "Уведомления (*.xml)|*.xml|All files (*.*)|*.*"
CommonDialog1.DefaultExt = "txt"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.InitDir = "C:\XML"
CommonDialog1.ShowOpen

'The FileName property gives you the variable you need to use
PostForm.openDoc (CommonDialog1.filename)

End Sub


Private Sub SvedNedvData_change()
      Dim i As Integer
      Dim selectedSND As SvedNedvData
      
      For i = 0 To SvedNedvData.ListCount - 1
          If SvedNedvData.Selected(i) Then
              Set selectedSND = dictSND(i + 1)
              Exit For
          End If
      Next i
      
      SND_ActDate.Text = selectedSND.ActDate
      AdresNedv.Text = selectedSND.AdresNedv
      KadastrN.Text = selectedSND.KadastrN
      NaimNedv.Text = selectedSND.NaimNedv
      RegisterDate.Text = selectedSND.RegisterDate
      SNedv.Text = selectedSND.SNedv
      
End Sub

Private Sub IdentificationData_change()
      Dim i As Integer
      Dim selectedID As IdentificationData
      
      For i = 0 To IdentificationData.ListCount - 1
          If IdentificationData.Selected(i) Then
              Set selectedID = dictID(i + 1)
              Exit For
          End If
      Next i
      
      Doc_ActDate.Text = selectedID.ActDate
      Doc_DateDoc.Text = selectedID.DateDoc
      Doc_FirstName.Text = selectedID.FirstName
      Doc_IssuedDoc.Text = selectedID.IssuedDoc
      Doc_NumDoc.Text = selectedID.NumDoc
      Doc_Patronymic.Text = selectedID.Patronymic
      Doc_SerDoc.Text = selectedID.SerDoc
      Doc_Surname.Text = selectedID.Surname
      setTypeDoc (selectedID.TypeDoc)
      
End Sub

Private Sub SvedDoxodData_Change()
      Dim i As Integer
      Dim selectedSDD As SvedDoxodData
      
      For i = 0 To SvedDoxodData.ListCount - 1
          If SvedDoxodData.Selected(i) Then
              Set selectedSDD = dictSDD(i + 1)
              Exit For
          End If
      Next i
      
    D_ActDate.Text = selectedSDD.ActDate
    DataDox.Text = selectedSDD.DataDox
    Ground.Text = selectedSDD.Ground
    SumDox.Text = selectedSDD.SumDox
End Sub


Private Sub AddressData_Change()
      Dim i As Integer
      Dim selectedAD As AddressData
      
      For i = 0 To AddressData.ListCount - 1
          If AddressData.Selected(i) Then
              Set selectedAD = dictAD(i + 1)
              Exit For
          End If
      Next i
      
    AD_ActDate.Text = selectedAD.ActDate
    AD_StrAddr.Text = selectedAD.StrAddr
    AD_RegDate.Text = selectedAD.regDate
    AD_countryCode.Text = selectedAD.countryCode
    AD_OKATO.Text = selectedAD.Okato
    AD_OKTMO.Text = selectedAD.Oktmo
    AD_zipCode.Text = selectedAD.zipCode
    AD_area.Text = selectedAD.area
    AD_city.Text = selectedAD.city
    AD_street.Text = selectedAD.street
    AD_houseNumber.Text = selectedAD.houseNumber
    AD_flatNumber.Text = selectedAD.flatNumber
    AD_addressText.Text = selectedAD.addressText
End Sub

Private Sub PaymentProperties_Change()

      Dim i As Integer
      Dim selectedPP As PaymentProperty
      

      For i = 0 To PaymentProperties.ListCount - 1
          If PaymentProperties.Selected(i) Then
              Set selectedPP = dictPP(i + 1)
              Exit For
          End If
      Next i
     
       
        RecpName.Text = selectedPP.RecpName
        RecpBank.Text = selectedPP.RecpBank
        RecpBIK.Text = selectedPP.RecpBIK
        RecpCnt.Text = selectedPP.RecpCnt
        RecpINN.Text = selectedPP.RecpINN
        RecpKPP.Text = selectedPP.RecpKPP
        Okato.Text = selectedPP.Okato
        Oktmo.Text = selectedPP.Oktmo
        Kbk.Text = selectedPP.Kbk
        PokPl.Text = selectedPP.PokPl
        RekvSum.Text = selectedPP.RekvSum
        
End Sub

Private Sub Form_Load()


flgPP = False
idxPP = 0

flgAD = False
idxAD = 0

flgID = False
idxID = 0

flgTD = False
idxTD = 0

flgSDD = False
idxSDD = 0

flgSND = False
idxSND = 0

IDType.AddItem (" ")
IDType.AddItem ("1-исполнительный лист")
IDType.AddItem ("2-нотариально удостоверенное соглашение об уплате алиментов")
IDType.AddItem ("3-акт по делу об административном правонарушении")
IDType.AddItem ("4-судебный приказ")
IDType.AddItem ("5-акт органа, осуществляющего контрольные функции")
IDType.AddItem ("6-удостоверение комиссии по трудовым спорам")
IDType.AddItem ("7-акт другого органа")
IDType.AddItem ("10-исполнительная надпись нотариуса")
IDType.AddItem ("11-судебный акт по делу об административном правонарушении")
IDType.AddItem ("12-запрос центрального органа о розыске ребенка")

DebtorType.AddItem (" ")
DebtorType.AddItem ("1-юридическое лицо")
DebtorType.AddItem ("2-физическое лицо")
DebtorType.AddItem ("3-индивидуальный предприниматель")
DebtorType.AddItem ("1696-Российская Федерация")
DebtorType.AddItem ("1681-субъект Российской Федерации")
DebtorType.AddItem ("1682-муниципальное образование")

ClaimerType.AddItem (" ")
ClaimerType.AddItem ("1-юридическое лицо")
ClaimerType.AddItem ("2-физическое лицо")
ClaimerType.AddItem ("3-индивидуальный предприниматель")
ClaimerType.AddItem ("1696-Российская Федерация")
ClaimerType.AddItem ("1681-субъект Российской Федерации")
ClaimerType.AddItem ("1682-муниципальное образование")

IDSubj.AddItem (" ")
IDSubj.AddItem ("1000000-Имущественного характера")
IDSubj.AddItem ("1020000-Возмещение вреда здоровью, причиненного при исполнении трудовых обязанностей")
IDSubj.AddItem ("1030000-Задолженность по детским пособиям")
IDSubj.AddItem ("1040000-Оплата труда и иные выплаты по трудовым правоотношениям")
IDSubj.AddItem ("1050000-Исполнительский сбор")
IDSubj.AddItem ("1060000-Задолженность по платежам за жилую площадь, коммунальные платежи, включая пени")
IDSubj.AddItem ("1080000-Моральный вред как самостоятельное требование")
IDSubj.AddItem ("1090000-Налоги и сборы")
IDSubj.AddItem ("1090100-Госпошлина, присужденная судом")
IDSubj.AddItem ("1090200-Единый налог на вмененный доход (ЕНВД)")
IDSubj.AddItem ("1090300-Земельный налог")
IDSubj.AddItem ("1090400-Взыскание налогов и сборов, включая пени")
IDSubj.AddItem ("1090500-Налог на имущество")
IDSubj.AddItem ("1090600-Пени по налогам и сборам")
IDSubj.AddItem ("1090700-Транспортный налог")
IDSubj.AddItem ("1100000-Взыскание платы за содержание в медицинском вытрезвителе")
IDSubj.AddItem ("1110000-Расходы по совершению исполнительных действий в бюджет")
IDSubj.AddItem ("1120000-Процессуальные издержки в доход государства")
IDSubj.AddItem ("1130000-Таможенные платежи")
IDSubj.AddItem ("1150000-Штраф по законодательству об административных правонарушениях")
IDSubj.AddItem ("1150100-Штраф ГИБДД")
IDSubj.AddItem ("1150200-Иной штраф ОВД")
IDSubj.AddItem ("1150300-Штраф как вид наказания по делам об АП, назначенный судом (за исключением дел по протоколам ФССП)")
IDSubj.AddItem ("1150400-Штраф таможенного органа")
IDSubj.AddItem ("1150500-Штраф иного органа")
IDSubj.AddItem ("1150600-Иной штраф налогового органа по КоАП РФ")
IDSubj.AddItem ("1150700-Штраф по постановлению судебного пристава-исполнителя")
IDSubj.AddItem ("1150800-Штраф как вид наказания по делам об АП (при выдворении), назначенный судом")
IDSubj.AddItem ("1150900-Штраф как вид наказания по делам об АП, назначенный судом по протоколу должностного лица ФССП")
IDSubj.AddItem ("1151000-Штраф по постановлению должностного лица ФССП России")
IDSubj.AddItem ("1151100-Штраф за нарушение в области лесных отношений")
IDSubj.AddItem ("1160000-Иные взыскания имущественного характера не в бюджеты РФ")
IDSubj.AddItem ("1170000-Страховые взносы, включая пени")
IDSubj.AddItem ("1190000-Задолженность по кредитным платежам (кроме ипотеки)")
IDSubj.AddItem ("1200000-Периодические платежи (кроме алиментных платежей)")
IDSubj.AddItem ("1210000-Выплаты в связи со смертью кормильца")
IDSubj.AddItem ("1220000-Обращение взыскания на заложенное имущество")
IDSubj.AddItem ("1230000-Взыскание с инвестиционных и строительных организаций в пользу физических лиц - соинвесторов и вкладчиков")
IDSubj.AddItem ("1240000-Взыскание с финансовых пирамид (в т.ч. банков и иных кредитных организаций и др.) в пользу физических лиц - соинвесторов и вкладчиков")
IDSubj.AddItem ("1250000-Задолженность")
IDSubj.AddItem ("1270000-Задолженность по алиментам")
IDSubj.AddItem ("1280000-Штраф за нарушение законодательства, кроме законодательства об АП")
IDSubj.AddItem ("1280100-Налоговая санкция")
IDSubj.AddItem ("1280200-Иной штраф налогового органа по НК РФ")
IDSubj.AddItem ("1280300-Штраф органа пенсионного фонда")
IDSubj.AddItem ("1280400-Иной штраф (не по КоАП)")
IDSubj.AddItem ("1280500-Штраф таможенного органа по НК РФ")
IDSubj.AddItem ("1280600-Судебные штрафы, наложенные в порядке ст. 105, 106 ГПК РФ")
IDSubj.AddItem ("1280700-Штраф по страховым взносам (ФЗ-212)")
IDSubj.AddItem ("1280800-Денежные взыскания, наложенные в порядке ст. 118 УПК РФ")
IDSubj.AddItem ("1280900-Денежное взыскание за нарушение законодательства о суде и судоустройстве, об ИП и судебные штрафы (код главы по КБК для ФССП - 322)")
IDSubj.AddItem ("1290000-Расходы по совершению исполнительных действий (иные)")
IDSubj.AddItem ("1310000-Уголовные штрафы")
IDSubj.AddItem ("1310100-Уголовный штраф, как основной вид наказания (код главы по КБК для ФССП - 322)")
IDSubj.AddItem ("1310200-Уголовный штраф, как дополнительный вид наказания (код главы по КБК для ФССП - 322)")
IDSubj.AddItem ("1310300-Уголовный штраф, как основной вид наказания (иной администратор дохода с кодом главы по КБК кроме 322)")
IDSubj.AddItem ("1310400-Уголовный штраф, как дополнительный вид наказания (иной администратор дохода с кодом главы по КБК кроме 322)")
IDSubj.AddItem ("1310500-Уголовный штраф за коррупционное преступление, как основной вид наказания")
IDSubj.AddItem ("1310600-Уголовный штраф за коррупционное преступление, как дополнительный вид наказания")
IDSubj.AddItem ("1310700-Уголовный штраф за коррупционное преступление осн. вид наказ. (иной администратор дохода)")
IDSubj.AddItem ("1310800-Уголовный штраф за коррупционное преступление доп. вид наказ. (иной администратор дохода)")
IDSubj.AddItem ("1320000-Иные взыскания имущественного характера в пользу бюджетов Российской Федерации")
IDSubj.AddItem ("1330000-Процессуальные издержки в пользу иных лиц, кроме расходов на экспертизу")
IDSubj.AddItem ("1340000-Задолженность по кредитным платежам (ипотека)")
IDSubj.AddItem ("1350000-Расходы по судебным экcпертизам и экспертным исследованиям")
IDSubj.AddItem ("1360000-Алиментные платежи")
IDSubj.AddItem ("1360200-Алименты на содержание детей")
IDSubj.AddItem ("1360300-Алименты на содержание детей в связи с наличием задолженности по окончании срока уплаты")
IDSubj.AddItem ("1360400-Алименты на содержание детей, находящихся в детских домах и иных учреждениях")
IDSubj.AddItem ("1360500-Алименты на содержание родителей")
IDSubj.AddItem ("1360600-Алименты на содержание супругов")
IDSubj.AddItem ("1360700-Алименты на содержание других членов семьи")
IDSubj.AddItem ("1390000-Материальный ущерб")
IDSubj.AddItem ("1390100-Взыскание ущерба в пользу ФССП России в порядке регресса")
IDSubj.AddItem ("1390200-Материальный ущерб")
IDSubj.AddItem ("1390300-Ущерб за нарушение лесного законодательства")
IDSubj.AddItem ("1390400-Ущерб, причиненный преступлением")
IDSubj.AddItem ("1390500-Ущерб, причиненный административным правонарушением")
IDSubj.AddItem ("1390600-Ущерб, причиненный преступлением (код главы по КБК для ФССП — 322)")
IDSubj.AddItem ("1390700-Ущерб за нарушение лесного законодательства, причиненного правонарушением")
IDSubj.AddItem ("1400000-Возмещение вреда здоровью, причиненного преступлением, правонарушением")
IDSubj.AddItem ("1410000-Конфискация по УК РФ")
IDSubj.AddItem ("1420000-Наложение ареста")
IDSubj.AddItem ("1430000-Снятие ареста")
IDSubj.AddItem ("1440000-Присуждение имущества в натуре")
IDSubj.AddItem ("1450000-Конфискация орудия совершения или предмета административного правонарушения, изъятие из оборота")
IDSubj.AddItem ("1460000-Задолженность по платежам за газ, тепло и электроэнергию")
IDSubj.AddItem ("1470000-Взыскание процентов за пользование чужими денежными средствами")
IDSubj.AddItem ("1480000-Взыскание компенсации в связи с неисполнением должником неденежного требования")
IDSubj.AddItem ("1490000-Задолженность по платежам за услуги связи")
IDSubj.AddItem ("1500000-Задолженность по оплате за использование лесов с арендаторов лесных участков и иных лиц, использующих леса")
IDSubj.AddItem ("2000000-Неимущественного характера")
IDSubj.AddItem ("2010000-Розыск ребенка, незаконно перемещенного в РФ")
IDSubj.AddItem ("2050000-Административное приостановление деятельности")
IDSubj.AddItem ("2060000-Восстановление на работе")
IDSubj.AddItem ("2070000-Вселение")
IDSubj.AddItem ("2080000-Выселение")
IDSubj.AddItem ("2090000-Лишение, ограничение родительских прав")
IDSubj.AddItem ("2100000-Передача (отобрание) ребенка")
IDSubj.AddItem ("2110000-Иной вид исполнения неимущественного характера")
IDSubj.AddItem ("2130000-Обеспечительная мера неимущественного характера")
IDSubj.AddItem ("2160000-Снос самовольно возведенных строений")
IDSubj.AddItem ("2170000-Обязание юридического лица освободить помещение")
IDSubj.AddItem ("2190000-Обязание физического лица освободить помещение")
IDSubj.AddItem ("2200000-Принудительное выдворение за пределы РФ")
IDSubj.AddItem ("2210000-О направлении на принудительное лечение в психиатрический стационар")
IDSubj.AddItem ("2220000-Административное наказание в виде обязательных работ")
IDSubj.AddItem ("2230000-Предоставление жилья, выдача жилищных сертификатов")
IDSubj.AddItem ("2230100-Предоставление жилого помещения иным лицам")
IDSubj.AddItem ("2230200-Предоставление жилья военнослужащим")
IDSubj.AddItem ("2230300-Предоставление жилья детям-сиротам")
IDSubj.AddItem ("2230400-Предоставление жилья нуждающимся в улучшении жилищных условий")
IDSubj.AddItem ("2240000-Определение порядка общения с несовершеннолетними детьми")
IDSubj.AddItem ("2250000-Определение места жительства ребенка")
IDSubj.AddItem ("1281000-Судебный штраф, назначенный в качестве меры уголовно-правового характера (код главы по КБК для ФССП - 322)")
IDSubj.AddItem ("1281100-Судебный штраф, назначенный в качестве меры уголовно-правового характера (иной администратор дохода с кодом главы по КБК кроме 322)")
IDSubj.AddItem ("2260000-Административное наказание в виде обязательных работ, наложенных судом по протоколу должностного лица ФССП России")
IDSubj.AddItem ("1510000-Моральный вред, причиненный преступлением")


SvedDoxodData.AddItem ("2018-03-31")
SvedDoxodData.AddItem ("2017-09-30")
SvedDoxodData.AddItem ("2017-10-31")
SvedDoxodData.AddItem ("2019-03-20")

TypeDoc.AddItem (" ")
TypeDoc.AddItem ("01-паспорт гражданина Союза Советских Социалистических Республик")
TypeDoc.AddItem ("02-загранпаспорт гражданина Союза Советских Социалистических Республик")
TypeDoc.AddItem ("03-свидетельство о рождении")
TypeDoc.AddItem ("04-удостоверение личности офицера")
TypeDoc.AddItem ("05-справка об освобождении из места лишения свободы")
TypeDoc.AddItem ("06-паспорт Минморфлота СССР")
TypeDoc.AddItem ("07-военный билет солдата (матроса, сержанта, старшины)")
TypeDoc.AddItem ("08-временное удостоверение, выданное взамен военного билета")
TypeDoc.AddItem ("09-дипломатический паспорт гражданина Российской Федерации")
TypeDoc.AddItem ("10-иностранный паспорт")
TypeDoc.AddItem ("11-свидетельство о рассмотрении ходатайства о признании беженцем на территории Российской Федерации по существу")
TypeDoc.AddItem ("12-вид на жительство лица без гражданства")
TypeDoc.AddItem ("13-удостоверение беженца в Российской Федерации")
TypeDoc.AddItem ("14-временное удостоверение личности гражданина Российской Федерации")
TypeDoc.AddItem ("19-разрешение на временное проживание в Российской Федерации")
TypeDoc.AddItem ("20-свидетельство о предоставлении временного убежища на территории Российской Федерации")
TypeDoc.AddItem ("21-паспорт гражданина Российской Федерации")
TypeDoc.AddItem ("22-заграничный паспорт гражданина Российской Федерации")
TypeDoc.AddItem ("23-свидетельство о рождении, выданное уполномоченным органом иностранного государства")
TypeDoc.AddItem ("24-удостоверение личности военнослужащего Российской Федерации")
TypeDoc.AddItem ("26-паспорт моряка")
TypeDoc.AddItem ("27-военный билет офицера запаса")
TypeDoc.AddItem ("60-документы, подтверждающие факт регистрации по месту жительства")
TypeDoc.AddItem ("91-иные документы, предусмотренные законодательством Российской Федерации")

Command1_Click

End Sub

Sub openDoc(ByVal filename As String)
    Dim XDoc As Object
    On Error GoTo error_open_doc
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load (filename)
    
    'Get Document Elements
    Set Lists = XDoc.DocumentElement
    
    For Each listNode In Lists.ChildNodes
        Debug.Print "----" & listNode.BaseName & "----"
    
        Select Case listNode.BaseName
            Case "PaymentProperties"
                If (flgPP = False) Then
                    Set pp = New PaymentProperty
                    flgPP = True
                Else
                    dictPP.Add pp
                    idxPP = idxPP + 1
                    Set pp = New PaymentProperty
                End If
        End Select
        
        For Each fieldNode In listNode.ChildNodes
            Debug.Print "[" & fieldNode.BaseName & "] = [" & fieldNode.Text & "]"
            Select Case fieldNode.BaseName
                Case "SvedDoxodData"
                    
                    Set sdd = New SvedDoxodData
                    
                    For Each fNode In fieldNode.ChildNodes
                        Debug.Print "SvedDoxodData____[" & fNode.BaseName & "] = [" & fNode.Text & "]"
                        Select Case fNode.BaseName
                            Case "ActDate"
                                sdd.ActDate = fNode.Text
                            Case "KindData"
                                sdd.KindData = fNode.Text
                            Case "DataDox"
                                sdd.DataDox = fNode.Text
                            Case "SumDox"
                                sdd.SumDox = fNode.Text
                            Case "Ground"
                                sdd.Ground = fNode.Text
                        End Select
                    Next fNode
                    
                    dictSDD.Add sdd
                    idxSDD = idxSDD + 1
                    
                Case "SvedNedvData"
                    Set snd = New SvedNedvData
                    For Each fNode In fieldNode.ChildNodes
                        Debug.Print "SvedNedvData____[" & fNode.BaseName & "] = [" & fNode.Text & "]"
                        Select Case fNode.BaseName
                            Case "ActDate"
                                snd.ActDate = fNode.Text
                            Case "NaimNedv"
                                snd.NaimNedv = fNode.Text
                            Case "KadastrN"
                                snd.KadastrN = fNode.Text
                            Case "AdresNedv"
                                snd.AdresNedv = fNode.Text
                            Case "SNedv"
                                snd.SNedv = fNode.Text
                            Case "RegisterDate"
                                snd.RegisterDate = fNode.Text
                        End Select
                    Next fNode
                    dictSND.Add snd
                    idxSND = idxSND + 1

                Case "TransportData"
                    Set td = New TransportData
                    For Each fNode In fieldNode.ChildNodes
                        Debug.Print "TransportData____[" & fNode.BaseName & "] = [" & fNode.Text & "]"
                        Select Case fNode.BaseName
                            Case "ActDate"
                                td.ActDate = fNode.Text
                            Case "AutomType"
                                td.AutomType = fNode.Text
                            Case "RegNo"
                                td.RegNo = fNode.Text
                            Case "Producer"
                                td.Producer = fNode.Text
                            Case "VIN"
                                td.vin = fNode.Text
                            Case "Engine"
                                td.Engine = fNode.Text
                            Case "MadeYear"
                                td.MadeYear = fNode.Text
                        End Select
                    Next fNode
                    dictTD.Add td
                    idxTD = idxTD + 1
                Case "IdentificationData"
                    Set id = New IdentificationData
                    
                    For Each fNode In fieldNode.ChildNodes
                        Debug.Print "IdentificationData____[" & fNode.BaseName & "] = [" & fNode.Text & "]"
                        Select Case fNode.BaseName
                            Case "ActDate"
                                id.ActDate = fNode.Text
                            Case "KindData"
                                id.KindData = fNode.Text
                            Case "TypeDoc"
                                id.TypeDoc = fNode.Text
                            Case "SerDoc"
                                id.SerDoc = fNode.Text
                            Case "NumDoc"
                                id.NumDoc = fNode.Text
                            Case "IssuedDoc"
                                id.IssuedDoc = fNode.Text
                            Case "DateDoc"
                                id.DateDoc = fNode.Text
                            Case "FIODoc"
                                For Each ffNode In fNode.ChildNodes
                                Debug.Print "FIODoc____[" & fNode.BaseName & "] = [" & fNode.Text & "]"
                                    Select Case ffNode.BaseName
                                        Case "Surname"
                                            id.Surname = ffNode.Text
                                        Case "FirstName"
                                            id.FirstName = ffNode.Text
                                        Case "Patronymic"
                                            id.Patronymic = ffNode.Text
                                    End Select
                                Next ffNode
                        End Select
                    Next fNode
                    
                    dictID.Add id ', idxID
                    idxID = idxID + 1
                Case "AddressData"
                    Set ad = New AddressData
                    
                    For Each fNode In fieldNode.ChildNodes
                        Debug.Print "AddressData____[" & fNode.BaseName & "] = [" & fNode.Text & "]"
                        Select Case fNode.BaseName
                            Case "ActDate"
                                ad.ActDate = fNode.Text
                            Case "StrAddr"
                                ad.StrAddr = fNode.Text
                            Case "RegDate"
                                ad.regDate = fNode.Text
                            Case "Address"
                                For Each ffNode In fNode.ChildNodes
                                Debug.Print "Address____[" & fNode.BaseName & "] = [" & fNode.Text & "]"
                                    Select Case ffNode.BaseName
                                        Case "addressText"
                                            ad.addressText = ffNode.Text
                                        Case "area"
                                            ad.area = ffNode.Text
                                        Case "city"
                                            ad.city = ffNode.Text
                                        Case "countryCode"
                                            ad.countryCode = ffNode.Text
                                        Case "flatNumber"
                                            ad.flatNumber = ffNode.Text
                                        Case "houseNumber"
                                            ad.houseNumber = ffNode.Text
                                        Case "OKATO"
                                            ad.Okato = ffNode.Text
                                        Case "OKTMO"
                                            ad.Oktmo = ffNode.Text
                                        Case "street"
                                            ad.street = ffNode.Text
                                        Case "zipCode"
                                            ad.zipCode = ffNode.Text
                                    End Select
                                    Next ffNode
                        End Select
                    Next fNode
                    
                    dictAD.Add ad
                    idxAD = idxAD + 1

            End Select
            Call upd(listNode.BaseName, fieldNode.BaseName, fieldNode.Text)
        Next fieldNode
    Next listNode
 
    If (flgPP = True) Then
        dictPP.Add pp
    End If
    
    PaymentProperties.Clear
    For Each it In dictPP
        Dim p  As PaymentProperty
        Set p = it
        PaymentProperties.AddItem p.RecpName
        RecpName.Text = p.RecpName
        RecpBank.Text = p.RecpBank
        RecpBIK.Text = p.RecpBIK
        RecpCnt.Text = p.RecpCnt
        RecpINN.Text = p.RecpINN
        RecpKPP.Text = p.RecpKPP
        Okato.Text = p.Okato
        Oktmo.Text = p.Oktmo
        Kbk.Text = p.Kbk
        PokPl.Text = p.PokPl
        RekvSum.Text = p.RekvSum
    Next it
 
    SvedDoxodData.Clear
    For Each it In dictSDD
        Dim d As SvedDoxodData
        Set d = it
        SvedDoxodData.AddItem d.DataDox
        D_ActDate.Text = d.ActDate
        DataDox.Text = d.DataDox
        Ground.Text = d.Ground
        SumDox.Text = d.SumDox
    Next it
    
    AddressData.Clear
    For Each it In dictAD
        Dim d1 As AddressData
        Set d1 = it
        AddressData.AddItem d1.addressText
        AD_ActDate.Text = d1.ActDate
        AD_addressText.Text = d1.addressText
        AD_area.Text = d1.area
        AD_city.Text = d1.city
        AD_countryCode.Text = d1.countryCode
        AD_flatNumber.Text = d1.flatNumber
        AD_houseNumber.Text = d1.houseNumber
        AD_OKATO.Text = d1.Okato
        AD_OKTMO.Text = d1.Oktmo
        AD_street.Text = d1.street
        AD_zipCode.Text = d1.zipCode
        AD_RegDate.Text = d1.regDate
        AD_StrAddr.Text = d1.StrAddr
    Next it
    
    IdentificationData.Clear
    For Each it In dictID
        Dim d2 As IdentificationData
        Set d2 = it
        IdentificationData.AddItem d2.NumDoc
        Doc_ActDate.Text = d2.ActDate
        Doc_DateDoc.Text = d2.DateDoc
        Doc_FirstName.Text = d2.FirstName
        Doc_IssuedDoc.Text = d2.IssuedDoc
        Doc_NumDoc.Text = d2.NumDoc
        Doc_Patronymic.Text = d2.Patronymic
        Doc_SerDoc.Text = d2.SerDoc
        Doc_Surname.Text = d2.Surname
        setTypeDoc (d2.TypeDoc)
    Next it
    
    SvedNedvData.Clear
    For Each it In dictSND
        Dim d3 As SvedNedvData
        Set d3 = it
        SvedNedvData.AddItem d3.NaimNedv
        SND_ActDate.Text = d3.ActDate
        AdresNedv.Text = d3.AdresNedv
        KadastrN.Text = d3.KadastrN
        NaimNedv.Text = d3.NaimNedv
        RegisterDate.Text = d3.RegisterDate
        SNedv.Text = d3.SNedv
    Next it
    
    TransportData.Clear
    For Each it In dictTD
        Dim d4 As TransportData
        Set d4 = it
        TransportData.AddItem d4.AutomType
        Tr_ActDate.Text = d4.ActDate
        AutomType.Text = d4.AutomType
        Engine.Text = d4.Engine
        MadeYear.Text = d4.MadeYear
        Producer.Text = d4.Producer
        RegNo.Text = d4.RegNo
        vin.Text = d4.vin
    Next it
    
    Call Show
error_open_doc:
End Sub

Private Sub upd(ByVal name As String, ByVal subname As String, ByVal value As String)
Select Case name
    Case "IDType"
        setIDTypeValue (value)
    Case "IDNum"
        IDNum.Text = value
    Case "ExternalKey"
        ExternalKey.Text = value
    Case "DocDate"
        DocDate.Text = value
    Case "IDDate"
        IDDate.Text = value
    Case "IDSubj"
        IDSubj.Text = value
    Case "IDSubjName"
        IDSubjName.Text = value
    Case "IDSum"
        IDSum.Text = value
    Case "DeloNum"
        DeloNum.Text = value
    Case "DeloDate"
        DeloDate.Text = value
    Case "OrganCode"
        OrganCode.Text = value
    Case "Organ"
        Organ.Text = value
    Case "OrganAdr"
        OrganAdr.Text = value
    Case "OrganSignPost"
        OrganSignPost.Text = value
    Case "OrganSignFIO"
        OrganSignFIO.Text = value
    Case "AktDate"
        AktDate.Text = value
    Case "DebtorType"
        'DebtorType.text = valus
        setDebtorType (value)
    Case "DebtorName"
        DebtorName.Text = value
    Case "Surname"
        Surname.Text = value
    Case "FirstName"
        FirstName.Text = value
    Case "Patronymic"
        Patronymic.Text = value
    Case "DebtorAdr"
        DebtorAdr.Text = value
    Case "DebtorBirthDate"
        DebtorBirthDate.Text = value
    Case "DebtorBirthPlace"
        DebtorBirthPlace.Text = value
    Case "DebtorINN"
        DebtorINN.Text = value
    Case "DebtorRegDate"
        DebtorRegDate.Text = value
    Case "ClaimerType"
        'ClaimerType.text = value
        setClaimerType (value)
    Case "ClaimerName"
        ClaimerName.Text = value
    Case "ClaimerAdr"
        ClaimerAdr.Text = value
    Case Else
        Debug.Print ("Поле не определено " & name)
End Select

Select Case subname
    Case "RecpName"
        pp.RecpName = value
    Case "RecpBank"
        pp.RecpBank = value
    Case "RecpBIK"
        pp.RecpBIK = value
    Case "RecpCnt"
        pp.RecpCnt = value
    Case "RecpINN"
        pp.RecpINN = value
    Case "RecpKPP"
        pp.RecpKPP = value
    Case "Okato"
        pp.Okato = value
    Case "Oktmo"
        pp.Oktmo = value
    Case "Kbk"
        pp.Kbk = value
    Case "PokPl"
        pp.PokPl = value
    Case "RekvSum"
        pp.RekvSum = value
    Case "SignedInfo"
        SignedInfo.Text = value
    Case "SignatureValue"
        SignatureValue.Text = value
    Case "KeyInfo"
        X509Certificate.Text = value
    Case "Surname"
        Surname.Text = value
    Case "FirstName"
        FirstName.Text = value
    Case "Patronymic"
        Patronymic.Text = value
End Select

End Sub

Private Sub setIDTypeValue(ByVal value As String)
IDType.Text = value
'Dim selvalue As String
'For Each l In IDType
' If InStr(l, value) > 0 Then
'    selvalue = l
'    Exit For
' End If
' Next l
' IDType.Text = selvalue
End Sub

Private Sub setTypeDoc(ByVal value As String)
TypeDoc.Text = value
'Dim selvalue As String
'For Each l In TypeDoc
' If InStr(l, value) > 0 Then
'    selvalue = l
'    Exit For
' End If
' Next l
' TypeDoc.Text = selvalue
End Sub

Private Sub setClaimerType(ByVal value As String)
ClaimerType.Text = value
'Dim selvalue As String
'For Each l In ClaimerType
' If InStr(l, value) > 0 Then
'    selvalue = l
'    Exit For
' End If
' Next l
' ClaimerType.Text = selvalue
End Sub

Private Sub setDebtorType(ByVal value As String)
DebtorType.Text = value
'Dim selvalue As String
'For Each l In DebtorType
' If InStr(l, value) > 0 Then
'    selvalue = l
'    Exit For
' End If
' Next l
' DebtorType.Text = selvalue
End Sub


Private Sub IdentificationData_Click()
    IdentificationData_change
End Sub

Private Sub IdentificationData_KeyPress(KeyAscii As Integer)
    IdentificationData_change
End Sub

Private Sub PaymentProperties_Click()
    PaymentProperties_Change
End Sub

Private Sub PaymentProperties_KeyPress(KeyAscii As Integer)
    PaymentProperties_Change
End Sub

Private Sub SvedDoxodData_Click()
    SvedDoxodData_Change
End Sub

Private Sub SvedDoxodData_KeyPress(KeyAscii As Integer)
    SvedDoxodData_Change
End Sub

Private Sub SvedNedvData_Click()
    SvedNedvData_change
End Sub

Private Sub TransportData_change()
      Dim i As Integer
      Dim selectedTD As TransportData
      
      For i = 0 To TransportData.ListCount - 1
          If TransportData.Selected(i) Then
              Set selectedTD = dictTD(i + 1)
              Exit For
          End If
      Next i

Tr_ActDate.Text = selectedTD.ActDate
AutomType.Text = selectedTD.AutomType
RegNo.Text = selectedTD.RegNo
Producer.Text = selectedTD.Producer
vin.Text = selectedTD.vin
Engine.Text = selectedTD.Engine
MadeYear.Text = selectedTD.MadeYear

End Sub

Private Sub SvedNedvData_KeyPress(KeyAscii As Integer)
    SvedNedvData_change
End Sub

Private Sub TransportData_Click()
    TransportData_change
End Sub

Private Sub TransportData_KeyPress(KeyAscii As Integer)
    TransportData_change
End Sub
