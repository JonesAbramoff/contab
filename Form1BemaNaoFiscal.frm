VERSION 5.00
Begin VB.Form Form1BemaNaoFiscal 
   Caption         =   "Teste de Bematech Não Fiscal"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Modelo 
      Height          =   330
      Left            =   1335
      TabIndex        =   3
      Text            =   "7"
      Top             =   420
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3195
      TabIndex        =   2
      Top             =   1170
      Width           =   855
   End
   Begin VB.TextBox Porta 
      Height          =   285
      Left            =   1290
      TabIndex        =   1
      Top             =   1170
      Width           =   1530
   End
   Begin VB.Label Label2 
      Caption         =   "Modelo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   540
      TabIndex        =   4
      Top             =   450
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Porta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   435
      TabIndex        =   0
      Top             =   1185
      Width           =   630
   End
End
Attribute VB_Name = "Form1BemaNaoFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Dim i As Integer, sTexto As String

    i = BematechNaoFiscal_ConfiguraModeloImpressora(CInt(Modelo.Text))
    MsgBox (CStr(i))
    i = BematechNaoFiscal_IniciaPorta(Trim(Porta.Text))
    MsgBox (CStr(i))
   sTexto = "Total bruto: 12.500,00" + Chr(10)
    sTexto = sTexto + "Total líquido: 9.600,00" + Chr(10)
    i = BematechNaoFiscal_BematechTX(sTexto)
    MsgBox (CStr(i))

    i = BematechNaoFiscal_FormataTX("12345678901234567890123456789012345678901234567890123456789012345678901234567890" & Chr$(10), 1, 0, 0, 0, 0)
    MsgBox (CStr(i))
    i = BematechNaoFiscal_FormataTX(sTexto, 2, 0, 0, 0, 0)
    MsgBox (CStr(i))
    i = BematechNaoFiscal_FormataTX(sTexto, 3, 0, 0, 0, 0)
    MsgBox (CStr(i))
    
'    i = BematechNaoFiscal_ImprimeCodigoQRCODE(1, 10, 0, 10, 1, "http://www4.fazenda.rj.gov.br/consultaNFCe/QRCode?chNFe=33141173841488000153654030000002361018561963&nVersao=100&tpAmb=2&dhEmi=323031342d31312d30355431323a34363a35372d30323a3030&vNF=6.17&vICMS=0.00&digVal=426e7432695a47754a586e4e7855415834586e5a7376382b41626f3d&cIdToken=000001&cHashQRCode=ab09c149f29ecd560561213def0ea328ab337a42")
'    MsgBox (CStr(i))
    
    i = BematechNaoFiscal_FechaPorta
    MsgBox (CStr(i))

End Sub

