VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ProgData 
   Caption         =   "Programação das Datas"
   ClientHeight    =   1665
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1665
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Fator"
      Height          =   555
      Left            =   195
      TabIndex        =   2
      Top             =   465
      Width           =   4185
      Begin VB.OptionButton OptCorridos 
         Caption         =   "Dias corridos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2055
         TabIndex        =   4
         Top             =   255
         Width           =   1695
      End
      Begin VB.OptionButton OptUteis 
         Caption         =   "Dias úteis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   525
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2460
      Picture         =   "ProgData.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1095
      Width           =   990
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1035
      Picture         =   "ProgData.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1005
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Left            =   2370
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   225
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   300
      Left            =   1245
      TabIndex        =   6
      Top             =   120
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NumDias 
      Height          =   285
      Left            =   3795
      TabIndex        =   8
      Top             =   105
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      Caption         =   "Núm.Dias:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   28
      Left            =   2820
      TabIndex        =   9
      Top             =   135
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data Base:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   225
      TabIndex        =   7
      Top             =   150
      Width           =   960
   End
End
Attribute VB_Name = "ProgData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gobjControle As Object
Dim iAlterado As Integer

Private Sub BotaoCancela_Click()
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbCancel
    'Fecha a tela
    Unload Me
End Sub

Private Sub BotaoOK_Click()
    
Dim lErro As Long
Dim dtData As Date
    
On Error GoTo Erro_BotaoOK_Click
    
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbOK
    
    Call Calcula_Data(dtData)
    
    Call DateParaMasked(gobjControle, dtData)
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208700)

    End Select

    Exit Sub
    
End Sub

Private Sub Calcula_Data(dtData As Date)
    
Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objParc As New ClassCondicaoPagtoParc

On Error GoTo Erro_Calcula_Data
      
    objCondicaoPagto.colParcelas.Add objParc
    
    objParc.iTipoDataBase = CONDPAGTO_TIPODATABASE_EMISSAO
    objCondicaoPagto.dtDataEmissao = StrParaDate(Data.Text)
    
    objParc.iModificador = 0
    objParc.iDias = StrParaInt(NumDias.Text)
    
    If OptUteis.Value Then
        objParc.iTipoIntervalo = CONDPAGTO_TIPOINTERVALO_DIAS_UTEIS
    Else
        objParc.iTipoIntervalo = CONDPAGTO_TIPOINTERVALO_DIAS
    End If
    
    objParc.dPercReceb = 1

    lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True, False)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    dtData = objParc.dtVencimento
    
    Exit Sub
    
Erro_Calcula_Data:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208702)

    End Select

    Exit Sub
    
End Sub

Function Trata_Parametros(ByVal objControle As Object, ByVal dtDataBase As Date, ByVal iDiasPadrao As Integer, ByVal iPadraoUteisCorridos As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjControle = objControle
    
    Call DateParaMasked(Data, dtDataBase)
    
    NumDias.PromptInclude = False
    NumDias.Text = CStr(iDiasPadrao)
    NumDias.PromptInclude = True
    
    If iPadraoUteisCorridos = CONDPAGTO_TIPOINTERVALO_DIAS_UTEIS Then
        OptUteis.Value = True
    Else
        OptCorridos.Value = True
    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208703)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)
    Set gobjControle = Nothing
End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190610)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    Data.SetFocus

    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190612)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) <> 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190614)

    End Select

    Exit Sub

End Sub
