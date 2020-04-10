VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl PlanoContaRefOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10665
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   10665
   Begin VB.Frame Frame3 
      Caption         =   "Plano de Contas Referencial"
      Height          =   4905
      Left            =   60
      TabIndex        =   18
      Top             =   1005
      Width           =   4245
      Begin MSComctlLib.TreeView TvwContas 
         Height          =   4575
         Left            =   30
         TabIndex        =   1
         Top             =   270
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   8070
         _Version        =   393217
         Indentation     =   453
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contas"
      Height          =   4905
      Left            =   4320
      TabIndex        =   15
      Top             =   1005
      Width           =   6315
      Begin VB.TextBox ContaImp 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1035
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   150
         Width           =   2520
      End
      Begin VB.TextBox Conta 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1035
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   150
         Width           =   2520
      End
      Begin VB.Frame Frame4 
         Caption         =   "Validade"
         Height          =   555
         Left            =   1035
         TabIndex        =   28
         Top             =   1455
         Width           =   5250
         Begin VB.Label ValidadeAte 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3405
            TabIndex        =   32
            Top             =   180
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
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
            Height          =   195
            Index           =   5
            Left            =   3000
            TabIndex        =   31
            Top             =   225
            Width           =   345
         End
         Begin VB.Label ValidadeDe 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1200
            TabIndex        =   30
            Top             =   180
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "De:"
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
            Height          =   195
            Index           =   4
            Left            =   795
            TabIndex        =   29
            Top             =   225
            Width           =   315
         End
      End
      Begin VB.TextBox DescConta 
         BackColor       =   &H8000000F&
         Height          =   930
         Left            =   1035
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   495
         Width           =   5235
      End
      Begin VB.Frame GrupoConfig 
         Caption         =   "Grupos de Contas Associadas"
         Height          =   2760
         Left            =   45
         TabIndex        =   19
         Top             =   2100
         Width           =   6240
         Begin VB.CommandButton BotaoAlterar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   4845
            Picture         =   "PlanoContaRefOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CommandButton BotaoContas 
            Caption         =   "Contas Final"
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
            Index           =   1
            Left            =   75
            TabIndex        =   5
            Top             =   2475
            Width           =   1350
         End
         Begin VB.CommandButton BotaoCcl 
            Caption         =   "Ccl Final"
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
            Index           =   1
            Left            =   1560
            TabIndex        =   7
            Top             =   2475
            Width           =   1350
         End
         Begin VB.CheckBox Subtrai 
            Height          =   195
            Left            =   4515
            TabIndex        =   24
            Top             =   1410
            Width           =   525
         End
         Begin VB.CommandButton BotaoCcl 
            Caption         =   "Ccl Inicial"
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
            Index           =   0
            Left            =   1575
            TabIndex        =   6
            Top             =   2175
            Width           =   1350
         End
         Begin VB.CommandButton BotaoContas 
            Caption         =   "Contas Inicial"
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
            Index           =   0
            Left            =   75
            TabIndex        =   4
            Top             =   2175
            Width           =   1350
         End
         Begin MSMask.MaskEdBox CclFim 
            Height          =   225
            Left            =   3645
            TabIndex        =   20
            Top             =   1035
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CclInicio 
            Height          =   225
            Left            =   2820
            TabIndex        =   21
            Top             =   1050
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ContaFim 
            Height          =   225
            Left            =   1500
            TabIndex        =   22
            Top             =   1065
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ContaInicio 
            Height          =   225
            Left            =   180
            TabIndex        =   23
            Top             =   1080
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridContas 
            Height          =   645
            Left            =   30
            TabIndex        =   3
            Top             =   195
            Width           =   6180
            _ExtentX        =   10901
            _ExtentY        =   1138
            _Version        =   393216
         End
      End
      Begin VB.Label TipoConta 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4425
         TabIndex        =   27
         Top             =   150
         Width           =   1860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Height          =   195
         Index           =   7
         Left            =   3960
         TabIndex        =   26
         Top             =   195
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Height          =   195
         Index           =   3
         Left            =   75
         TabIndex        =   25
         Top             =   555
         Width           =   945
      End
      Begin VB.Label LabelConta 
         AutoSize        =   -1  'True
         Caption         =   "Conta:"
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
         Height          =   195
         Left            =   435
         TabIndex        =   16
         Top             =   225
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modelo"
      Height          =   915
      Left            =   60
      TabIndex        =   12
      Top             =   45
      Width           =   8865
      Begin VB.TextBox Tipo 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1140
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   525
         Width           =   7680
      End
      Begin VB.TextBox AnoVigencia 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   4755
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   180
         Width           =   660
      End
      Begin VB.TextBox DescricaoModelo 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1140
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   525
         Visible         =   0   'False
         Width           =   7680
      End
      Begin VB.TextBox CodigoModelo 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1140
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   180
         Width           =   1050
      End
      Begin VB.CheckBox Oficial 
         Caption         =   "Oficial para empresa"
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
         Left            =   6060
         TabIndex        =   0
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   38
         Top             =   570
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Válido a partir (ano):"
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
         Height          =   195
         Index           =   1
         Left            =   2985
         TabIndex        =   17
         Top             =   225
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   570
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label LabelCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Height          =   195
         Left            =   390
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   13
         Top             =   195
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8970
      ScaleHeight     =   495
      ScaleWidth      =   1590
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   105
      Width           =   1650
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "PlanoContaRefOcx.ctx":1926
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "PlanoContaRefOcx.ctx":1A80
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1095
         Picture         =   "PlanoContaRefOcx.ctx":1FB2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "PlanoContaRefOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoConta As AdmEvento
Attribute objEventoConta.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1

Dim gobjPlanoContaRefModelo As ClassPlanoContaRefModelo

Dim giCcl As Integer
Dim giConta As Integer

Dim objGridContas As AdmGrid
Dim iGrid_ContaInicio_Col As Integer
Dim iGrid_ContaFinal_Col As Integer
Dim iGrid_CclInicio_Col As Integer
Dim iGrid_CclFinal_Col As Integer
Dim iGrid_Subtrai_Col As Integer

Dim iAlterado As Integer

Private Sub BotaoAlterar_Click()

Dim lErro As Long
Dim objContaRef As New ClassPlanoContaRef
Dim objContaConfig As New ClassPlanoContaRefConfig
Dim objContaRefAux As ClassPlanoContaRef
Dim objNode As Node
Dim sContaMascarada As String
Dim bContaNova As Boolean
Dim iContaPreenchida As Integer
Dim iIndice As Integer
Dim sConta As String

On Error GoTo Erro_BotaoAlterar_Click
       
    lErro = Mover_Conta_Memoria(objContaRef, bContaNova)
    If lErro <> SUCESSO Then gError 200891
   
    If bContaNova Then gError 200892 ' Só se pode alterar contas que já existam

    sConta = objContaRef.sConta

    Set objContaRefAux = gobjPlanoContaRefModelo.colContas(sConta)
    
    objContaRefAux.sDescricao = objContaRef.sDescricao
    objContaRefAux.iTipo = objContaRef.iTipo
    
    For iIndice = objContaRefAux.colConfig.Count To 1 Step -1
        objContaRefAux.colConfig.Remove iIndice
    Next
    For Each objContaConfig In objContaRef.colConfig
        objContaRefAux.colConfig.Add objContaConfig
    Next
    
'    If objContaRef.iTipo = CONTA_ANALITICA Then
'        sConta = "A" & objContaRef.sConta
'    Else
'        sConta = "S" & objContaRef.sConta
'    End If
    sConta = "X" & objContaRef.sConta
    
    Set objNode = TvwContas.Nodes(sConta)

    'coloca a conta no formato que é exibida na tela
    'lErro = Mascara_Mascarar_ContaRef(objContaRef.sConta, sContaMascarada)
    'If lErro <> SUCESSO Then gError 200893
        
    'objNode.Text = sContaMascarada & SEPARADOR & objContaRef.sDescricao
    
    objNode.Text = objContaRefAux.sContaImp & SEPARADOR & objContaRefAux.sDescricao
    
    Call Limpa_Conta

    Exit Sub
    
Erro_BotaoAlterar_Click:
   
    Select Case gErr
    
        Case 200890, 200891, 200893
        
        Case 200892
            Call Rotina_Erro(vbOKOnly, "ERRO_ALT_CONTAREF_NAO_EXISTE", gErr)
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200894)
        
    End Select

    Exit Sub
    
End Sub

'Private Sub BotaoRemover_Click()
'
'Dim lErro As Long
'Dim objContaRef As New ClassPlanoContaRef
'Dim objContaRefAux As New ClassPlanoContaRef
'Dim objNode As Node
'Dim sConta As String
'Dim sContaMascarada As String
'Dim bContaNova As Boolean
'Dim iContaPreenchida As Integer
'Dim iIndice As Integer
'Dim bAchou As Boolean
'Dim iNivelConta As Integer
'Dim iNivelContaPai As Integer
'Dim vbConfirma As VbMsgBoxResult
'
'On Error GoTo Erro_BotaoRemover_Click
'
'    lErro = Mover_Conta_Memoria(objContaRef, bContaNova)
'    If lErro <> SUCESSO Then gError 200896
'
'    If bContaNova Then gError 200897 ' Só se pode excluir contas que já existam
'
'    sConta = objContaRef.sConta
'
'    iIndice = 0
'    bAchou = False
'    For Each objContaRefAux In gobjPlanoContaRefModelo.colContas
'        iIndice = iIndice + 1
'        If objContaRefAux.sConta = objContaRef.sConta Then
'            If objContaRefAux.iTipo <> CONTA_ANALITICA Then
'                vbConfirma = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CONTAREF_SINTETICA")
'                If vbConfirma = vbNo Then gError 200898
'            End If
'            'gobjPlanoContaRefModelo.colContas.Remove iIndice
'            Call ContaRef_Nivel(objContaRefAux.sConta, iNivelContaPai)
'            bAchou = True
'        End If
'        If bAchou Then
'            'Remove os filhos
'            Call ContaRef_Nivel(objContaRefAux.sConta, iNivelConta)
'            If iNivelConta > iNivelContaPai Or objContaRefAux.sConta = objContaRef.sConta Then
'                gobjPlanoContaRefModelo.colContas.Remove iIndice
'                iIndice = iIndice - 1
'            Else
'                Exit For
'            End If
'        End If
'    Next
'
'    'REMOVER A CONTA
'    Call Exclui_Arvore_Conta(objContaRef)
'
'    Call Limpa_Conta
'
'    Exit Sub
'
'Erro_BotaoRemover_Click:
'
'    Select Case gErr
'
'        Case 200895, 200896, 200898
'
'        Case 200897
'            Call Rotina_Erro(vbOKOnly, "ERRO_EXC_CONTAREF_NAO_EXISTE", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200899)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoIncluir_Click()
'
'Dim lErro As Long
'Dim objContaRef As New ClassPlanoContaRef
'Dim objNode As Node
'Dim sConta As String
'Dim sContaMascarada As String
'Dim bContaNova As Boolean
'Dim sContaPai As String
'Dim iContaPreenchida As Integer
'Dim colSaida As New Collection
'Dim colCampos As New Collection
'Dim iIndice As Integer
'
'On Error GoTo Erro_BotaoIncluir_Click
'
'    lErro = Mover_Conta_Memoria(objContaRef, bContaNova)
'    If lErro <> SUCESSO Then gError 200901
'
'    If Not bContaNova Then gError 200902 ' Só se pode incluir contas que não existam
'
'    sConta = objContaRef.sConta
'
'    ''coloca a conta no formato que é exibida na tela
'    'lErro = Mascara_Mascarar_ContaRef(objContaRef.sConta, sContaMascarada)
'    'If lErro <> SUCESSO Then gError 200903
'    sContaMascarada = objContaRef.sConta
'
'     If objContaRef.iTipo = CONTA_ANALITICA Then
'         sConta = "A" & objContaRef.sConta
'     Else
'         sConta = "S" & objContaRef.sConta
'     End If
'
'     sContaPai = ""
'
'     lErro = Retorna_ContaRef_Pai(objContaRef.sConta, sContaPai)
'     If lErro <> SUCESSO Then gError 200904
'
'     If Len(Trim(sContaPai)) = 0 Then
'
'         Set objNode = TvwContas.Nodes.Add(, tvwLast, sConta, sContaMascarada & SEPARADOR & objContaRef.sDescricao)
'        TvwContas.Sorted = True
'
'     Else
'
'        sContaPai = "S" & sContaPai
'
'        Set objNode = TvwContas.Nodes.Add(TvwContas.Nodes.Item(sContaPai), tvwChild, sConta, sContaMascarada & SEPARADOR & objContaRef.sDescricao)
'        TvwContas.Nodes.Item(sContaPai).Sorted = True
'
'    End If
'
'    objNode.Tag = objContaRef.sConta
'
'    colCampos.Add "sConta"
'
'    gobjPlanoContaRefModelo.colContas.Add objContaRef
'
'    Call Ordena_Colecao(gobjPlanoContaRefModelo.colContas, colSaida, colCampos)
'
'    For iIndice = colSaida.Count To 1 Step -1
'        gobjPlanoContaRefModelo.colContas.Remove iIndice
'    Next
'    For Each objContaRef In colSaida
'        gobjPlanoContaRefModelo.colContas.Add objContaRef, objContaRef.sConta
'    Next
'
'    Call Limpa_Conta
'
'    Exit Sub
'
'Erro_BotaoIncluir_Click:
'
'    Select Case gErr
'
'        Case 200900, 200901, 200903, 200904
'
'        Case 200902
'            Call Rotina_Erro(vbOKOnly, "ERRO_INC_CONTAREF_EXISTE", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200905)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoExcluir_Click()
'
'Dim lErro As Long
'Dim vbMsgRes As VbMsgBoxResult
'Dim objPlanoContaRefModelo As New ClassPlanoContaRefModelo
'
'On Error GoTo Erro_BotaoExcluir_Click
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    '#####################
'    'CRITICA DADOS DA TELA
'    If Len(Trim(CodigoModelo.Text)) = 0 Then gError 200906
'    '#####################
'
'    objPlanoContaRefModelo.lCodigo = StrParaLong(CodigoModelo.Text)
'
'    'Pergunta ao usuário se confirma a exclusão
'    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PLANOCONTAREF")
'
'    If vbMsgRes = vbYes Then
'
'        'Exclui a requisição de consumo
'        lErro = CF("PlanoContaRefModelo_Exclui", objPlanoContaRefModelo)
'        If lErro <> SUCESSO Then gError 200907
'
'        'Limpa Tela
'        Call Limpa_Tela_PlanoConta
'
'    End If
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'Erro_BotaoExcluir_Click:
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case gErr
'
'        Case 200906
'            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
'            CodigoModelo.SetFocus
'
'        Case 200907
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200908)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub CodigoModelo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodigoModelo_Validate

    'Verifica se CodigoModelo está preenchida
    If Len(Trim(CodigoModelo.Text)) <> 0 Then

       'Critica a CodigoModelo
       lErro = Long_Critica(CodigoModelo.Text)
       If lErro <> SUCESSO Then gError 200909

    End If

    Exit Sub

Erro_CodigoModelo_Validate:

    Cancel = True

    Select Case gErr

        Case 200909

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200910)

    End Select

    Exit Sub

End Sub

'Private Sub CodigoModelo_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(CodigoModelo, iAlterado)
'
'End Sub

Private Sub CodigoModelo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 200911
    
    Call Limpa_Tela_PlanoConta
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 200911
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200912)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(CodigoModelo.Text)) = 0 Then gError 200913
    If StrParaInt(AnoVigencia.Text) = 0 Then gError 200914
    If Len(Trim(DescricaoModelo.Text)) = 0 Then gError 200915
    
    'Preenche o Modelo
    lErro = Mover_Tela_Memoria(gobjPlanoContaRefModelo)
    If lErro <> SUCESSO Then gError 200916

    lErro = Trata_Alteracao(gobjPlanoContaRefModelo, gobjPlanoContaRefModelo.lCodigo)
    If lErro <> SUCESSO Then gError 200917
    
    'Grava o Plano de Conta Referencial
    lErro = CF("PlanoContaRefModelo_Grava", gobjPlanoContaRefModelo)
    If lErro <> SUCESSO Then gError 200918

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault
       
     Select Case gErr
     
        Case 200913
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTAREF_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 200914
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTAREF_ANO_NAO_PREENCHIDO", gErr)
        
        Case 200915
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTAREF_DESCRICAO_NAO_PREENCHIDO", gErr)
            
        Case 200916, 200917, 200918
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200919)
        
    End Select

    Exit Function

End Function
        
Private Function Mover_Tela_Memoria(ByVal objPlanoContaRefModelo As ClassPlanoContaRefModelo) As Long

Dim lErro As Long

On Error GoTo Erro_Mover_Tela_Memoria

    gobjPlanoContaRefModelo.lCodigo = StrParaLong(CodigoModelo.Text)
    gobjPlanoContaRefModelo.sDescricao = DescricaoModelo.Text
    
    If Oficial.Value = vbChecked Then
        gobjPlanoContaRefModelo.iOficial = MARCADO
    Else
        gobjPlanoContaRefModelo.iOficial = DESMARCADO
    End If
    gobjPlanoContaRefModelo.iAnoVigencia = StrParaInt(AnoVigencia.Text)
    gobjPlanoContaRefModelo.iTipo = Codigo_Extrai(Tipo.Text)

    Mover_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Mover_Tela_Memoria:

    Mover_Tela_Memoria = gErr
    
    Select Case gErr
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200920)
        
    End Select

    Exit Function
        
End Function

Private Function Mover_Conta_Memoria(ByVal objPlanoContaRef As ClassPlanoContaRef, bContaNova As Boolean) As Long

Dim lErro As Long
Dim objConfig As ClassPlanoContaRefConfig
Dim iLinha As Integer
Dim sConta As String
Dim iContaPreenchida As Integer
Dim sCcl As String
Dim iCclPreenchida As Integer
Dim iIndiceConta As Integer

On Error GoTo Erro_Mover_Conta_Memoria

    'lErro = ContaRef_Formata(Conta.Text, sConta, iContaPreenchida)
    'If lErro <> SUCESSO Then gError 200921
    sConta = Conta.Text
    
    objPlanoContaRef.sConta = sConta
    objPlanoContaRef.sDescricao = DescConta.Text
    objPlanoContaRef.lCodigoModelo = StrParaLong(CodigoModelo.Text)
    objPlanoContaRef.iTipo = Codigo_Extrai(TipoConta.Caption)
    
    For iLinha = objPlanoContaRef.colConfig.Count To 1 Step -1
        objPlanoContaRef.colConfig.Remove (iLinha)
    Next
    
    For iLinha = 1 To objGridContas.iLinhasExistentes
        
        Set objConfig = New ClassPlanoContaRefConfig
        
        objConfig.sConta = objPlanoContaRef.sConta
        objConfig.lCodigoModelo = objPlanoContaRef.lCodigoModelo
        objConfig.iSeq = iLinha
    
        'critica o formato da conta
        lErro = CF("Conta_Formata", GridContas.TextMatrix(iLinha, iGrid_ContaInicio_Col), sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 200922
    
        objConfig.sContaInicial = sConta
        
        'critica o formato da conta
        lErro = CF("Conta_Formata", GridContas.TextMatrix(iLinha, iGrid_ContaFinal_Col), sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 200923
    
        objConfig.sContaFinal = sConta
        
        lErro = CF("Ccl_Formata", GridContas.TextMatrix(iLinha, iGrid_CclInicio_Col), sCcl, iCclPreenchida)
        If lErro <> SUCESSO Then gError 200924
        
        objConfig.sCclInicial = sCcl
        
        lErro = CF("Ccl_Formata", GridContas.TextMatrix(iLinha, iGrid_CclFinal_Col), sCcl, iCclPreenchida)
        If lErro <> SUCESSO Then gError 200925
        
        objConfig.sCclFinal = sCcl
        
        objConfig.iSubtrai = StrParaInt(GridContas.TextMatrix(iLinha, iGrid_Subtrai_Col))
        
        If Len(Trim(objConfig.sContaInicial)) > 0 And Len(Trim(objConfig.sContaFinal)) > 0 Then
            If objConfig.sContaInicial > objConfig.sContaFinal Then gError 200926
        End If
    
        If Len(Trim(objConfig.sCclInicial)) > 0 And Len(Trim(objConfig.sCclFinal)) > 0 Then
            If objConfig.sCclInicial > objConfig.sCclFinal Then gError 200927
        End If
        
        If objConfig.sContaInicial <> "" Or objConfig.sContaFinal <> "" Or objConfig.sCclInicial <> "" Or objConfig.sCclFinal <> "" Then
            objPlanoContaRef.colConfig.Add objConfig
        End If
    Next
    
    '=======> FAZER VALIDAÇÕES EM CIMA DA CONTA
    'SE TEM PAI, TIPO, ETC
    lErro = Conta_Critica(objPlanoContaRef)
    If lErro <> SUCESSO Then gError 200928

    '=======> FAZER VALIDAÇÕES EM CIMA DA CONTA
    'VERIFICAR SE ESSA CONTA JÁ EXISTE OU É NOVA
    Call Verifica_Existencia_Conta(objPlanoContaRef.sConta, bContaNova, iIndiceConta)
    
    Mover_Conta_Memoria = SUCESSO
    
    Exit Function
    
Erro_Mover_Conta_Memoria:

    Mover_Conta_Memoria = gErr
    
    Select Case gErr
    
        Case 200921 To 200925, 200928
        
        Case 200926
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INICIAL_MAIOR", gErr)
        
        Case 200927
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200929)
        
    End Select

    Exit Function
        
End Function

Private Function Conta_Critica(ByVal objPlanoContaRef As ClassPlanoContaRef) As Long

'Dim lErro As Long
'Dim sContaPai As String
'Dim iIndicePai As Integer
'Dim bContaNova As Boolean
'Dim objPlanoContaRefPai As ClassPlanoContaRef
'
'On Error GoTo Erro_Conta_Critica
'
'    lErro = Retorna_ContaRef_Pai(objPlanoContaRef.sConta, sContaPai)
'    If lErro <> SUCESSO Then gError 200929
'
'    'Tem Pai
'    If Len(Trim(sContaPai)) > 0 Then
'
'        Call Verifica_Existencia_Conta(sContaPai, bContaNova, iIndicePai)
'
'        'Pai não existe
'        If bContaNova Then gError 200930
'
'        Set objPlanoContaRefPai = gobjPlanoContaRefModelo.colContas(sContaPai)
'
'        'verifica se a conta pai é sintetica. Se for analitica é erro.
'        If objPlanoContaRefPai.iTipo = CONTA_ANALITICA Then gError 200931
'
'    End If
'
'    '===> Poderia verificar se é analítica e tem filhos
'
'    '===> Poderia verificar se a faixa de conta e ccl do config não inclui conta\ccl que
'    'já esteja relacionado a outro plano de conta referencial
'
'    Conta_Critica = SUCESSO
'
'    Exit Function
'
'Erro_Conta_Critica:
'
'    Conta_Critica = gErr
'
'    Select Case gErr
'
'        Case 200929
'
'        Case 200930
'            Call Rotina_Erro(vbOKOnly, "ERRO_CONTAPAI_INEXISTENTE", gErr)
'
'        Case 200931
'            Call Rotina_Erro(vbOKOnly, "ERRO_CONTAPAI_ANALITICA", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200932)
'
'    End Select
'
'    Exit Function
'
End Function

Private Function Traz_Conta_Tela(ByVal objPlanoContaRef As ClassPlanoContaRef) As Long

Dim lErro As Long
Dim sContaEnxuta As String
Dim iIndice As Integer
Dim objConfig As ClassPlanoContaRefConfig
Dim iLinhas As Integer
Dim sContaMascarada As String
Dim sCclMascarada As String

On Error GoTo Erro_Traz_Conta_Tela
    
    'Call Combo_Seleciona_ItemData(TipoConta, objPlanoContaRef.iTipoImp)
    
    TipoConta.Caption = CStr(objPlanoContaRef.iTipoImp) & "-" & IIf(objPlanoContaRef.iTipoImp = CONTA_ANALITICA, "ANALÍTICA", "SINTÉTICA")
    ContaImp.Text = objPlanoContaRef.sContaImp
    
    'lErro = Retorna_ContaRef_Enxuto(objPlanoContaRef.sConta, sContaEnxuta)
    'If lErro <> SUCESSO Then gError 200933
    
    'Conta.PromptInclude = False
    Conta.Text = objPlanoContaRef.sConta 'sContaEnxuta
    'Conta.PromptInclude = True
    
    'Call Combo_Seleciona_ItemData(TipoConta, objPlanoContaRef.iTipo)
    
    DescConta.Text = objPlanoContaRef.sDescricao
    
    
    Call Grid_Limpa(objGridContas)
    
    'Impede definir a conta pois ela é o somatório das filhas
    If objPlanoContaRef.iTipoImp = CONTA_ANALITICA Then
        GrupoConfig.Enabled = True
    Else
        GrupoConfig.Enabled = False
    End If
    
    If objPlanoContaRef.dtValidadeDe <> DATA_NULA Then
        ValidadeDe.Caption = Format(objPlanoContaRef.dtValidadeDe, "dd/mm/yyyy")
    Else
        ValidadeDe.Caption = ""
    End If
    
    If objPlanoContaRef.dtValidadeAte <> DATA_NULA Then
        ValidadeAte.Caption = Format(objPlanoContaRef.dtValidadeAte, "dd/mm/yyyy")
    Else
        ValidadeAte.Caption = ""
    End If
    
    iLinhas = 0
    
    For Each objConfig In objPlanoContaRef.colConfig
    
        iLinhas = iLinhas + 1
    
        If Len(objConfig.sContaInicial) > 0 Then

            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
    
            lErro = Mascara_RetornaContaEnxuta(objConfig.sContaInicial, sContaMascarada)
            If lErro <> SUCESSO Then gError 200934
    
            ContaInicio.PromptInclude = False
            ContaInicio.Text = sContaMascarada
            ContaInicio.PromptInclude = True
            
            GridContas.TextMatrix(iLinhas, iGrid_ContaInicio_Col) = ContaInicio.Text

        End If

        If Len(objConfig.sContaFinal) > 0 Then

            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
    
            lErro = Mascara_RetornaContaEnxuta(objConfig.sContaFinal, sContaMascarada)
            If lErro <> SUCESSO Then gError 200935
    
            ContaFim.PromptInclude = False
            ContaFim.Text = sContaMascarada
            ContaFim.PromptInclude = True

            GridContas.TextMatrix(iLinhas, iGrid_ContaFinal_Col) = ContaFim.Text

        End If

        If Len(objConfig.sCclInicial) > 0 Then

            'mascara o ccl
            sCclMascarada = String(STRING_CCL, 0)
    
            lErro = Mascara_RetornaCclEnxuta(objConfig.sCclInicial, sCclMascarada)
            If lErro <> SUCESSO Then gError 200936
    
            CclInicio.PromptInclude = False
            CclInicio.Text = sCclMascarada
            CclInicio.PromptInclude = True

            GridContas.TextMatrix(iLinhas, iGrid_CclInicio_Col) = CclInicio.Text

        End If

        If Len(objConfig.sCclFinal) > 0 Then

            'mascara o ccl
            sCclMascarada = String(STRING_CCL, 0)
    
            lErro = Mascara_RetornaCclEnxuta(objConfig.sCclFinal, sCclMascarada)
            If lErro <> SUCESSO Then gError 200937
    
            CclFim.PromptInclude = False
            CclFim.Text = sCclMascarada
            CclFim.PromptInclude = True

            GridContas.TextMatrix(iLinhas, iGrid_CclFinal_Col) = CclFim.Text

        End If
        
        GridContas.TextMatrix(iLinhas, iGrid_Subtrai_Col) = CStr(objConfig.iSubtrai)
    
    Next
    
    objGridContas.iLinhasExistentes = iLinhas
    
    Call Grid_Refresh_Checkbox(objGridContas)

    Traz_Conta_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Conta_Tela:

    Traz_Conta_Tela = gErr
    
    Select Case gErr
    
        Case 200933 To 200937
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200938)
        
    End Select

    Exit Function
        
End Function

Private Function Traz_Modelo_Tela(ByVal objPlanoContaRefModelo As ClassPlanoContaRefModelo) As Long

Dim lErro As Long
Dim objTipo As New ClassTiposPlanoContaRef

On Error GoTo Erro_Traz_Modelo_Tela

    Call Limpa_Tela_PlanoConta
    
    'Lê o PlanoContaRefModelo que está sendo Passado
    lErro = CF("PlanoContaRefModelo_Le", objPlanoContaRefModelo)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200939

    If lErro = SUCESSO Then
    
        'CodigoModelo.PromptInclude = False
        CodigoModelo.Text = CStr(objPlanoContaRefModelo.lCodigo)
        'CodigoModelo.PromptInclude = True

        'AnoVigencia.PromptInclude = False
        AnoVigencia.Text = CStr(objPlanoContaRefModelo.iAnoVigencia)
        'AnoVigencia.PromptInclude = True
        
        DescricaoModelo.Text = objPlanoContaRefModelo.sDescricao
        
        If objPlanoContaRefModelo.iOficial = MARCADO Then
            Oficial.Value = vbChecked
        Else
            Oficial.Value = vbUnchecked
        End If
        
        If objPlanoContaRefModelo.iTipo = 0 Then
            Tipo.Text = "0-DESCONTINUADO"
        Else
        
            objTipo.iTipo = objPlanoContaRefModelo.iTipo
            objTipo.iAnoBase = objPlanoContaRefModelo.iAnoVigencia
            
            lErro = CF("TiposPlanoContaRef_Le", objTipo)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200939
    
            Tipo.Text = CStr(objPlanoContaRefModelo.iTipo) & "-" & objTipo.sDescricao
            
        End If

        lErro = Carrega_Arvore(objPlanoContaRefModelo)
        If lErro <> SUCESSO Then gError 200940
        
    End If

    Set gobjPlanoContaRefModelo = objPlanoContaRefModelo

    Traz_Modelo_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Modelo_Tela:

    Traz_Modelo_Tela = gErr
    
    Select Case gErr
    
        Case 200939, 200940
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200941)
        
    End Select

    Exit Function
        
End Function

Private Sub BotaoLimpar_Click()

Dim iLote As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 200942

    Call Limpa_Tela_PlanoConta
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr
    
        Case 200942
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200943)
        
    End Select
        
    Exit Sub
    
End Sub

'Private Sub Conta_Change()
'        iAlterado = REGISTRO_ALTERADO
'End Sub

'Private Sub Conta_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim sContaFormatada As String
'Dim iContaPreenchida As Integer
'Dim objConta As New ClassPlanoContaRef
'Dim bContaNova As Boolean
'
'On Error GoTo Erro_Conta_Validate
'
'    If Len(Conta.ClipText) > 0 Then
'
'        'critica o formato da conta
'        lErro = ContaRef_Formata(Conta.Text, sContaFormatada, iContaPreenchida)
'        If lErro <> SUCESSO Then gError 200943
'
'        If iContaPreenchida = CONTA_PREENCHIDA Then
'
'            lErro = Mover_Conta_Memoria(objConta, bContaNova)
'            If lErro <> SUCESSO Then gError 200944
'
'            lErro = Conta_Critica(objConta)
'            If lErro <> SUCESSO Then gError 200945
'
'        Else
'            sContaFormatada = ""
'        End If
'
'    End If
'
'    Exit Sub
'
'Erro_Conta_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 200943 To 200945
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200946)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Private Sub DescConta_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub

Public Sub Form_Load()

Dim objOrigem As ClassOrigemContab
Dim iIndice As Integer
Dim iLote As Integer
Dim lErro As Long
Dim iIndice1 As Integer
Dim iHabilitaSaldoInicial As Integer

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    Set objEventoConta = New AdmEvento
    Set objEventoCcl = New AdmEvento
       
    Set gobjPlanoContaRefModelo = New ClassPlanoContaRefModelo
    
    ''inicializa a mascara de conta
    'lErro = Inicializa_Mascara_ContaRef(Conta)
    'If lErro <> SUCESSO Then gError 200947
    
    ''inicializar os tipos de conta
    'For iIndice = 1 To gobjColTipoConta.Count
    '    TipoConta.AddItem gobjColTipoConta.Item(iIndice).sDescricao
    '    TipoConta.ItemData(TipoConta.NewIndex) = gobjColTipoConta.Item(iIndice).iTipo
    'Next
        
    ''selecionar o tipo de conta atual
    'For iIndice = 0 To TipoConta.ListCount - 1
    '    If gobjColTipoConta.TipoConta(TipoConta.List(iIndice)) = giTipoConta Then
    '        TipoConta.ListIndex = iIndice
    '        Exit For
    '    End If
    'Next

   'inicializa a mascara de conta
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaInicio)
    If lErro <> SUCESSO Then gError 200948

    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaFim)
    If lErro <> SUCESSO Then gError 200949
               
    'inicializa a mascara de ccl
    lErro = CF("Inicializa_Mascara_Ccl_MaskEd", CclInicio)
    If lErro <> SUCESSO Then gError 200950

    lErro = CF("Inicializa_Mascara_Ccl_MaskEd", CclFim)
    If lErro <> SUCESSO Then gError 200951
    
    Set objGridContas = New AdmGrid
    
    lErro = Inicializa_Grid_Contas(objGridContas)
    If lErro <> SUCESSO Then gError 200952
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
        
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 200947 To 200952
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200953)
        
    End Select
    
    iAlterado = 0
    
    Exit Sub
        
End Sub

Function Trata_Parametros(Optional objPlanoContaRefModelo As ClassPlanoContaRefModelo) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há uma conta selecionada, exibir seus dados
    If Not (objPlanoContaRefModelo Is Nothing) Then
        
        lErro = Traz_Modelo_Tela(objPlanoContaRefModelo)
        If lErro <> SUCESSO Then gError 200954
               
    End If

    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 200954
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200955)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)
           
    Set objEventoCodigo = Nothing
    Set objEventoConta = Nothing
    Set objEventoCcl = Nothing
    
    Set gobjPlanoContaRefModelo = Nothing
    
End Sub

Private Sub LabelCodigo_Click()

Dim objModelo As New ClassPlanoContaRefModelo
Dim colSelecao As New Collection

    objModelo.lCodigo = StrParaLong(CodigoModelo.Text)
    
    Call Chama_Tela("PlanoContaRefModeloLista", colSelecao, objModelo, objEventoCodigo)
    
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objPlanoContaRefModelo As ClassPlanoContaRefModelo
Dim lErro As Long

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objPlanoContaRefModelo = obj1
    
    lErro = Traz_Modelo_Tela(objPlanoContaRefModelo)
    If lErro <> SUCESSO Then gError 200956
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoCodigo_evSelecao:

    Select Case gErr
    
        Case 200956
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200957)
        
    End Select

    Exit Sub
    
End Sub

'Private Sub TipoConta_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_TipoConta_Click
'
'    iAlterado = REGISTRO_ALTERADO
'
'    Exit Sub
'
'Erro_TipoConta_Click:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200958)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub TipoConta_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_TipoConta_Validate
'
'
'    Exit Sub
'
'Erro_TipoConta_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200959)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Function Limpa_Tela_PlanoConta() As Long

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridContas)
    
    TvwContas.Nodes.Clear
    
    Set gobjPlanoContaRefModelo = New ClassPlanoContaRefModelo
        
    Limpa_Tela_PlanoConta = SUCESSO

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "PlanoContaRefModelo"
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", StrParaLong(CodigoModelo.Text), 0, "Codigo"
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200960)

    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objPlanoContaRefModelo As New ClassPlanoContaRefModelo

On Error GoTo Erro_Tela_Preenche

    objPlanoContaRefModelo.lCodigo = colCampoValor.Item("Codigo").vValor

    If objPlanoContaRefModelo.lCodigo <> 0 Then
    
        lErro = Traz_Modelo_Tela(objPlanoContaRefModelo)
        If lErro <> SUCESSO Then gError 200961

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 200961

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200962)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub
'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PLANO_CONTAS
    Set Form_Load_Ocx = Me
    Caption = "Plano de Contas Referencial"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PlanoContaRef"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CodigoModelo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is ContaInicio Then
            Call BotaoContas_Click(0)
        ElseIf Me.ActiveControl Is ContaFim Then
            Call BotaoContas_Click(1)
        ElseIf Me.ActiveControl Is CclInicio Then
            Call BotaoCcl_Click(0)
        ElseIf Me.ActiveControl Is CclFim Then
            Call BotaoCcl_Click(1)
        End If
    
    End If

End Sub

Function Carrega_Arvore(ByVal objPlanoContaRefModelo As ClassPlanoContaRefModelo) As Long
'preenche a treeview Roteiro com a composicao de objRoteirosDeFabricacao
   
Dim objNode As Node
Dim lErro As Long
Dim sContaMascarada As String
Dim iProxChave As Integer
Dim objPlanoContaRef As ClassPlanoContaRef
Dim objPlanoContaRefAux As ClassPlanoContaRef
Dim sConta As String
Dim sContaPai As String

On Error GoTo Erro_Carrega_Arvore

    TvwContas.Nodes.Clear
    
    For Each objPlanoContaRef In objPlanoContaRefModelo.colContas

        ''coloca a conta no formato que é exibida na tela
        'lErro = Mascara_Mascarar_ContaRef(objPlanoContaRef.sConta, sContaMascarada)
        'If lErro <> SUCESSO Then gError 200963
        sContaMascarada = objPlanoContaRef.sContaImp
                
'        If objPlanoContaRef.iTipo = CONTA_ANALITICA Then
'            sConta = "A" & objPlanoContaRef.sConta
'        Else
'            sConta = "S" & objPlanoContaRef.sConta
'        End If
        sConta = "X" & objPlanoContaRef.sConta
        
        sContaPai = ""

        If objPlanoContaRefModelo.iAnoVigencia < 2014 Then
            lErro = Retorna_ContaRef_Pai(objPlanoContaRef.sConta, sContaPai)
            If lErro <> SUCESSO Then gError 200964
        Else
            sContaPai = objPlanoContaRef.sContaPai
        End If

        If Len(Trim(sContaPai)) = 0 Then

            Set objNode = TvwContas.Nodes.Add(, tvwLast, sConta)

        Else
        
            sContaPai = "X" & sContaPai

            Set objNode = TvwContas.Nodes.Add(TvwContas.Nodes.Item(sContaPai), tvwChild, sConta)

       End If
        
        objNode.Text = sContaMascarada & SEPARADOR & objPlanoContaRef.sDescricao
                
        objNode.Tag = objPlanoContaRef.sConta
        
    Next

    Carrega_Arvore = SUCESSO

    Exit Function

Erro_Carrega_Arvore:

    Carrega_Arvore = gErr

    Select Case gErr
    
        Case 200963, 200964
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200965)
            
            Resume Next

    End Select

    Exit Function

End Function

'********************************
'Funções relativas ao GridContas
'********************************

Private Sub GridContas_Click()
Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridContas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContas, iAlterado)
    End If

End Sub

Private Sub GridContas_EnterCell()
    Call Grid_Entrada_Celula(objGridContas, iAlterado)
End Sub

Private Sub GridContas_GotFocus()
    Call Grid_Recebe_Foco(objGridContas)
End Sub

Private Sub GridContas_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridContas)
End Sub

Private Sub GridContas_KeyPress(KeyAscii As Integer)
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridContas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContas, iAlterado)
    End If

End Sub

Private Sub GridContas_LeaveCell()
    Call Saida_Celula(objGridContas)
End Sub

Private Sub GridContas_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridContas)
End Sub

Private Sub GridContas_RowColChange()
    Call Grid_RowColChange(objGridContas)
End Sub

Private Sub GridContas_Scroll()
    Call Grid_Scroll(objGridContas)
End Sub

Private Function Inicializa_Grid_Contas(objGridInt As AdmGrid) As Long
'Inicializa o grid de contas

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Contas

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Conta Início")
    objGridInt.colColuna.Add ("Conta Fim")
    objGridInt.colColuna.Add ("Ccl Início")
    objGridInt.colColuna.Add ("Ccl Fim")
    objGridInt.colColuna.Add ("-")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ContaInicio.Name)
    objGridInt.colCampo.Add (ContaFim.Name)
    objGridInt.colCampo.Add (CclInicio.Name)
    objGridInt.colCampo.Add (CclFim.Name)
    objGridInt.colCampo.Add (Subtrai.Name)

    'Colunas do Grid
    iGrid_ContaInicio_Col = 1
    iGrid_ContaFinal_Col = 2
    iGrid_CclInicio_Col = 3
    iGrid_CclFinal_Col = 4
    iGrid_Subtrai_Col = 5

    'Grid do GridInterno
    objGridInt.objGrid = GridContas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_CONTAS_DRE + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridContas.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Contas = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Contas:

    Inicializa_Grid_Contas = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200966)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Contas(objGridInt As AdmGrid) As Long
''Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Contas

    If objGridInt.objGrid Is GridContas Then

        Select Case GridContas.Col

            Case iGrid_ContaInicio_Col
                lErro = Saida_Celula_ContaInicio(objGridInt)
                If lErro <> SUCESSO Then gError 200967

            Case iGrid_ContaFinal_Col
                lErro = Saida_Celula_ContaFim(objGridInt)
                If lErro <> SUCESSO Then gError 200968

            Case iGrid_CclInicio_Col
                lErro = Saida_Celula_CclInicio(objGridInt)
                If lErro <> SUCESSO Then gError 200969

            Case iGrid_CclFinal_Col
                lErro = Saida_Celula_CclFim(objGridInt)
                If lErro <> SUCESSO Then gError 200970

        End Select

    End If

    Saida_Celula_Contas = SUCESSO

    Exit Function

Erro_Saida_Celula_Contas:

    Saida_Celula_Contas = gErr

    Select Case gErr

        Case 200967 To 200970

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200971)

    End Select

    Exit Function

End Function

Private Sub ContaInicio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaInicio_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContas)
End Sub

Private Sub ContaInicio_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContas)
End Sub

Private Sub ContaInicio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContas.objControle = ContaInicio
    lErro = Grid_Campo_Libera_Foco(objGridContas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ContaFim_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaFim_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContas)
End Sub

Private Sub ContaFim_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContas)
End Sub

Private Sub ContaFim_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContas.objControle = ContaFim
    lErro = Grid_Campo_Libera_Foco(objGridContas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Function Saida_Celula_ContaInicio(objGridInt As AdmGrid) As Long
'faz a critica da celula GeraOP do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objPlanoConta As New ClassPlanoConta
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_ContaInicio

    Set objGridInt.objControle = ContaInicio

    If Len(Trim(ContaInicio.ClipText)) > 0 Then

        sContaFormatada = String(STRING_CONTA, 0)
        
        lErro = CF("Conta_Formata", ContaInicio.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then gError 200972
    
        If iContaPreenchida = CONTA_PREENCHIDA Then
    
            'verifica se a Conta Final existe
            lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
            If lErro <> SUCESSO And lErro <> 6030 Then gError 200973

            If lErro = 6030 Then gError 200974

        End If

        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridContas.Row - GridContas.FixedRows) = objGridContas.iLinhasExistentes Then
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 200975

    Saida_Celula_ContaInicio = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaInicio:

    Saida_Celula_ContaInicio = gErr

    Select Case gErr

        Case 200972, 200973, 200975
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 200974
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr, ContaInicio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200976)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaFim(objGridInt As AdmGrid) As Long
'faz a critica da celula GeraOP do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Saida_Celula_ContaFim

    Set objGridInt.objControle = ContaFim

    If Len(Trim(ContaFim.ClipText)) > 0 Then

        sContaFormatada = String(STRING_CONTA, 0)
        
        lErro = CF("Conta_Formata", ContaFim.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then gError 200977
    
        If iContaPreenchida = CONTA_PREENCHIDA Then
    
            'verifica se a Conta Final existe
            lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
            If lErro <> SUCESSO And lErro <> 6030 Then gError 200978

            If lErro = 6030 Then gError 200979

        End If

        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridContas.Row - GridContas.FixedRows) = objGridContas.iLinhasExistentes Then
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 200980

    Saida_Celula_ContaFim = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaFim:

    Saida_Celula_ContaFim = gErr

    Select Case gErr

        Case 200977, 200978, 200980
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 200979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr, ContaFim.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200981)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub CclInicio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CclInicio_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContas)
End Sub

Private Sub CclInicio_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContas)
End Sub

Private Sub CclInicio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContas.objControle = CclInicio
    lErro = Grid_Campo_Libera_Foco(objGridContas)
    If lErro <> SUCESSO Then Cancel = True
    
    
End Sub

Private Sub CclFim_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CclFim_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContas)
End Sub

Private Sub CclFim_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContas)
End Sub

Private Sub CclFim_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContas.objControle = CclFim
    lErro = Grid_Campo_Libera_Foco(objGridContas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Function Saida_Celula_CclInicio(objGridInt As AdmGrid) As Long
'faz a critica da celula GeraOP do grid que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_CclInicio

    Set objGridInt.objControle = CclInicio

    If Len(Trim(CclInicio.ClipText)) > 0 Then

        'Retorna Ccl formatada como no BD
        lErro = CF("Ccl_Formata", CclInicio.Text, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then gError 200982
    
        objCcl.sCcl = sCclFormatada
    
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 200983
    
        If lErro = 5599 Then gError 200984
        
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridContas.Row - GridContas.FixedRows) = objGridContas.iLinhasExistentes Then
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 200985

    Saida_Celula_CclInicio = SUCESSO

    Exit Function

Erro_Saida_Celula_CclInicio:

    Saida_Celula_CclInicio = gErr

    Select Case gErr

        Case 200982, 200983, 200985
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 200984
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CclInicio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200986)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CclFim(objGridInt As AdmGrid) As Long
'faz a critica da celula GeraOP do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_CclFim

    Set objGridInt.objControle = CclFim

    If Len(Trim(CclFim.ClipText)) > 0 Then

        'Retorna Ccl formatada como no BD
        lErro = CF("Ccl_Formata", CclFim.Text, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then gError 200987
    
        objCcl.sCcl = sCclFormatada
    
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 200988
    
        If lErro = 5599 Then gError 200989
        
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridContas.Row - GridContas.FixedRows) = objGridContas.iLinhasExistentes Then
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 200990

    Saida_Celula_CclFim = SUCESSO

    Exit Function

Erro_Saida_Celula_CclFim:

    Saida_Celula_CclFim = gErr

    Select Case gErr

        Case 200987, 200988, 200990
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 200989
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CclFim.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200991)

    End Select

    Exit Function

End Function

Private Sub SubTrai_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SubTrai_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContas)
End Sub

Private Sub SubTrai_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContas)
End Sub

Private Sub SubTrai_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContas.objControle = Subtrai
    lErro = Grid_Campo_Libera_Foco(objGridContas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub
'********************************
' fim do tratamento do GridContas
'********************************

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridDescontos
            Case GridContas.Name

                lErro = Saida_Celula_Contas(objGridInt)
                If lErro <> SUCESSO Then gError 200992

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 200993

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 200992

        Case 200993
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200994)

    End Select

    Exit Function

End Function

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long
Dim objPlanoContaRef As ClassPlanoContaRef

On Error GoTo Erro_TvwContas_NodeClick
    
    Set objPlanoContaRef = gobjPlanoContaRefModelo.colContas(Node.Tag)
            
    lErro = Traz_Conta_Tela(objPlanoContaRef)
    If lErro <> SUCESSO Then gError 200995
    
    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case gErr
    
        Case 200995
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200996)
            
    End Select
        
    Exit Sub
    
End Sub

Public Sub BotaoContas_Click(iIndice As Integer)

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim sContaOrigem As String
Dim iContaPreenchida As Integer
Dim lErro As Long
Dim objControle As Object

On Error GoTo Erro_BotaoContas_Click

    If GridContas.Row = 0 Then gError 200997

    If iIndice = 0 Then
        Set objControle = ContaInicio
    Else
        Set objControle = ContaFim
    End If
    giConta = iIndice

    If Len(Trim(objControle.ClipText)) > 0 Then
    
        lErro = CF("Conta_Formata", objControle.Text, sContaOrigem, iContaPreenchida)
        If lErro <> SUCESSO Then gError 200998

        If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sContaOrigem
    Else
        objPlanoConta.sConta = ""
    End If
           
    'Chama a tela que lista os vendedores
    Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoConta)

    Exit Sub
    
Erro_BotaoContas_Click:

    Select Case gErr
        
        Case 200997
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 200998
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200999)
            
    End Select

    Exit Sub
    
End Sub

Private Sub objEventoConta_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sConta As String
Dim sContaEnxuta As String
Dim objControle As Object

On Error GoTo Erro_objEventoConta_evSelecao
    
    Set objPlanoConta = obj1
    
    sConta = objPlanoConta.sConta
    
    sContaEnxuta = String(STRING_CONTA, 0)

    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then gError 202000

    If giConta = 0 Then
        Set objControle = ContaInicio
    Else
        Set objControle = ContaFim
    End If
    
    objControle.PromptInclude = False
    objControle.Text = sContaEnxuta
    objControle.PromptInclude = True
    
    If Not (Me.ActiveControl Is objControle) Then
        If giConta = 0 Then
            GridContas.TextMatrix(GridContas.Row, iGrid_ContaInicio_Col) = objControle.Text
        Else
            GridContas.TextMatrix(GridContas.Row, iGrid_ContaFinal_Col) = objControle.Text
        End If
        
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridContas.Row - GridContas.FixedRows) = objGridContas.iLinhasExistentes Then
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
        End If
    End If

    Me.Show
    
    Exit Sub
    
Erro_objEventoConta_evSelecao:

    Select Case gErr

        Case 202000

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202001)
        
    End Select

    Exit Sub

End Sub

Public Sub BotaoCcl_Click(iIndice As Integer)

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection
Dim sCclOrigem As String
Dim iCclPreenchida As Integer
Dim lErro As Long
Dim objControle As Object

On Error GoTo Erro_BotaoCcl_Click

    If GridContas.Row = 0 Then gError 202002

    If iIndice = 0 Then
        Set objControle = CclInicio
    Else
        Set objControle = CclFim
    End If
    giCcl = iIndice
    
    If Len(Trim(objControle.ClipText)) > 0 Then
    
        lErro = CF("Ccl_Formata", objControle.Text, sCclOrigem, iCclPreenchida)
        If lErro <> SUCESSO Then gError 202003

        If iCclPreenchida = CCL_PREENCHIDA Then objCcl.sCcl = sCclOrigem
    Else
        objCcl.sCcl = ""
    End If

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)
    
    Exit Sub
    
Erro_BotaoCcl_Click:

    Select Case gErr
        
        Case 202002
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 202003
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202004)
            
    End Select

    Exit Sub

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objCcl As ClassCcl
Dim sCclEnxuta As String
Dim objControle As Object

On Error GoTo Erro_objEventoCcl_evSelecao
    
    Set objCcl = obj1

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclEnxuta)
    If lErro <> SUCESSO Then gError 202005

    If giCcl = 0 Then
        Set objControle = CclInicio
    Else
        Set objControle = CclFim
    End If
    
    objControle.PromptInclude = False
    objControle.Text = sCclEnxuta
    objControle.PromptInclude = True
    
    If Not (Me.ActiveControl Is objControle) Then
        If giCcl = 0 Then
            GridContas.TextMatrix(GridContas.Row, iGrid_CclInicio_Col) = objControle.Text
        Else
            GridContas.TextMatrix(GridContas.Row, iGrid_CclFinal_Col) = objControle.Text
        End If
    
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridContas.Row - GridContas.FixedRows) = objGridContas.iLinhasExistentes Then
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
        End If
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 202005

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202006)
        
    End Select

    Exit Sub

End Sub

Sub Exclui_Arvore_Conta(objPlanoContaRef As ClassPlanoContaRef)

Dim objNode As Node
Dim bContaNova As Boolean
Dim iIndiceConta As Integer
    
    Call Verifica_Existencia_Conta(objPlanoContaRef.sConta, bContaNova, iIndiceConta)
    
    If Not bContaNova Then
        TvwContas.Nodes.Remove (iIndiceConta)
    End If
    
End Sub

Sub Verifica_Existencia_Conta(ByVal sConta As String, bContaNova As Boolean, iIndiceConta As Integer)

Dim objNode As Node
    
    bContaNova = True
    iIndiceConta = 0
    
    For Each objNode In TvwContas.Nodes
        If objNode.Tag = sConta Then
            bContaNova = False
            iIndiceConta = objNode.Index
            Exit For
        End If
    Next
    
End Sub

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click
    
    lErro = CF("Config_ObterAutomatico", "CTBConfig", "NUM_PROX_COD_MODELO_CONTA_REF", "PlanoContaRefModelo", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 202007
    
    'CodigoModelo.PromptInclude = False
    CodigoModelo.Text = CStr(lCodigo)
    'CodigoModelo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 202007

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202008)
    
    End Select

    Exit Sub
    
End Sub

Public Sub Limpa_Conta()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Conta
    
    'Conta.PromptInclude = False
    Conta.Text = ""
    'Conta.PromptInclude = True
    
    DescConta.Text = ""
    
    TipoConta.Caption = ""
    ValidadeDe.Caption = ""
    ValidadeAte.Caption = ""

    'selecionar o tipo de conta atual
    'For iIndice = 0 To TipoConta.ListCount - 1
    '    If gobjColTipoConta.TipoConta(TipoConta.List(iIndice)) = giTipoConta Then
    '        TipoConta.ListIndex = iIndice
    '        Exit For
    '    End If
    'Next
    
    Call Grid_Limpa(objGridContas)
    
    Exit Sub

Erro_Limpa_Conta:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202009)
    
    End Select

    Exit Sub
    
End Sub
