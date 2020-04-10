VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpDistMPMaqOcx 
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   KeyPreview      =   -1  'True
   ScaleHeight     =   4635
   ScaleMode       =   0  'User
   ScaleWidth      =   6804.255
   Begin VB.Frame FrameMaquinas 
      Caption         =   "Máquinas"
      Height          =   825
      Left            =   90
      TabIndex        =   20
      Top             =   750
      Width           =   6825
      Begin VB.TextBox MaquinaFinal 
         Height          =   300
         Left            =   3810
         MaxLength       =   26
         TabIndex        =   2
         Top             =   300
         Width           =   2895
      End
      Begin VB.TextBox MaquinaInicial 
         Height          =   300
         Left            =   450
         MaxLength       =   26
         TabIndex        =   1
         Top             =   300
         Width           =   2895
      End
      Begin VB.Label LabelMaquinaFinal 
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
         Left            =   3450
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   360
         Width           =   360
      End
      Begin VB.Label LabelMaquinaInicial 
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4770
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpDistMPMaquinasOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpDistMPMaquinasOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpDistMPMaquinasOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpDistMPMaquinasOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpDistMPMaquinasOcx.ctx":0994
      Left            =   735
      List            =   "RelOpDistMPMaquinasOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2220
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3180
      Picture         =   "RelOpDistMPMaquinasOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   1395
   End
   Begin VB.Frame FrameProduto 
      Caption         =   "Produtos"
      Height          =   1215
      Left            =   90
      TabIndex        =   17
      Top             =   3300
      Width           =   6825
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   645
         TabIndex        =   10
         Top             =   735
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   645
         TabIndex        =   9
         Top             =   330
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelProdutoAte 
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
         Left            =   270
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   795
         Width           =   360
      End
      Begin VB.Label LabelProdutoDe 
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
         Left            =   315
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   31
         Top             =   390
         Width           =   315
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2160
         TabIndex        =   30
         Top             =   735
         Width           =   4455
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2160
         TabIndex        =   29
         Top             =   330
         Width           =   4455
      End
   End
   Begin VB.Frame FrameOrdemProducao 
      Caption         =   "Ordem de Produção"
      Height          =   1605
      Left            =   90
      TabIndex        =   16
      Top             =   1650
      Width           =   6825
      Begin VB.Frame FrameOPCodigo 
         Caption         =   "Código"
         Height          =   1185
         Left            =   180
         TabIndex        =   26
         Top             =   270
         Width           =   3165
         Begin VB.TextBox OpCodigoInicial 
            Height          =   300
            Left            =   840
            TabIndex        =   3
            Top             =   300
            Width           =   1695
         End
         Begin VB.TextBox OpCodigoFinal 
            Height          =   300
            Left            =   840
            TabIndex        =   4
            Top             =   690
            Width           =   1695
         End
         Begin VB.Label LabelOpFinal 
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
            Left            =   450
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   750
            Width           =   360
         End
         Begin VB.Label LabelOpInicial 
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
            Left            =   480
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   27
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.Frame FrameOPData 
         Caption         =   "Data"
         Height          =   1185
         Left            =   3510
         TabIndex        =   23
         Top             =   270
         Width           =   3165
         Begin MSComCtl2.UpDown UpDownOPDataInicial 
            Height          =   315
            Left            =   2235
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   300
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataInicial 
            Height          =   300
            Left            =   1110
            TabIndex        =   5
            Top             =   300
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownOPDataFinal 
            Height          =   315
            Left            =   2235
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   690
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataFinal 
            Height          =   300
            Left            =   1110
            TabIndex        =   7
            Top             =   690
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelOPDataInicial 
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
            Height          =   240
            Left            =   720
            TabIndex        =   25
            Top             =   337
            Width           =   345
         End
         Begin VB.Label LabelOPDataFinal 
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
            Left            =   690
            TabIndex        =   24
            Top             =   750
            Width           =   360
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   19
      Top             =   270
      Width           =   825
   End
End
Attribute VB_Name = "RelOpDistMPMaqOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Dim giProdInicial As Integer
Dim giOp_Inicial As Integer
Dim giMaquinaInicial As Integer

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoMaquinaDe As AdmEvento
Attribute objEventoMaquinaDe.VB_VarHelpID = -1
Private WithEvents objEventoMaquinaAte As AdmEvento
Attribute objEventoMaquinaAte.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1
Private WithEvents objEventoOpDe As AdmEvento
Attribute objEventoOpDe.VB_VarHelpID = -1
Private WithEvents objEventoOpAte As AdmEvento
Attribute objEventoOpAte.VB_VarHelpID = -1

Public Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 103088

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 103089

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 103089

        Case 103088
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168392)

    End Select

End Function

Private Sub BotaoFechar_Click()
'Sai da Tela

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()
'Faz a Limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 106462
    
    ComboOpcoes.Text = ""
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 106462
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168393)

    End Select


End Sub

Private Sub LabelMaquinaFinal_Click()

Dim lErro As Long
Dim objMaquina As New ClassMaquinas
Dim colSelecao As Collection

On Error GoTo Erro_LabelMaquinaFinal_Click

    giMaquinaInicial = 0

    If Len(Trim(MaquinaFinal.Text)) <> 0 Then

        objMaquina.sNomeReduzido = MaquinaFinal.Text

    End If

    Call Chama_Tela("MaquinasLista", colSelecao, objMaquina, objEventoMaquinaAte)

    Exit Sub

Erro_LabelMaquinaFinal_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168394)

    End Select

End Sub

Private Sub LabelMaquinaInicial_Click()

Dim lErro As Long
Dim objMaquina As New ClassMaquinas
Dim colSelecao As Collection

On Error GoTo Erro_LabelMaquinaInicial_Click

    giMaquinaInicial = 1

    If Len(Trim(MaquinaInicial.Text)) <> 0 Then

        objMaquina.sNomeReduzido = MaquinaInicial.Text

    End If

    Call Chama_Tela("MaquinasLista", colSelecao, objMaquina, objEventoMaquinaDe)

    Exit Sub

Erro_LabelMaquinaInicial_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168395)

    End Select

End Sub

Private Sub MaquinaFinal_GotFocus()

    giMaquinaInicial = 0

End Sub

Private Sub MaquinaInicial_GotFocus()

    giMaquinaInicial = 1

End Sub

Private Sub MaquinaInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_MaquinaInicial_Validate

    If Len(Trim(MaquinaInicial.Text)) > 0 Then
        
        lErro = CF("TP_Maquina_Le2", MaquinaInicial, objMaquina)
        If lErro <> SUCESSO And lErro <> 106451 And lErro <> 106453 Then gError 106458
        
        'Se nao encontrou => Erro
        If lErro = 106451 Or lErro = 106453 Then gError 106459
        
    End If

    Exit Sub

Erro_MaquinaInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 106458
        
        Case 106459
                Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, MaquinaInicial.Text)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168396)

    End Select

End Sub

Private Sub objEventoMaquinaAte_evSelecao(obj1 As Object)
'Evento do Browser

Dim lErro As Long
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_objEventoMaquinaAte_evSelecao

    Set objMaquina = obj1
    
    objMaquina.iFilialEmpresa = giFilialEmpresa
    
    'Tenta Ler a maquina
    lErro = CF("Maquinas_Le", objMaquina)
    If lErro <> SUCESSO And lErro <> 103090 Then gError 106445
    
    'Se nao Encontrou => Erro
    If lErro = 103090 Then gError 106446
    
    'Coloca na Tela o Codigo "-" NomeReduzido
    MaquinaFinal.Text = objMaquina.iCodigo & SEPARADOR & objMaquina.sNomeReduzido
    
    Me.Show
    
    Exit Sub

Erro_objEventoMaquinaAte_evSelecao:

    Select Case gErr
    
        Case 106445
        
        Case 106446
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, objMaquina.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168397)
            
    End Select

End Sub

Private Sub objEventoMaquinaDe_evSelecao(obj1 As Object)
'Evento do Browser

Dim lErro As Long
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_objEventoMaquinaDe_evSelecao

    Set objMaquina = obj1
    
    objMaquina.iFilialEmpresa = giFilialEmpresa
    
    'Tenta Ler a maquina
    lErro = CF("Maquinas_Le", objMaquina)
    If lErro <> SUCESSO And lErro <> 103090 Then gError 106447
    
    'Se nao Encontrou => Erro
    If lErro = 103090 Then gError 106448
    
    'Coloca na Tela o Codigo "-" NomeReduzido
    MaquinaInicial.Text = objMaquina.iCodigo & SEPARADOR & objMaquina.sNomeReduzido
    
    Me.Show
    
    Exit Sub

Erro_objEventoMaquinaDe_evSelecao:

    Select Case gErr
    
        Case 106447
        
        Case 106448
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, objMaquina.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168398)
            
    End Select

End Sub

Private Sub objEventoOpDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOp As New ClassOrdemDeProducao

On Error GoTo Erro_objEventoOpDe_evSelecao

    Set objOp = obj1
    
    objOp.iFilialEmpresa = giFilialEmpresa

    'Tenta ler a OP
    lErro = CF("OrdemDeProducao_Le_SemItens", objOp)
    If lErro <> SUCESSO And lErro <> 34455 Then gError 106458
    
    'Se nao existir => Erro
    If lErro = 34455 Then gError 106459
    
    'Coloca na tela o Código da OP
    OpCodigoInicial.Text = objOp.sCodigo
    
    Me.Show
    
    Exit Sub

Erro_objEventoOpDe_evSelecao:

    Select Case gErr
    
        Case 106458
        
        Case 106459
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE", gErr)
    
       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168399)

    End Select

End Sub

Private Sub objEventoOpAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOp As New ClassOrdemDeProducao

On Error GoTo Erro_objEventoOpAte_evSelecao

    Set objOp = obj1

    objOp.iFilialEmpresa = giFilialEmpresa
    
    'Tenta ler a OP
    lErro = CF("OrdemDeProducao_Le_SemItens", objOp)
    If lErro <> SUCESSO And lErro <> 34455 Then gError 106460
    
    'Se nao existir => Erro
    If lErro = 34455 Then gError 106461
    
    'Coloca na tela o Código da OP
    OpCodigoFinal.Text = objOp.sCodigo
    
    Me.Show
    
    Exit Sub

Erro_objEventoOpAte_evSelecao:

    Select Case gErr
    
        Case 106460
        
        Case 106461
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE", gErr)
    
       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168400)

    End Select

End Sub

Private Function Valida_OrdProd(sCodigoOP As String) As Long

Dim objOp As New ClassOrdemDeProducao
Dim lErro As Long

On Error GoTo Erro_Valida_OrdProd

    objOp.iFilialEmpresa = giFilialEmpresa
    objOp.sCodigo = sCodigoOP

    giOp_Inicial = 1
    giProdInicial = 1

    'busca ordem de produção aberta
    lErro = CF("OrdemDeProducao_Le_SemItens", objOp)
    If lErro <> SUCESSO And lErro <> 34455 Then

        gError 103078

    Else

        'se não existe ordem de produção aberta
        If lErro <> SUCESSO Then

            'busca ordem de produção baixada
            lErro = CF("OPBaixada_Le_SemItens", objOp)
            If lErro <> SUCESSO And lErro <> 34459 Then gError 103079

            If lErro = 34459 Then gError 103080

        End If

    End If

    Valida_OrdProd = SUCESSO

    Exit Function

Erro_Valida_OrdProd:

    Valida_OrdProd = gErr

    Select Case gErr

        Case 103078, 103079

        Case 103080
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168401)

    End Select

End Function

Private Sub LabelOpFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objOp As ClassOrdemDeProducao

On Error GoTo Erro_LabelOpFinal_Click

    giOp_Inicial = 0

    If Len(Trim(OpCodigoFinal.Text)) <> 0 Then

        Set objOp = New ClassOrdemDeProducao
        objOp.sCodigo = OpCodigoFinal.Text

    End If

    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOp, objEventoOpAte)
    
   Exit Sub

Erro_LabelOpFinal_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168402)

    End Select

End Sub

Private Sub LabelOpInicial_Click()

Dim lErro As Long
Dim objOp As ClassOrdemDeProducao
Dim colSelecao As Collection

On Error GoTo Erro_LabelOpInicial_Click

    giOp_Inicial = 1

    If Len(Trim(OpCodigoInicial.Text)) <> 0 Then

        Set objOp = New ClassOrdemDeProducao
        objOp.sCodigo = OpCodigoInicial.Text

    End If

    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOp, objEventoOpDe)

    Exit Sub

Erro_LabelOpInicial_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168403)

    End Select

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 103064

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 103065

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 103066

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 103064, 103066

        Case 103065
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168404)

    End Select

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 103067

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 103068

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 103069

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 103067, 103069

        Case 103068
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168405)

    End Select

End Sub

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoAte_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoFinal.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 103070

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 103070

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168406)

    End Select

End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoDe_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoInicial.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 103071

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 103071

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168407)

    End Select

End Sub

Private Sub DataInicial_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataFinal_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 103072

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 103072

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168408)

    End Select

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 103073

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 103073

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168409)

    End Select

End Sub

Private Sub OpCodigoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpCodigoFinal_Validate

    giOp_Inicial = 0

    If Len(Trim(OpCodigoFinal.Text)) <> 0 Then

        lErro = Valida_OrdProd(OpCodigoFinal.Text)
        If lErro <> SUCESSO Then gError 103082

    End If

    Exit Sub

Erro_OpCodigoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 103082

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168410)

    End Select

End Sub

Private Sub OpCodigoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpCodigoInicial_Validate

    giOp_Inicial = 1

    If Len(Trim(OpCodigoInicial.Text)) <> 0 Then

        lErro = Valida_OrdProd(OpCodigoInicial.Text)
        If lErro <> SUCESSO Then gError 103081

    End If

    Exit Sub

Erro_OpCodigoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 103081

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168411)

    End Select

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoFinal_Validate

    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdFormatado, iProdPreenchido)
    If lErro <> SUCESSO Then gError 108511

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 108512

        'Se nao Encontrou => Erro
        If lErro = 28030 Then gError 108513
        
'*************************
        'se for gerencial => Erro
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 108591
        
        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 108592
        
        'Se não controla estoque => Erro
        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 108593
        
        'Se nao for um produto produzido => Erro
        If objProduto.iCompras = PRODUTO_COMPRAVEL Then gError 108594
'*************************
    
        DescProdFim.Caption = objProduto.sDescricao
        
    End If
    
    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 108512, 108511

        Case 108513
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoFinal.Text)

'*************************
        Case 108591
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, ProdutoFinal.Text)
            
        Case 108592
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, ProdutoFinal.Text)
        
        Case 108593
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_COM_ESTOQUE", gErr, ProdutoFinal.Text)
            
        Case 108594
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, ProdutoFinal.Text)
'*************************

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168412)

    End Select

End Sub

Private Sub ProdutoInicial_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProdutoInicial)
    
End Sub

Private Sub ProdutoFinal_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProdutoFinal)
    
End Sub


Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoInicial_Validate

    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdFormatado, iProdPreenchido)
    If lErro <> SUCESSO Then gError 108511

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 108512

        'Se nao Encontrou => Erro
        If lErro = 28030 Then gError 108513
        
'*************************
        'se for gerencial => Erro
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 108591
        
        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 108592
        
        'Se não controla estoque => Erro
        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 108593
        
        'Se nao for um produto produzido => Erro
        If objProduto.iCompras = PRODUTO_COMPRAVEL Then gError 108594
'*************************
    
        DescProdInic.Caption = objProduto.sDescricao
        
    End If
    
    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 108512, 108511

        Case 108513
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoInicial.Text)

'*************************
        Case 108591
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, ProdutoInicial.Text)
            
        Case 108592
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, ProdutoInicial.Text)
        
        Case 108593
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_COM_ESTOQUE", gErr, ProdutoInicial.Text)
            
        Case 108594
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, ProdutoInicial.Text)
'*************************

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168413)

    End Select

End Sub

Private Sub MaquinaFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_MaquinaFinal_Validate

    If Len(Trim(MaquinaFinal.Text)) > 0 Then
        
        lErro = CF("TP_Maquina_Le2", MaquinaFinal, objMaquina)
        If lErro <> SUCESSO And lErro <> 106451 And lErro <> 106453 Then gError 103103
        
        'Se nao encontrou => Erro
        If lErro = 106451 Or lErro = 106453 Then gError 106457
        
    End If

    Exit Sub

Erro_MaquinaFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 103103
        
        Case 106457
                Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, MaquinaFinal.Text)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168414)

    End Select

End Sub

Private Sub UpDownOPDataFinal_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownOPDataFinal_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 103076

    Exit Sub

Erro_UpDownOPDataFinal_DownClick:

    Select Case gErr

        Case 103076
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168415)

    End Select

End Sub

Private Sub UpDownOPDataFinal_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownOPDataFinal_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 103077

    Exit Sub

Erro_UpDownOPDataFinal_UpClick:

    Select Case gErr

        Case 103077
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168416)

    End Select

End Sub

Private Sub UpDownOPDataInicial_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownOPDataInicial_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 103074

    Exit Sub

Erro_UpDownOPDataInicial_DownClick:

    Select Case gErr

        Case 103074
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168417)

    End Select

End Sub

Private Sub UpDownOPDataInicial_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownOPDataInicial_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 103075

    Exit Sub

Erro_UpDownOPDataInicial_UpClick:

    Select Case gErr

        Case 103075
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168418)

    End Select

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load

    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoOpDe = New AdmEvento
    Set objEventoOpAte = New AdmEvento
    Set objEventoMaquinaDe = New AdmEvento
    Set objEventoMaquinaAte = New AdmEvento

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 103051

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 103052

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 103051, 103052, 103087

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168419)

    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set gobjRelOpcoes = Nothing
    Set gobjRelatorio = Nothing
    Set objEventoMaquinaDe = Nothing
    Set objEventoMaquinaAte = Nothing
    Set objEventoOpDe = Nothing
    Set objEventoOpAte = Nothing

End Sub


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_FALTAS
    Set Form_Load_Ocx = Me
    Caption = "Distribuição de matéria-prima por máquina"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpDistMPMaq"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        ElseIf Me.ActiveControl Is MaquinaFinal Then
            Call LabelMaquinaFinal_Click
        ElseIf Me.ActiveControl Is MaquinaInicial Then
            Call LabelMaquinaInicial_Click
        ElseIf Me.ActiveControl Is OpCodigoInicial Then
            Call LabelOpInicial_Click
        ElseIf Me.ActiveControl Is OpCodigoFinal Then
            Call LabelOpFinal_Click
        End If

    End If

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelMaquinaFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelMaquinaFinal, Source, X, Y)
End Sub

Private Sub LabelMaquinaFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelMaquinaFinal, Button, Shift, X, Y)
End Sub

Private Sub LabelMaquinaInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelMaquinaInicial, Source, X, Y)
End Sub

Private Sub LabelMaquinaInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelMaquinaInicial, Button, Shift, X, Y)
End Sub

Private Sub LabelOpInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOpInicial, Source, X, Y)
End Sub

Private Sub LabelOpInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOpInicial, Button, Shift, X, Y)
End Sub

Private Sub LabelOpFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOpFinal, Source, X, Y)
End Sub

Private Sub LabelOpFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOpFinal, Button, Shift, X, Y)
End Sub

Private Sub LabelOPDataInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOPDataInicial, Source, X, Y)
End Sub

Private Sub LabelOPDataInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOPDataInicial, Button, Shift, X, Y)
End Sub

Private Sub LabelOPDataFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOPDataFinal, Source, X, Y)
End Sub

Private Sub LabelOPDataFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOPDataFinal, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoDe, Source, X, Y)
End Sub

Private Sub LabelProdutoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoAte, Source, X, Y)
End Sub

Private Sub LabelProdutoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoAte, Button, Shift, X, Y)
End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 106470

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 106471

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 106472
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 106473
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 106470
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 106471, 106472, 106473

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168420)

    End Select

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutar As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String, lNumIntRel As Long

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
       
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 106473
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 106474
         
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 106475

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 106476
    
    lErro = objRelOpcoes.IncluirParametro("TCOPINI", OpCodigoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 106477
    
    lErro = objRelOpcoes.IncluirParametro("TCOPFIM", OpCodigoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 106478
    
    lErro = objRelOpcoes.IncluirParametro("NMAQINI", Codigo_Extrai(MaquinaInicial.Text))
    If lErro <> AD_BOOL_TRUE Then gError 106479
    
    lErro = objRelOpcoes.IncluirParametro("NMAQFIM", Codigo_Extrai(MaquinaFinal.Text))
    If lErro <> AD_BOOL_TRUE Then gError 106480
    
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAINI", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAINI", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 106481
    
    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 106481
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 106483

    If bExecutar Then
    
        lErro = CF("ItensOPRel_Prepara2", giFilialEmpresa, lNumIntRel, sProd_I, sProd_F, MaskedParaDate(DataInicial), MaskedParaDate(DataFinal), OpCodigoInicial.Text, OpCodigoFinal.Text)
        If lErro <> SUCESSO Then gError 106483
        
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 106747
    
    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 106473 To 106483, 106747

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168421)

    End Select

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long, bPodeSeguir As Boolean

On Error GoTo Erro_Formata_E_Critica_Parametros

    bPodeSeguir = False
    
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 106465

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 106466

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 106467

        bPodeSeguir = True
        
    End If

   'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If StrParaDate(DataInicial.Text) > StrParaDate(DataFinal.Text) Then gError 106468
    
         bPodeSeguir = True
    
    End If
    
    'op inicial não pode ser maior que a op final
    If Trim(OpCodigoInicial.Text) <> "" And Trim(OpCodigoFinal.Text) <> "" Then
    
        If CStr(OpCodigoInicial.Text) > CStr(OpCodigoFinal.Text) Then gError 106469
    
        bPodeSeguir = True
    
    End If
    
    'maquina inicial não pode ser maior que a maquina final
    If Trim(MaquinaInicial.Text) <> "" And Trim(MaquinaFinal.Text) <> "" Then
    
        If Codigo_Extrai(CStr(MaquinaInicial.Text)) > Codigo_Extrai(CStr(MaquinaFinal.Text)) Then gError 106471
    
    End If
    
    If bPodeSeguir = False Then gError 106746
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
        
        Case 106746
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_FILTRO_OP_DATA_OU_PRODUTO", gErr)
        
        Case 106465
            ProdutoInicial.SetFocus

        Case 106466
            ProdutoFinal.SetFocus

        Case 106467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
             
        Case 106468
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
            
        Case 106469
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_INICIAL_MAIOR", gErr)
            OpCodigoInicial.SetFocus
        
        Case 106471
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168422)

    End Select

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""
    
    If sProd_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = "PRODUTO >= " & Forprint_ConvTexto(sProd_I)

    End If

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PRODUTO <= " & Forprint_ConvTexto(sProd_F)

    End If
        
    If Trim(DataInicial.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DATAOP >= " & Forprint_ConvData(StrParaDate(DataInicial.Text))

    End If
    
    If Trim(DataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DATAOP <= " & Forprint_ConvData(StrParaDate(DataFinal.Text))

    End If
 
    If Trim(MaquinaInicial.Text) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "MAQUINA >= " & Forprint_ConvInt(Codigo_Extrai(MaquinaInicial.Text))

    End If
    
    If Trim(MaquinaFinal.Text) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "MAQUINA <= " & Forprint_ConvInt(Codigo_Extrai(MaquinaFinal.Text))

    End If
 
    If Trim(OpCodigoInicial.Text) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CODIGOOP >= " & Forprint_ConvTexto(OpCodigoInicial.Text)

    End If
    
    If Trim(OpCodigoFinal.Text) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CODIGOOP <= " & Forprint_ConvTexto(OpCodigoFinal.Text)

    End If
 
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168423)

    End Select

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 106485
   
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 106486

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 106487

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 106488

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 106489
    
    'pega Maquina Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NMAQINI", sParam)
    If lErro <> SUCESSO Then gError 106490
    If sParam = "0" Then sParam = ""
    MaquinaInicial.Text = sParam
    Call MaquinaInicial_Validate(bSGECancelDummy)
    
    'pega Maquina Final e exibe
    lErro = objRelOpcoes.ObterParametro("NMAQFIM", sParam)
    If lErro <> SUCESSO Then gError 106491
    If sParam = "0" Then sParam = ""
    MaquinaFinal.Text = sParam
    Call MaquinaFinal_Validate(bSGECancelDummy)
    
    'pega a OP Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCOPINI", sParam)
    If lErro <> SUCESSO Then gError 106492
    OpCodigoInicial.Text = sParam
    Call OpCodigoInicial_Validate(bSGECancelDummy)
    
    'pega a OP Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCOPFIM", sParam)
    OpCodigoFinal.Text = sParam
    If lErro <> SUCESSO Then gError 106493
    Call OpCodigoFinal_Validate(bSGECancelDummy)

    'pega a Data Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINI", sParam)
    If lErro <> SUCESSO Then gError 106494
    Call DateParaMasked(DataInicial, StrParaDate(sParam))
    
    'pega a Data Final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAFIM", sParam)
    If lErro <> SUCESSO Then gError 106495
    Call DateParaMasked(DataFinal, StrParaDate(sParam))
    
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 106485 To 106495

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168424)

    End Select

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 106496

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_REL_OP_DIST_MP_REATOR")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 106497

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
         lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 106498
    
        ComboOpcoes.Text = ""
        DescProdInic.Caption = ""
        DescProdFim.Caption = ""
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 106496
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 106497, 106498

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168425)

    End Select

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 108500

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 108500

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168426)

    End Select

    Exit Sub

End Sub
