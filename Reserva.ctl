VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl Reserva 
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   9510
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2220
      Picture         =   "Reserva.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   285
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7230
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Reserva.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Reserva.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Reserva.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Reserva.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Almoxarifados 
      Height          =   3960
      Left            =   6720
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   1350
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.ListBox Reservas 
      Height          =   3960
      Left            =   6705
      TabIndex        =   11
      Top             =   1350
      Width           =   2670
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datas"
      Height          =   795
      Left            =   135
      TabIndex        =   24
      Top             =   3945
      Width           =   6285
      Begin MSComCtl2.UpDown UpDownValidade 
         Height          =   300
         Left            =   5670
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   330
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataValidade 
         Height          =   300
         Left            =   4620
         TabIndex        =   8
         Top             =   315
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Reserva:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   28
         Top             =   360
         Width           =   780
      End
      Begin VB.Label DataReserva 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1305
         TabIndex        =   27
         Top             =   330
         Width           =   1170
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Validade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3720
         TabIndex        =   26
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Quantidades"
      Height          =   795
      Left            =   135
      TabIndex        =   20
      Top             =   2160
      Width           =   6285
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   315
         Left            =   1290
         TabIndex        =   5
         Top             =   315
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Disponível:"
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
         Left            =   3570
         TabIndex        =   23
         Top             =   360
         Width           =   990
      End
      Begin VB.Label QuantDisponivel 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4605
         TabIndex        =   22
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Reservada:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   225
         TabIndex        =   21
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.CommandButton Padrao 
      Caption         =   "Padrão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3630
      TabIndex        =   4
      Top             =   1710
      Width           =   930
   End
   Begin VB.Frame Frame1 
      Caption         =   "Origem da Reserva"
      Height          =   795
      Left            =   150
      TabIndex        =   17
      Top             =   3045
      Width           =   6285
      Begin VB.ComboBox TipoDocAssoc 
         Height          =   315
         ItemData        =   "Reserva.ctx":0A7E
         Left            =   1305
         List            =   "Reserva.ctx":0A80
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   330
         Width           =   1950
      End
      Begin MSMask.MaskEdBox Documento 
         Height          =   315
         Left            =   4605
         TabIndex        =   7
         Top             =   330
         Visible         =   0   'False
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Origem:"
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
         Left            =   540
         TabIndex        =   19
         Top             =   375
         Width           =   660
      End
      Begin VB.Label DocumentoLabel 
         AutoSize        =   -1  'True
         Caption         =   "Ped. Venda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3525
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   375
         Visible         =   0   'False
         Width           =   1065
      End
   End
   Begin VB.TextBox Responsabilidade 
      Height          =   300
      Left            =   1425
      MaxLength       =   50
      TabIndex        =   9
      Top             =   4920
      Width           =   4995
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   315
      Left            =   1425
      TabIndex        =   2
      Top             =   735
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Almoxarifado 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   1710
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   1410
      TabIndex        =   0
      Top             =   270
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin VB.Label LabelAlmoxarifado 
      AutoSize        =   -1  'True
      Caption         =   "Almoxarifados"
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
      Left            =   6690
      TabIndex        =   37
      Top             =   1110
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label UnidadeMedida 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1395
      TabIndex        =   36
      Top             =   1215
      Width           =   810
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "U.M.:"
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
      Left            =   840
      TabIndex        =   35
      Top             =   1260
      Width           =   480
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2940
      TabIndex        =   34
      Top             =   735
      Width           =   3480
   End
   Begin VB.Label CodigoLabel 
      AutoSize        =   -1  'True
      Caption         =   " Número:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   540
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   33
      Top             =   300
      Width           =   780
   End
   Begin VB.Label ProdutoLabel 
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   600
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   32
      Top             =   765
      Width           =   735
   End
   Begin VB.Label AlmoxarifadoLabel 
      AutoSize        =   -1  'True
      Caption         =   "Almoxarifado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   165
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   31
      Top             =   1755
      Width           =   1155
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Responsável:"
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
      Left            =   180
      TabIndex        =   30
      Top             =   4965
      Width           =   1170
   End
   Begin VB.Label LabelReserva 
      AutoSize        =   -1  'True
      Caption         =   "Reservas"
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
      Left            =   6690
      TabIndex        =   29
      Top             =   1110
      Width           =   1185
   End
End
Attribute VB_Name = "Reserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iCodigoAlterado As Integer
Dim iProdutoAlterado As Integer
Dim iAlmoxarifadoAlterado As Integer
Dim iDocumentoAlterado As Integer

Dim WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Dim WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Dim WithEvents objEventoDocumento As AdmEvento
Attribute objEventoDocumento.VB_VarHelpID = -1
Dim WithEvents objEventoAlmoxarifado As AdmEvento
Attribute objEventoAlmoxarifado.VB_VarHelpID = -1

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera o codigo automatico
    lErro = CF("Reserva_Automatico", lCodigo)
    If lErro <> SUCESSO Then Error 57526

    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57526
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174089)
    
    End Select

    Exit Sub

End Sub

Private Sub Almoxarifado_Change()

    iAlmoxarifadoAlterado = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Almoxarifado_GotFocus()
'Mostra os Almoxarifados no lugar das Reservas

    Almoxarifados.Visible = True
    Reservas.Visible = False
'    Produtos.Visible = False

    LabelAlmoxarifado.Visible = True
    LabelReserva.Visible = False
'    LabelProduto.Visible = False

End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Almoxarifado_Validate

    'Se o almoxarifado foi alterado
    If iAlmoxarifadoAlterado <> 0 Then

        If Len(Trim(Almoxarifado.Text)) = 0 Then
            QuantDisponivel.Caption = ""

        Else
            objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

            lErro = TP_Almoxarifado_Le(Almoxarifado, objAlmoxarifado)
            If lErro <> SUCESSO Then Error 23841

            'Se produto estiver preenchido --> Calcula o disponível
            If Len(Produto.ClipText) > 0 Then

                'Formata produto como no Banco de Dados
                lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then Error 55871

                lErro = CalculaDisponivel(sProdutoFormatado, Almoxarifado.Text, Documento.Text, Codigo.Text)
                If lErro <> SUCESSO Then Error 23959

                If Len(Quantidade.Text) <> 0 Then
                    If CDbl(Quantidade.Text) > CDbl(QuantDisponivel.Caption) Then Error 55507
                End If

            End If

        End If

        iAlmoxarifadoAlterado = 0

    End If

    Exit Sub

Erro_Almoxarifado_Validate:

    Cancel = True

    Select Case Err

        Case 23841, 23959
            Quantidade.Text = ""
            QuantDisponivel.Caption = Formata_Estoque(0)

        Case 55507
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_RESERVADA_MAIOR", Err)

        Case 55871

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174090)

    End Select

    Exit Sub

End Sub

Private Sub AlmoxarifadoLabel_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoLabel_Click
    
    If Len(Trim(Almoxarifado.Text)) <> 0 Then

        objAlmoxarifado.sNomeReduzido = Almoxarifado.Text
        
        'Lê o Almoxarifado
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then Error 55568
    
        'Se não encontrou o Almoxaifado --> erro
        If lErro = 25060 Then Error 55569
    
    End If

    Call Chama_Tela("AlmoxarifadoLista", colSelecao, objAlmoxarifado, objEventoAlmoxarifado)

    Exit Sub

Erro_AlmoxarifadoLabel_Click:

    Select Case Err

        Case 55568

        Case 55569
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", Err, objAlmoxarifado.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174091)

    End Select

    Exit Sub

End Sub

Private Sub objEventoAlmoxarifado_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objAlmoxarifado As ClassAlmoxarifado
Dim bCancel As Boolean

On Error GoTo Erro_objEventoAlmoxarifado_evSelecao

    Set objAlmoxarifado = obj1

    Almoxarifado.Text = objAlmoxarifado.iCodigo
    Call Almoxarifado_Validate(bCancel)

    If bCancel = True Then Almoxarifado.Text = ""

    Me.Show

    Exit Sub

Erro_objEventoAlmoxarifado_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174092)

    End Select

    Exit Sub

End Sub

Private Sub Almoxarifados_DblClick()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Almoxarifados_DblClick

    If Almoxarifados.ListIndex = -1 Then Exit Sub

    If Almoxarifado.Text = Almoxarifados.List(Almoxarifados.ListIndex) Then Exit Sub

    'Se o produto estiver preenchido Calcula o Total Disponível
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Formata produto como no Banco de Dados
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 55872

        lErro = CalculaDisponivel(sProdutoFormatado, Almoxarifados.List(Almoxarifados.ListIndex), Documento.Text, Codigo.Text)
        If lErro <> SUCESSO Then Error 23946

    End If

    Almoxarifado.Text = Almoxarifados.List(Almoxarifados.ListIndex)

    Exit Sub

Erro_Almoxarifados_DblClick:

    Select Case Err

        Case 23946
            Quantidade.Text = ""
            QuantDisponivel.Caption = Formata_Estoque(0)

        Case 55872

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174093)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim lCodigo As Long
Dim objReserva As New ClassReserva
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Error 23998

    objReserva.lCodigo = CLng(Codigo.Text)
    
    'Lê a Reserva
    lErro = CF("Reserva_Le", objReserva)
    If lErro <> SUCESSO And lErro <> 23928 Then Error 23999

    'Se não encontrou a Reserva --> erro
    If lErro = 23928 Then Error 30000

    vbMsg = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RESERVA", objReserva.lCodigo)

    If vbMsg = vbYes Then

        'Chama Reserva_Exclui
        lErro = CF("Reserva_Exclui", objReserva)
        If lErro <> SUCESSO And lErro <> 30007 Then Error 30001

        If lErro = 30007 Then Error 30026

        Call Exclui_ReservaListBox(objReserva.lCodigo)

        lErro = Limpa_Tela_Reserva()
        If lErro <> SUCESSO Then Error 55496

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 23998
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 23999

        Case 30000, 30026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RESERVA_NAO_CADASTRADA1", Err, giFilialEmpresa, objReserva.lCodigo)

        Case 30001, 55496
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174094)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava a Reserva
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 23859

    lErro = Limpa_Tela_Reserva()
    If lErro <> SUCESSO Then Error 55497

    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 23859, 55497

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174095)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 23913

    lErro = Limpa_Tela_Reserva()
    If lErro <> SUCESSO Then Error 55498
        
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 23913, 55498

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174096)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iCodigoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()

Dim lErro As Long
Dim iCodigoAux As Integer

    Call ProdutoAlmoxarifado_PerdeFoco
    
    iCodigoAux = iCodigoAlterado
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    iCodigoAlterado = iCodigoAux
    
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Codigo_Validate

    If iCodigoAlterado = 0 Then Exit Sub

    If Len(Trim(Codigo.Text)) <> 0 Then

        lErro = Long_Critica(Codigo.Text)
        If lErro <> SUCESSO Then Error 23838

        'Se produto estiver preenchido --> Calcula o disponível
        If Len(Produto.ClipText) > 0 And Len(Trim(Almoxarifado.Text)) > 0 Then

            'Formata produto como no Banco de Dados
            lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then Error 55873

            lErro = CalculaDisponivel(sProdutoFormatado, Almoxarifado.Text, Documento.Text, Codigo.Text)
            If lErro <> SUCESSO Then Error 55510

            If Len(Quantidade.Text) <> 0 Then
                If CDbl(Quantidade.Text) > CDbl(QuantDisponivel.Caption) Then Error 55511
            End If

        End If

    End If

    iCodigoAlterado = 0

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 23838, 55873
            
        Case 55510
            Quantidade.Text = ""
            QuantDisponivel.Caption = Formata_Estoque(0)
        
        Case 55511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_RESERVADA_MAIOR", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174097)

    End Select

    Exit Sub

End Sub

Private Sub DataValidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataValidade_GotFocus()

    Call ProdutoAlmoxarifado_PerdeFoco
    
    Call MaskEdBox_TrataGotFocus(DataValidade, iAlterado)
    
End Sub

Private Sub DataValidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataValidade_Validate

    If Len(DataValidade.ClipText) = 0 Then Exit Sub

    lErro = Data_Critica(DataValidade.Text)
    If lErro <> SUCESSO Then Error 23857

    If Len(Trim(DataReserva.Caption)) <> 0 Then
        If CDate(DataValidade.Text) < CDate(DataReserva.Caption) Then Error 23858
    End If

    Exit Sub

Erro_DataValidade_Validate:

    Cancel = True


    Select Case Err

        Case 23857

        Case 23858
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VALIDADE_MENOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174098)

    End Select

    Exit Sub

End Sub

Private Sub Documento_Change()

    iAlterado = REGISTRO_ALTERADO
    iDocumentoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Documento_GotFocus()

Dim iDocAux As Integer

    Call ProdutoAlmoxarifado_PerdeFoco
    
    iDocAux = iDocumentoAlterado
    Call MaskEdBox_TrataGotFocus(Documento, iAlterado)
    iDocumentoAlterado = iDocAux
    
End Sub

Private Sub Documento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colItemPedido As New colItemPedido
Dim objItemPV As ClassItemPedido
Dim sProduto As String
Dim iPreenchido As Integer
Dim iProdutoUtilizado As Integer
Dim objItemRomaneioGrade As ClassItemRomaneioGrade

On Error GoTo Erro_Documento_Validate

    If iDocumentoAlterado = 0 Then Exit Sub

    If Len(Trim(Documento.Text)) = 0 Then Exit Sub

    'Verifica se é Long e Positivo
    lErro = Long_Critica(Documento.Text)
    If lErro <> SUCESSO Then Error 23849
    
    If TipoDocAssoc.Text = TIPO_PEDIDO Then

        objPedidoVenda.lCodigo = CLng(Documento.Text)
        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
        
        'Lê Pedido de Venda
        lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
        If lErro <> SUCESSO And lErro <> 26509 Then Error 23850
    
        'Se não achou o Pedido de Venda --> erro
        If lErro = 26509 Then Error 23851
    
        'se o produto estiver preenchido,
        If Len(Trim(Produto.ClipText)) > 0 Then
    
            'verificar se o pv possui algum item p/o produto corrente
    
            'Lê os ítens do Pedido de Venda
            lErro = CF("PedidoDeVenda_Le_Itens", objPedidoVenda)
            If lErro <> SUCESSO Then Error 23852
    
            lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
            If lErro <> SUCESSO Then Error 23844
    
            iProdutoUtilizado = False
    
            For iIndice = 1 To objPedidoVenda.colItensPedido.Count
                If objPedidoVenda.colItensPedido(iIndice).sProduto = sProduto Then
                    iProdutoUtilizado = True
                    Exit For
                End If
                For Each objItemRomaneioGrade In objPedidoVenda.colItensPedido(iIndice).colItensRomaneioGrade
                    If objItemRomaneioGrade.sProduto = sProduto Then
                        iProdutoUtilizado = True
                        Exit For
                    End If
                Next
            
            Next
            
            If Not iProdutoUtilizado Then Error 30066
    
        End If
        
    ElseIf TipoDocAssoc.Text = TIPO_PEDIDO_SRV Then

        objPedidoVenda.lCodigo = CLng(Documento.Text)
        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
        
        'Lê Pedido de Venda
        lErro = CF("PedidoServico_Le", objPedidoVenda)
        If lErro <> SUCESSO And lErro <> 188828 Then Error 23850
    
        'Se não achou o Pedido de Venda --> erro
        If lErro = 188828 Then Error 23851
    
        'se o produto estiver preenchido,
        If Len(Trim(Produto.ClipText)) > 0 Then
    
            'verificar se o pv possui algum item p/o produto corrente
    
            'Lê os ítens do Pedido de Venda
            lErro = CF("PedidoServico_Le_Itens", objPedidoVenda)
            If lErro <> SUCESSO Then Error 23852
    
            lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
            If lErro <> SUCESSO Then Error 23844
    
            iProdutoUtilizado = False
    
            For iIndice = 1 To objPedidoVenda.colItensPedido.Count
                If objPedidoVenda.colItensPedido(iIndice).sProduto = sProduto Then
                    iProdutoUtilizado = True
                    Exit For
                End If
                For Each objItemRomaneioGrade In objPedidoVenda.colItensPedido(iIndice).colItensRomaneioGrade
                    If objItemRomaneioGrade.sProduto = sProduto Then
                        iProdutoUtilizado = True
                        Exit For
                    End If
                Next
            
            Next
            
            If Not iProdutoUtilizado Then Error 30066
    
        End If
    End If

    iDocumentoAlterado = 0

    Exit Sub

Erro_Documento_Validate:

    Cancel = True


    Select Case Err

        Case 23844, 23850

        Case 23849

        Case 23851
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", Err, objPedidoVenda.lCodigo)

        Case 23852

        Case 30066
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITENSPV_NAO_UTILIZAM_PRODUTO", Err, objPedidoVenda.lCodigo, Produto.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174099)

    End Select

    Exit Sub

End Sub

Private Sub DocumentoLabel_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colSelecao As Collection
Dim sNomeBrowse As String

On Error GoTo Erro_DocumentoLabel_Click

    If TipoDocAssoc.Text = TIPO_PEDIDO Or TipoDocAssoc.Text = TIPO_PEDIDO_SRV Then
    
        If TipoDocAssoc.Text = TIPO_PEDIDO Then
            sNomeBrowse = "PedidoVendaLista"
        Else
            sNomeBrowse = "PedidoServico_Lista"
        End If

        If Len(Trim(Documento.Text)) > 0 Then
            objPedidoVenda.lCodigo = CLng(Documento.Text)
        End If
    
        Call Chama_Tela(sNomeBrowse, colSelecao, objPedidoVenda, objEventoDocumento)
        
    End If

    Exit Sub

Erro_DocumentoLabel_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174100)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigo As New Collection

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoDocumento = New AdmEvento
    Set objEventoAlmoxarifado = New AdmEvento

    'Inicializa Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then Error 23809

    'Formata quantidade no Formato Estoque
    Quantidade.Format = FORMATO_ESTOQUE

    'Coloca em Data de Reserva a data atual
    DataReserva.Caption = Format(gdtDataHoje, "dd/mm/yyyy")

'    'Carrega TreeView de Produtos
'    lErro = CF("Carga_Arvore_Produto_Reserva",Produtos.Nodes)
'    If lErro <> SUCESSO Then Error 23810

    'Preenche ListBox de Reservas com código das reservas no Banco de Dados
    lErro = CF("Reservas_Le_Codigo", colCodigo)
    If lErro <> SUCESSO Then Error 23811

    For iIndice = 1 To colCodigo.Count

        Reservas.AddItem colCodigo(iIndice)
        Reservas.ItemData(Reservas.NewIndex) = colCodigo(iIndice)

    Next

    'carrega a listbox de almoxarifados
    lErro = Carga_Almoxarifados("")
    If lErro <> SUCESSO Then Error 55494

    'Preenche List de ComboBox TipoDocAssoc
    TipoDocAssoc.AddItem TIPO_MANUTENCAO
    TipoDocAssoc.ItemData(TipoDocAssoc.NewIndex) = TIPO_MANUTENCAO_COD

    'Se modulo de Faturamento faz parte do pacote
    If gcolModulo.Ativo(MODULO_FATURAMENTO) = MODULO_ATIVO Then

        TipoDocAssoc.AddItem TIPO_PEDIDO
        TipoDocAssoc.ItemData(TipoDocAssoc.NewIndex) = TIPO_PEDIDO_COD

    End If
    
    'Se modulo de Faturamento faz parte do pacote
    If gcolModulo.Ativo(MODULO_SERVICOS) = MODULO_ATIVO Then

        TipoDocAssoc.AddItem TIPO_PEDIDO_SRV
        TipoDocAssoc.ItemData(TipoDocAssoc.NewIndex) = TIPO_PEDIDO_SRV_COD

    End If

    'Seleciona o primeiro item da List(TIPO_MANUTENCAO)
    TipoDocAssoc.ListIndex = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 23809, 23810, 23811, 55494

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174101)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objReserva As ClassReserva) As Long

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objReserva Is Nothing) Then

        'Lê Reserva no Banco de Dados a partir do código
        lErro = CF("Reserva_Le", objReserva)
        If lErro <> SUCESSO And lErro <> 23928 Then Error 23813

        If lErro <> 23928 Then 'Se a Reserva Existe

            lErro = Preenche_Tela(objReserva)
            If lErro <> SUCESSO Then Error 23814

        Else
            'Apenas exibe o código
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objReserva.lCodigo)
            Codigo.PromptInclude = True

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 23813, 23814

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174102)

    End Select

    Exit Function

End Function

Private Function Preenche_Tela(objReserva As ClassReserva) As Long
'Mostra os dados da Reserva na tela

Dim lErro As Long
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim colItemPedido As New colItemPedido
Dim sDocumento As String
Dim sCodigo As String

On Error GoTo Erro_Preenche_Tela

    objProduto.sCodigo = objReserva.sProduto
    
    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 23820

    'Se não achou o Produto --> erro
    If lErro = 28030 Then Error 23821

    objAlmoxarifado.iCodigo = objReserva.iAlmoxarifado
    
    'Lê o Almoxarifado
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then Error 23822

    'Se não achou o Almoxarifado --> erro
    If lErro = 25056 Then Error 23823

    'Se (Documento não estiver preenchido ) e (almoxarifado e Produto estiverem preenchidos ) --> Calcula o Disponível
    If Len(Trim(objAlmoxarifado.sNomeReduzido)) <> 0 And Len(objProduto.sCodigo) <> 0 Then

        sDocumento = ""
        If objReserva.lDocOrigem <> 0 Then sDocumento = CStr(objReserva.lDocOrigem)
         
        sCodigo = ""
        If objReserva.lCodigo <> 0 Then sCodigo = CStr(objReserva.lCodigo)

        lErro = CalculaDisponivel(objProduto.sCodigo, objAlmoxarifado.sNomeReduzido, sDocumento, sCodigo)
        If lErro <> SUCESSO Then Error 23931
        
    Else
    
        QuantDisponivel.Caption = ""

    End If

    'Preenche a Tela
    Codigo.PromptInclude = False
    Codigo.Text = CStr(objReserva.lCodigo)
    Codigo.PromptInclude = True

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, Descricao)
    If lErro <> SUCESSO Then Error 23824

    UnidadeMedida.Caption = objProduto.sSiglaUMEstoque

    Almoxarifado.Text = objAlmoxarifado.sNomeReduzido

    'Formata quantidade antes de colocar na Tela
    Quantidade.Text = Formata_Estoque(objReserva.dQuantidade)

    If objReserva.iTipoDoc = TIPO_PEDIDO_GRADE Then objReserva.iTipoDoc = TIPO_PEDIDO_COD

    'Seleciona TipoDocAssoc
    For iIndice = 0 To TipoDocAssoc.ListCount - 1

        If TipoDocAssoc.ItemData(iIndice) = objReserva.iTipoDoc Then
            TipoDocAssoc.ListIndex = iIndice
            Exit For
        End If

    Next

    If TipoDocAssoc.Text = TIPO_PEDIDO Or TipoDocAssoc.Text = TIPO_PEDIDO_SRV Then

        Documento.PromptInclude = False
        Documento.Text = CStr(objReserva.lDocOrigem)
        Documento.PromptInclude = True
    Else
        Documento.PromptInclude = False
        Documento.Text = ""
        Documento.PromptInclude = True
    End If

    DataReserva.Caption = Format(objReserva.dtDataReserva, "dd/mm/yyyy")

    Call DateParaMasked(DataValidade, objReserva.dtDataValidade)

    Responsabilidade.Text = objReserva.sResponsavel

    'carrega a listbox de almoxarifados
    lErro = Carga_Almoxarifados(objProduto.sCodigo)
    If lErro <> SUCESSO Then Error 55501

    iAlterado = 0

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = Err

    Select Case Err

        Case 23821
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case 23823
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", Err, objAlmoxarifado.iCodigo)

        Case 23822, 23824, 55501

        Case 23931
            Quantidade.Text = ""
            QuantDisponivel.Caption = Formata_Estoque(0)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174103)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim objReserva As New ClassReserva

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Reserva"

    'Lê os atributos de objReserva que aparecem na Tela
    lErro = Move_Tela_Memoria(objReserva)
    If lErro <> SUCESSO Then Error 23825

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Quantidade", objReserva.dQuantidade, 0, "Quantidade"
    colCampoValor.Add "DataReserva", objReserva.dtDataReserva, 0, "DataReserva"
    colCampoValor.Add "DataValidade", objReserva.dtDataValidade, 0, "DataValidade"
    colCampoValor.Add "Almoxarifado", objReserva.iAlmoxarifado, 0, "Almoxarifado"
    colCampoValor.Add "FilialEmpresa", objReserva.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "NumIntOrigem", objReserva.lNumIntOrigem, 0, "NumIntOrigem"
    colCampoValor.Add "TipoDoc", objReserva.iTipoDoc, 0, "TipoDoc"
    colCampoValor.Add "Codigo", objReserva.lCodigo, 0, "Codigo"
    colCampoValor.Add "DocOrigem", objReserva.lDocOrigem, 0, "DocOrigem"
    colCampoValor.Add "CodUsuario", objReserva.sCodUsuario, STRING_SIGLA_USUARIO, "CodUsuario"
    colCampoValor.Add "Produto", objReserva.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "Responsavel", objReserva.sResponsavel, STRING_RESPONSAVEL_RESERVA, "Responsavel"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 23825

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174104)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objReserva As New ClassReserva

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objReserva
    objReserva.dQuantidade = colCampoValor.Item("Quantidade").vValor
    objReserva.dtDataReserva = colCampoValor.Item("DataReserva").vValor
    objReserva.dtDataValidade = colCampoValor.Item("DataValidade").vValor
    objReserva.iAlmoxarifado = colCampoValor.Item("Almoxarifado").vValor
    objReserva.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objReserva.lNumIntOrigem = colCampoValor.Item("NumIntOrigem").vValor
    objReserva.iTipoDoc = colCampoValor.Item("TipoDoc").vValor
    objReserva.lCodigo = colCampoValor.Item("Codigo").vValor
    objReserva.lDocOrigem = colCampoValor.Item("DocOrigem").vValor
    objReserva.sCodUsuario = colCampoValor.Item("CodUsuario").vValor
    objReserva.sProduto = colCampoValor.Item("Produto").vValor
    objReserva.sResponsavel = colCampoValor.Item("Responsavel").vValor

    lErro = Preenche_Tela(objReserva)
    If lErro <> SUCESSO Then Error 23826

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 23826

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174105)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoCodigo = Nothing
    Set objEventoProduto = Nothing
    Set objEventoDocumento = Nothing
    Set objEventoAlmoxarifado = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
   lErro = ComandoSeta_Liberar(Me.Name)
   
End Sub


Private Sub objEventoDocumento_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

    Set objPedidoVenda = obj1

    Documento.PromptInclude = False
    Documento.Text = CStr(objPedidoVenda.lCodigo)
    Documento.PromptInclude = True

    Call Documento_Validate(bSGECancelDummy)

    Me.Show

End Sub

Private Sub Padrao_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iAlmoxarifadoPadrao As Integer
Dim bCancel As Boolean

On Error GoTo Erro_Padrao_Click

    'Se produto estiver preenchido
    If Len(Produto.ClipText) = 0 Then Exit Sub

    'Tenta selecionar Almoxarifado Padrão para o Produto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 23842

    'Lê o Almoxarifado Padrão
    lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, sProdutoFormatado, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO And lErro <> 23796 Then Error 23843

    If iAlmoxarifadoPadrao <> 0 Then
        Almoxarifado.Text = iAlmoxarifadoPadrao
        Call Almoxarifado_Validate(bCancel)
    End If

    Exit Sub

Erro_Padrao_Click:

    Select Case Err

        Case 23842, 23843

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174106)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Change()

    iProdutoAlterado = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

''Mostra os Produtos no lugar das Reservas
'
'    Produtos.Visible = True
'    Reservas.Visible = False
'    Almoxarifados.Visible = False
'
'    LabelProduto.Visible = True
'    LabelReserva.Visible = False
'    LabelAlmoxarifado.Visible = False
'
'End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Produto_Validate

    'Se o produto foi alterado
    If iProdutoAlterado <> 0 Then

        'Se produto não estiver preenchido --> limpa descrição e unidade de medida
        If Len(Trim(Produto.ClipText)) = 0 Then
            Descricao.Caption = ""
            UnidadeMedida.Caption = ""
            QuantDisponivel.Caption = ""

        Else
            'Caso esteja preenchido
            lErro = CF("Produto_Critica", Produto.Text, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 25041 Then Error 23839

            If lErro = 25041 Then Error 23840

            Descricao.Caption = ""
            UnidadeMedida.Caption = ""
            QuantDisponivel.Caption = ""

            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

                If objProduto.iControleEstoque <> PRODUTO_CONTROLE_RESERVA Then Error 23947

                Descricao.Caption = objProduto.sDescricao
                UnidadeMedida.Caption = objProduto.sSiglaUMEstoque

                'Se Almoxarifado estiver preenchido --> Calcula o Disponível
                If Len(Trim(Almoxarifado.Text)) > 0 Then

                    lErro = CalculaDisponivel(objProduto.sCodigo, Almoxarifado.Text, Documento.Text, Codigo.Text)
                    If lErro <> SUCESSO Then Error 23949

                    If Len(Quantidade.Text) <> 0 Then
                        If CDbl(Quantidade.Text) > CDbl(QuantDisponivel.Caption) Then Error 55508
                    End If

                End If

            End If

        End If

        'carrega a listbox de almoxarifados
        lErro = Carga_Almoxarifados(objProduto.sCodigo)
        If lErro <> SUCESSO Then Error 55499

        iProdutoAlterado = 0

    End If

    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case Err

        Case 23839, 55499

        Case 23949
            Quantidade.Text = ""
            QuantDisponivel.Caption = Formata_Estoque(0)

        Case 23840
            Descricao.Caption = ""
            UnidadeMedida.Caption = ""

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                Call Chama_Tela("Produto", objProduto)
            End If

        Case 23947
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_ADMITE_RESERVA", Err, objProduto.sCodigo)

        Case 55508
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_RESERVADA_MAIOR", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174107)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoLabel_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_ProdutoLabel_Click

    'Verifica se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 23828

        objProduto.sCodigo = sProdutoFormatado

    End If

    'Adiciona filtro para reserva e estoque
    colSelecao.Add PRODUTO_CONTROLE_RESERVA

    Call Chama_Tela("ProdutoLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_ProdutoLabel_Click:

    Select Case gErr

        Case 23828

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174108)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 30102

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 30103

    'Se Almoxarifado estiver preenchido Calcula o Disponível
    If Len(Trim(Almoxarifado.Text)) > 0 Then

        lErro = CalculaDisponivel(objProduto.sCodigo, Almoxarifado.Text, Documento.Text, Codigo.Text)
        If lErro <> SUCESSO Then gError 23943

    End If

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, Descricao)
    If lErro <> SUCESSO Then gError 23829

    'Preenche unidade de medida com SiglaUMEstoque
    UnidadeMedida.Caption = objProduto.sSiglaUMEstoque

    'carrega a listbox de almoxarifados
    lErro = Carga_Almoxarifados(objProduto.sCodigo)
    If lErro <> SUCESSO Then gError 55500

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 23829, 23942, 30102, 55500

        Case 23943
            Quantidade.Text = ""
            QuantDisponivel.Caption = Formata_Estoque(0)

        Case 30103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174109)

    End Select

    Exit Sub

End Sub

Private Sub CodigoLabel_Click()

Dim lErro As Long
Dim objReserva As New ClassReserva
Dim colSelecao As Collection

On Error GoTo Erro_CodigoLabel_Click

    'Chama Move_Tela_Memoria
    lErro = Move_Tela_Memoria(objReserva)
    If lErro <> SUCESSO Then Error 23830

    Call Chama_Tela("ReservaLista", colSelecao, objReserva, objEventoCodigo)

    Exit Sub

Erro_CodigoLabel_Click:

    Select Case Err

        Case 23830

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174110)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objReserva As ClassReserva

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objReserva = obj1

    lErro = Preenche_Tela(objReserva)
    If lErro <> SUCESSO Then Error 23831

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case Err

        Case 23831

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174111)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'Private Sub Produtos_Expand(ByVal objNode As MSComctlLib.Node)
'
'Dim lErro As Long
'
'On Error GoTo Erro_Produtos_Expand
'
'    If objNode.Tag <> NETOS_NA_ARVORE Then
'
'        'move os dados do plano de contas do banco de dados para a arvore colNodes.
'        lErro = CF("Carga_Arvore_Produto_Netos_Reserva",objNode, Produtos.Nodes)
'        If lErro <> SUCESSO Then Error 48086
'
'    End If
'
'    Exit Sub
'
'Erro_Produtos_Expand:
'
'    Select Case Err
'
'        Case 48086
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174112)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Produtos_NodeClick(ByVal Node As MSComctlLib.Node)
'
'Dim lErro As Long
'Dim objProduto As New ClassProduto
'
'On Error GoTo Erro_Produtos_NodeClick
'
'    'Se nó tiver Filhos --> Sai
'    If Node.Children > 0 Then Exit Sub
'
'    objProduto.sCodigo = Mid(Node.Key, 2)
'
'    'lê o Produto
'    lErro = CF("Produto_Le",objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then Error 23832
'
'    'Se não achou o Produto --> erro
'    If lErro = 28030 Then Error 23833
'
'    'Se produto for gerencial sai
'    If objProduto.iGerencial = GERENCIAL Then Error 49942
'
'    'Se Almoxarifado estiver preenchido --> Calcula o Disponível
'    If Len(Trim(Almoxarifado.Text)) > 0 Then
'
'        lErro = CalculaDisponivel(objProduto.sCodigo, Almoxarifado.Text, Documento.Text, Codigo.Text)
'        If lErro <> SUCESSO Then Error 55502
'
'    End If
'
'    lErro = CF("Traz_Produto_MaskEd",objProduto.sCodigo, Produto, Descricao)
'    If lErro <> SUCESSO Then Error 23834
'
'    UnidadeMedida.Caption = objProduto.sSiglaUMEstoque
'
'    'carrega a listbox de almoxarifados
'    lErro = Carga_Almoxarifados(objProduto.sCodigo)
'    If lErro <> SUCESSO Then Error 55503
'
'    Exit Sub
'
'Erro_Produtos_NodeClick:
'
'    Select Case Err
'
'        Case 23832, 49942, 55503
'
'        Case 23833
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, Node.Child.Text)
'
'        Case 23834, 55502
'            Quantidade.Text = ""
'            QuantDisponivel.Caption = ""
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174113)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Call ProdutoAlmoxarifado_PerdeFoco

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Quantidade_Validate

    'Veifica se Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) = 0 Then Exit Sub

    'Critica a Quantidade
    lErro = Valor_Positivo_Critica(Quantidade.Text)
    If lErro <> SUCESSO Then Error 23847

    If Len(QuantDisponivel.Caption) <> 0 Then
        If CDbl(Quantidade.Text) > CDbl(QuantDisponivel.Caption) Then Error 23848
    End If

    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True


    Select Case Err

        Case 23847

        Case 23848
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_RESERVADA_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174114)

    End Select

    Exit Sub

End Sub

Private Sub Reservas_DblClick()

Dim lErro As Long
Dim objReserva As New ClassReserva

On Error GoTo Erro_Reservas_DblClick

    If Reservas.ListIndex = -1 Then Exit Sub

    objReserva.lCodigo = Reservas.ItemData(Reservas.ListIndex)

    'Lê Reserva no Banco de Dados a partir do código
    lErro = CF("Reserva_Le", objReserva)
    If lErro <> SUCESSO And lErro <> 23928 Then Error 23835

    'Se não encontrou a Reserva --> erro
    If lErro = 23928 Then Error 23836

    lErro = Preenche_Tela(objReserva)
    If lErro <> SUCESSO Then Error 23837

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Reservas_DblClick:

    Select Case Err

        Case 23835

        Case 23836
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RESERVA_NAO_CADASTRADA", Err, objReserva.lCodigo)

        Case 23837

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174115)

    End Select

    Exit Sub

End Sub

Private Sub Responsabilidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Responsabilidade_GotFocus()

    Call ProdutoAlmoxarifado_PerdeFoco

End Sub

Private Sub TipoDocAssoc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoDocAssoc_Click()

    If TipoDocAssoc.ListIndex = -1 Then Exit Sub

    iAlterado = REGISTRO_ALTERADO

    If TipoDocAssoc.List(TipoDocAssoc.ListIndex) = TIPO_MANUTENCAO Then
        Documento.PromptInclude = False
        Documento.Text = ""
        Documento.PromptInclude = True

        Documento.Enabled = False
        DocumentoLabel.Enabled = False

        Documento.Visible = False
        DocumentoLabel.Visible = False

    ElseIf TipoDocAssoc.List(TipoDocAssoc.ListIndex) = TIPO_PEDIDO Then
        
        Documento.Enabled = True
        DocumentoLabel.Enabled = True

        DocumentoLabel.Caption = "Ped. Venda"

        Documento.Visible = True
        DocumentoLabel.Visible = True

    ElseIf TipoDocAssoc.List(TipoDocAssoc.ListIndex) = TIPO_PEDIDO_SRV Then
        
        Documento.Enabled = True
        DocumentoLabel.Enabled = True
        
        DocumentoLabel.Caption = "Pedido SRV"

        Documento.Visible = True
        DocumentoLabel.Visible = True

    End If

End Sub

Public Function Gravar_Registro()
'Valida os dados para gravar Reserva

Dim lErro As Long
Dim objReserva As New ClassReserva

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Error 23860

    'Verifica se produto foi preenchido
    If Len(Produto.ClipText) = 0 Then Error 23861

    'Verifica se Almoxarifado foi preenchido
    If Len(Trim(Almoxarifado.Text)) = 0 Then Error 23862

    'Verifica se a Quantidade foi preenchida
    If Len(Trim(Quantidade.Text)) = 0 Then Error 23863

    'Se Tipo do Documento estiver preenchido
    If TipoDocAssoc.Text = TIPO_PEDIDO Then

        'Verifica se Documento está preenchido
        If Len(Trim(Documento.Text)) = 0 Then Error 23864

    End If

    If Len(Trim(Quantidade.Text)) > 0 Then objReserva.dQuantidade = CDbl(Quantidade.Text)
    
    If objReserva.dQuantidade <= 0 Then Error 55874

    'Preenche objReserva
    lErro = Move_Tela_Memoria(objReserva)
    If lErro <> SUCESSO Then Error 23865

    'Grava Reserva no Banco de Dados
    lErro = CF("Reserva_Grava", objReserva)
    If lErro <> SUCESSO Then Error 23866

    'Exclui da ListBoxReservas se achar
    Reservas.ListIndex = -1

    'Obs.: objReserva.lCodigo pode ter sido trocado dentro de Reserva_Grava
    Call Exclui_ReservaListBox(objReserva.lCodigo)
    Call Inclui_ReservaListBox(objReserva.lCodigo)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 23860
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 23861
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_INFORMADO", Err)

        Case 23862
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", Err)

        Case 23863
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_RESERVA_NAO_PREENCHIDA", Err)

        Case 23864
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_PREENCHIDO", Err)

        Case 23865, 23866

        Case 55874
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_RESERVADA_NAO_POSITIVA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174116)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objReserva As ClassReserva) As Long
'Preenche o objReserva com os dados da tela

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objItemPedido As ClassItemPedido

On Error GoTo Erro_Move_Tela_Memoria

    If Len(Trim(DataReserva.Caption)) > 0 Then objReserva.dtDataReserva = CDate(DataReserva.Caption)

    If Len(Trim(DataValidade.ClipText)) > 0 Then
        objReserva.dtDataValidade = CDate(DataValidade.Text)
    Else
        objReserva.dtDataValidade = DATA_NULA
    End If

    If Len(Trim(Almoxarifado.Text)) > 0 Then

        objAlmoxarifado.sNomeReduzido = Almoxarifado.Text
        'Lê o Almoxarifado
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then Error 30067

        'Se não encontrou o Almoxaifado --> erro
        If lErro = 25060 Then Error 30068

        objReserva.iAlmoxarifado = objAlmoxarifado.iCodigo

    End If

    objReserva.iFilialEmpresa = giFilialEmpresa

    If TipoDocAssoc.Text = TIPO_PEDIDO Then

        objReserva.iTipoDoc = TipoDocAssoc.ItemData(TipoDocAssoc.ListIndex)

        'Verifica se o Documento asscociado foi preenchido
        If Len(Trim(Documento.Text)) > 0 Then

            objReserva.lDocOrigem = CLng(Documento.Text)

            lErro = PesquisaItem(objItemPedido)
            If lErro <> SUCESSO Then Error 30090

            If objItemPedido.iPossuiGrade = DESMARCADO Then
                objReserva.lNumIntOrigem = objItemPedido.lNumIntDoc
            Else
                objReserva.lNumIntOrigem = objItemPedido.colItensRomaneioGrade(1).lNumIntDoc
                objReserva.iTipoDoc = TIPO_PEDIDO_GRADE
            End If

        End If
    ElseIf TipoDocAssoc.Text = TIPO_PEDIDO_SRV Then

        objReserva.iTipoDoc = TipoDocAssoc.ItemData(TipoDocAssoc.ListIndex)

        'Verifica se o Documento asscociado foi preenchido
        If Len(Trim(Documento.Text)) > 0 Then
            
            objReserva.lDocOrigem = CLng(Documento.Text)
            
            lErro = PesquisaItemSRV(objItemPedido)
            If lErro <> SUCESSO Then Error 30090
            
            objReserva.lNumIntOrigem = objItemPedido.lNumIntDoc
        End If
        
    Else
        objReserva.iTipoDoc = TipoDocAssoc.ItemData(TipoDocAssoc.ListIndex)
    End If

    objReserva.sCodUsuario = gsUsuario

    If Len(Trim(Codigo.Text)) > 0 Then objReserva.lCodigo = CLng(Codigo.Text)

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 30070

    objReserva.sProduto = sProdutoFormatado

    objReserva.sResponsavel = Responsabilidade.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 30067, 30090

        Case 30068
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", Err, objAlmoxarifado.sNomeReduzido)

        Case 30070

        Case 55874
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_RESERVADA_NAO_POSITIVA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174117)

    End Select

    Exit Function

End Function

Private Sub TipoDocAssoc_GotFocus()

    Call ProdutoAlmoxarifado_PerdeFoco

End Sub

Private Sub UpDownValidade_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownValidade_DownClick

    'Verifica se a Data da Validade foi preenchida
    If Len(Trim(DataValidade.ClipText)) = 0 Then Exit Sub

    'Diminui a Data em um dia
    lErro = Data_Up_Down_Click(DataValidade, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 23916

    Exit Sub

Erro_UpDownValidade_DownClick:

    Select Case Err

        Case 23916

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174118)

    End Select

    Exit Sub

End Sub

Private Sub UpDownValidade_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownValidade_UpClick

    'Verifica se a Data de Validade foi preenchida
    If Len(Trim(DataValidade.ClipText)) = 0 Then Exit Sub

    'Aumenta a Data em um dia
    lErro = lErro = Data_Up_Down_Click(DataValidade, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 23915

    Exit Sub

Erro_UpDownValidade_UpClick:

    Select Case Err

        Case 23915

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174119)

    End Select

    Exit Sub

End Sub

Private Function PesquisaItem(Optional objItemPedido As ClassItemPedido) As Long
'Lê o pedido de venda correspondente e determina o ítem corresponde ao produto

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_PesquisaItem

    Set objItemPedido = New ClassItemPedido

    objItemPedido.lCodPedido = CLng(Documento.Text)

    'Formata o produto para o Banco de Dados
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 30069

    objItemPedido.sProduto = sProdutoFormatado
    
    'Lê um Item de pedido
    lErro = CF("ItemPedido_Le_Produto", objItemPedido)
    If lErro <> SUCESSO And lErro <> 30062 Then Error 23950

    'Se não encontrou o Item --> erro
    If lErro = 30062 Then
    
        lErro = CF("ItemPedidoGrade_Le_Produto", objItemPedido)
        If lErro <> SUCESSO And lErro <> 86382 Then Error 23950
    
        If lErro <> SUCESSO Then Error 23952

    End If
    
    PesquisaItem = SUCESSO

    Exit Function

Erro_PesquisaItem:

    PesquisaItem = Err

    Select Case Err

        Case 23950, 30069

        Case 23952, 30085
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_SEM_PRODUTO", Err, Produto.Text)
            Produto.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174120)

    End Select

    Exit Function

End Function

Private Function PesquisaItemSRV(Optional objItemPedido As ClassItemPedido) As Long
'Lê o pedido de venda correspondente e determina o ítem corresponde ao produto

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objItemPedidoAux As ClassItemPedido
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim bAchou As Boolean

On Error GoTo Erro_PesquisaItemSRV

    Set objItemPedido = New ClassItemPedido

    objPedidoVenda.iFilialEmpresa = giFilialEmpresa
    objPedidoVenda.lCodigo = CLng(Documento.Text)

    'Formata o produto para o Banco de Dados
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 30069
    
    'Lê um Item de pedido
    lErro = CF("PedidoServico_Le_Todos_Completo", objPedidoVenda)
    If lErro <> SUCESSO Then Error 23950

    bAchou = False
    For Each objItemPedidoAux In objPedidoVenda.colItensPedido
        If objItemPedidoAux.sProduto = sProdutoFormatado Then
            Set objItemPedido = objItemPedidoAux
            bAchou = True
            Exit For
        End If
    Next
    
    If Not bAchou Then Error 30085
    
    PesquisaItemSRV = SUCESSO

    Exit Function

Erro_PesquisaItemSRV:

    PesquisaItemSRV = Err

    Select Case Err

        Case 23950, 30069

        Case 30085
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_SEM_PRODUTO", Err, Produto.Text)
            Produto.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174120)

    End Select

    Exit Function

End Function

Function CalculaDisponivel(sProdutoFormatado As String, sAlmoxNomeRed As String, sDocumento As String, sCodigo As String) As Long
'Calcula a Quantidade Disponível

Dim lErro As Long
Dim dQuantDisponivel  As Double
Dim iProdutoPreenchido As Integer
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objReservaItem As New ClassReserva
Dim objItemPedido As New ClassItemPedido
Dim objReserva As New ClassReserva

On Error GoTo Erro_CalculaDisponivel

    objAlmoxarifado.sNomeReduzido = sAlmoxNomeRed
    
    'Procurar o código do almoxarifado a partir do NomeReduzido
    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25060 Then Error 23955

    'Se não encontrou o Almoxarifado --> erro
    If lErro = 25060 Then Error 23973

    objEstoqueProduto.sProduto = sProdutoFormatado
    objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
    
    'Lê Estoque do Produto neste almoxarifado
    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 21306 Then Error 23956

    'Se não encontrou o Estoque do Produto --> erro
    If lErro = 21306 Then Error 23957
    
    If Len(Trim(sCodigo)) <> 0 Then
    
        objReserva.lCodigo = CLng(sCodigo)

        'Lê a Reserva
        lErro = CF("Reserva_Le", objReserva)
        If lErro <> SUCESSO And lErro <> 23928 Then Error 30077

        'Não achou a Reserva
        If lErro = 23928 Then
            objReserva.dQuantidade = 0
        Else
            If Len(Trim(sDocumento)) <> 0 Then
                If CLng(sDocumento) <> objReserva.lDocOrigem Then Error 55504
            End If
        End If
        
        QuantDisponivel.Caption = Formata_Estoque(objEstoqueProduto.dQuantDisponivel + objReserva.dQuantidade)

    Else 'Outros casos

        QuantDisponivel.Caption = Formata_Estoque(objEstoqueProduto.dQuantDisponivel)

    End If

    iProdutoAlterado = 0

    CalculaDisponivel = SUCESSO

    Exit Function

Erro_CalculaDisponivel:

    CalculaDisponivel = Err

    Select Case Err

        Case 23954, 23955, 23956

        Case 23957
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_INEXISTENTE", Err, objEstoqueProduto.sProduto)
            QuantDisponivel.Caption = Formata_Estoque(0)
            Quantidade.Text = ""

        Case 23973
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", Err, objAlmoxarifado.sNomeReduzido)
            QuantDisponivel.Caption = ""
            Quantidade.Text = ""

        Case 55504
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_ORIGEM_RESERVA", Err, CLng(sDocumento), objReserva.lDocOrigem)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174121)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Reserva() As Long

Dim lErro As Long, lCodigo As Long

On Error GoTo Erro_Limpa_Tela_Reserva

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Funcção genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    'Limpa os campos da tela que não foram limpos pela função acima
    Descricao.Caption = ""
    UnidadeMedida.Caption = ""
    QuantDisponivel.Caption = ""
    DataReserva.Caption = Format(gdtDataHoje, "dd/mm/yyyy")

    'Seleciona Manutenção de Reserva
    TipoDocAssoc.ListIndex = 0
    
    'carrega a listbox de almoxarifados
    lErro = Carga_Almoxarifados("")
    If lErro <> SUCESSO Then Error 55495
    
    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True
    
    iCodigoAlterado = 0
    iDocumentoAlterado = 0
    iProdutoAlterado = 0
    iAlmoxarifadoAlterado = 0
    iAlterado = 0
    
    Documento.Enabled = False
    DocumentoLabel.Enabled = False

    Documento.Visible = False
    DocumentoLabel.Visible = False
    
    Limpa_Tela_Reserva = SUCESSO

    Exit Function

Erro_Limpa_Tela_Reserva:

    Limpa_Tela_Reserva = Err

    Select Case Err

        Case 55495

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174122)

    End Select

    Exit Function

End Function

Sub ProdutoAlmoxarifado_PerdeFoco()
'Mostra as Reservas no lugar dos Almoxarifados

    Reservas.Visible = True
    Almoxarifados.Visible = False
'    Produtos.Visible = False

    LabelReserva.Visible = True
    LabelAlmoxarifado.Visible = False
'    LabelProduto.Visible = False

End Sub

Sub Exclui_ReservaListBox(lCodigo As Long)

Dim iIndice As Integer

    For iIndice = 0 To Reservas.ListCount - 1

        If Reservas.ItemData(iIndice) = lCodigo Then
            Reservas.RemoveItem (iIndice)
            Exit For
        End If

    Next

End Sub

Sub Inclui_ReservaListBox(lCodigo As Long)

Dim iIndice As Integer

    For iIndice = 0 To Reservas.ListCount - 1

        If Reservas.ItemData(iIndice) > lCodigo Then Exit For

    Next

    Reservas.AddItem lCodigo, iIndice
    Reservas.ItemData(iIndice) = lCodigo

End Sub

Private Function Carga_Almoxarifados(sProduto As String) As Long
'carrega a listbox de almoxarifados

Dim lErro As Long
Dim colCodNome As New AdmColCodigoNome
Dim colAlmoxarifados As New Collection
Dim objAlmoxarifado As ClassAlmoxarifado
Dim iIndice As Integer

On Error GoTo Erro_Carga_Almoxarifados

    Almoxarifados.Clear

    'Se o produto não estiver preenchido
    If Len(sProduto) = 0 Then

        'Preenche ListBox Almoxarifados com Nomes Reduzidos de todos Almoxarifados de todas as filiais
        lErro = CF("Almoxarifados_Le_Cod_NomeRed", colCodNome)
        If lErro <> SUCESSO Then Error 55493
    
        For iIndice = 1 To colCodNome.Count
    
            Almoxarifados.AddItem colCodNome(iIndice).sNome
            Almoxarifados.ItemData(Almoxarifados.NewIndex) = colCodNome(iIndice).iCodigo
    
        Next

    Else
    
        'Preenche ListBox Almoxarifados com Nomes Reduzidos dos Almoxarifados associados ao produto sProduto
        lErro = CF("EstoqueProduto_Le_Almoxarifados1", sProduto, colAlmoxarifados)
        If lErro <> SUCESSO Then Error 55492
    
        For Each objAlmoxarifado In colAlmoxarifados
    
            Almoxarifados.AddItem objAlmoxarifado.sNomeReduzido
            Almoxarifados.ItemData(Almoxarifados.NewIndex) = objAlmoxarifado.iCodigo
    
        Next
    
    End If

    Carga_Almoxarifados = SUCESSO
    
    Exit Function
    
Erro_Carga_Almoxarifados:

    Carga_Almoxarifados = Err
    
    Select Case Err

        Case 55492, 55493

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174123)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RESERVA_PRODUTO
    Set Form_Load_Ocx = Me
    Caption = "Reserva"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Reserva"
    
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
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'**** fim do trecho a ser copiado *****

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call ProdutoLabel_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call AlmoxarifadoLabel_Click
        ElseIf Me.ActiveControl Is Documento Then
            Call DocumentoLabel_Click
        End If
        
    ElseIf KeyCode = KEYCODE_CODBARRAS Then
        Call Trata_CodigoBarras1
        
    End If
    
End Sub




Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub DataReserva_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataReserva, Source, X, Y)
End Sub

Private Sub DataReserva_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataReserva, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub QuantDisponivel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantDisponivel, Source, X, Y)
End Sub

Private Sub QuantDisponivel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantDisponivel, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub DocumentoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DocumentoLabel, Source, X, Y)
End Sub

Private Sub DocumentoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DocumentoLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelAlmoxarifado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAlmoxarifado, Source, X, Y)
End Sub

Private Sub LabelAlmoxarifado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAlmoxarifado, Button, Shift, X, Y)
End Sub

'Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelProduto, Source, X, Y)
'End Sub
'
'Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelProduto, Button, Shift, X, Y)
'End Sub

Private Sub UnidadeMedida_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UnidadeMedida, Source, X, Y)
End Sub

Private Sub UnidadeMedida_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UnidadeMedida, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub

Private Sub CodigoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabel, Source, X, Y)
End Sub

Private Sub CodigoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabel, Button, Shift, X, Y)
End Sub

Private Sub ProdutoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoLabel, Source, X, Y)
End Sub

Private Sub ProdutoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLabel, Button, Shift, X, Y)
End Sub

Private Sub AlmoxarifadoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AlmoxarifadoLabel, Source, X, Y)
End Sub

Private Sub AlmoxarifadoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AlmoxarifadoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub LabelReserva_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelReserva, Source, X, Y)
End Sub

Private Sub LabelReserva_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelReserva, Button, Shift, X, Y)
End Sub

Public Function Trata_CodigoBarras1() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoEnxuto As String
Dim sCodBarras As String
Dim sCodBarrasOriginal As String
Dim dCusto As Double

On Error GoTo Erro_Trata_CodigoBarras1

            
            objProduto.lErro = 1
    
            Call Chama_Tela_Modal("CodigoBarras", objProduto)
    
            
            If objProduto.sCodigoBarras <> "Cancel" Then
                If objProduto.lErro = SUCESSO Then
    
                    lErro = CF("INV_Trata_CodigoBarras", objProduto)
                    If lErro <> SUCESSO Then gError 210877
    
                End If
    
                'Lê os demais atributos do Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 210878
    
                'Se não encontrou o Produto --> Erro
                If lErro = 28030 Then gError 210879
    
                lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
                If lErro <> SUCESSO Then gError 210880
        
                Me.Show
        
                Produto.PromptInclude = False
                Produto.Text = sProdutoEnxuto
                Produto.PromptInclude = True
                
                Produto.SetFocus
                
            End If
            
    
    
    Trata_CodigoBarras1 = SUCESSO

    Exit Function

Erro_Trata_CodigoBarras1:

    Trata_CodigoBarras1 = gErr


    Select Case gErr

        Case 210877 To 210878

        Case 210879
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 210880
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210881)

    End Select

    Exit Function

End Function

