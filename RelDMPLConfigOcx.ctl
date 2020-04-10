VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl RelDMPLConfigOcx 
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
   KeyPreview      =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   7770
   Begin VB.Frame TipoElemento 
      Caption         =   "Elemento do Tipo Contas"
      Height          =   2235
      Index           =   0
      Left            =   165
      TabIndex        =   8
      Top             =   3585
      Visible         =   0   'False
      Width           =   7485
      Begin VB.CommandButton BotaoConta 
         Caption         =   "Plano de Contas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   5955
         TabIndex        =   34
         Top             =   420
         Width           =   1380
      End
      Begin MSMask.MaskEdBox ContaFim 
         Height          =   225
         Left            =   3390
         TabIndex        =   9
         Top             =   1245
         Width           =   2220
         _ExtentX        =   3916
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
         Left            =   1050
         TabIndex        =   10
         Top             =   1245
         Width           =   2220
         _ExtentX        =   3916
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
         Height          =   1635
         Left            =   195
         TabIndex        =   11
         Top             =   450
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   2884
         _Version        =   393216
      End
      Begin MSComctlLib.TreeView TvwContas 
         Height          =   1635
         Left            =   4335
         TabIndex        =   12
         Top             =   435
         Visible         =   0   'False
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   2884
         _Version        =   393217
         Indentation     =   453
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
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
      Begin VB.Label LabelContas 
         AutoSize        =   -1  'True
         Caption         =   "Plano de Contas"
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
         Left            =   4350
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Grupos de Contas Associadas"
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
         Left            =   255
         TabIndex        =   24
         Top             =   255
         Width           =   2550
      End
   End
   Begin VB.Frame TipoElemento 
      Caption         =   "Elemento do Tipo Fórmula"
      Height          =   2235
      Index           =   1
      Left            =   180
      TabIndex        =   15
      Top             =   3585
      Visible         =   0   'False
      Width           =   7485
      Begin VB.ListBox ListaFormula 
         Height          =   1620
         Left            =   4560
         TabIndex        =   21
         Top             =   450
         Width           =   2745
      End
      Begin VB.ComboBox SomaSubtrai 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "RelDMPLConfigOcx.ctx":0000
         Left            =   2835
         List            =   "RelDMPLConfigOcx.ctx":000A
         TabIndex        =   17
         Top             =   870
         Width           =   1065
      End
      Begin VB.TextBox Formula 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   570
         MaxLength       =   255
         TabIndex        =   16
         Top             =   750
         Width           =   2175
      End
      Begin MSFlexGridLib.MSFlexGrid GridFormulas 
         Height          =   1665
         Left            =   315
         TabIndex        =   18
         Top             =   435
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   2937
         _Version        =   393216
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Células Disponíveis"
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
         Left            =   4560
         TabIndex        =   26
         Top             =   255
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Fórmula"
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
         Left            =   330
         TabIndex        =   27
         Top             =   255
         Width           =   2025
      End
   End
   Begin VB.Frame FrameExercicio 
      Caption         =   "Exercício"
      Height          =   630
      Left            =   5040
      TabIndex        =   6
      Top             =   2850
      Visible         =   0   'False
      Width           =   2610
      Begin VB.OptionButton BotaoExercAnt 
         Caption         =   "Anterior"
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
         Left            =   1440
         TabIndex        =   13
         Top             =   270
         Width           =   1005
      End
      Begin VB.OptionButton BotaoExercAtual 
         Caption         =   "Atual"
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
         Left            =   345
         TabIndex        =   14
         Top             =   270
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin VB.Frame TipoElemento 
      Caption         =   "Elemento do Tipo Título"
      Height          =   2235
      Index           =   2
      Left            =   180
      TabIndex        =   19
      Top             =   3585
      Visible         =   0   'False
      Width           =   7485
      Begin VB.TextBox Titulo 
         Height          =   345
         Left            =   945
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   390
         Width           =   4995
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Título:"
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
         Left            =   270
         TabIndex        =   25
         Top             =   450
         Width           =   615
      End
   End
   Begin VB.ComboBox ComboModelos 
      Height          =   315
      Left            =   930
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   285
      Width           =   3135
   End
   Begin VB.Frame Tipos 
      Caption         =   "Tipo do Elemento"
      Height          =   615
      Left            =   195
      TabIndex        =   1
      Top             =   2850
      Width           =   4620
      Begin VB.OptionButton BotaoVazio 
         Caption         =   "Vazio"
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
         Left            =   3600
         TabIndex        =   22
         Top             =   270
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton BotaoFormula 
         Caption         =   "Fórmula"
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
         Left            =   1335
         TabIndex        =   4
         Top             =   285
         Width           =   1020
      End
      Begin VB.OptionButton BotaoContas 
         Caption         =   "Contas"
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
         Left            =   210
         TabIndex        =   3
         Top             =   285
         Width           =   1035
      End
      Begin VB.OptionButton BotaoTitulo 
         Caption         =   "Título"
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
         Left            =   2505
         TabIndex        =   2
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5550
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1575
         Picture         =   "RelDMPLConfigOcx.ctx":001D
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "RelDMPLConfigOcx.ctx":019B
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "RelDMPLConfigOcx.ctx":06CD
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "RelDMPLConfigOcx.ctx":0857
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridRelatorio 
      Height          =   1845
      Left            =   210
      TabIndex        =   7
      Top             =   915
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   3254
      _Version        =   393216
      Rows            =   29
      Cols            =   21
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
      AllowUserResizing=   3
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Layout do Relatório"
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
      Left            =   240
      TabIndex        =   28
      Top             =   705
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   180
      TabIndex        =   29
      Top             =   285
      Width           =   690
   End
End
Attribute VB_Name = "RelDMPLConfigOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim gcolRel As Collection
Dim gcolRelConta As Collection
Dim gcolRelFormula As Collection

Dim objGridFormulas As AdmGrid
Dim objGridContas As AdmGrid

Dim gsModelo As String
Dim gsRelatorio As String

Dim iGrid_Formula_Col As Integer
Dim iGrid_Operacao_Col As Integer
Dim iGrid_ContaInicio_Col As Integer
Dim iGrid_ContaFinal_Col As Integer

Private WithEvents objEventoConta As AdmEvento
Attribute objEventoConta.VB_VarHelpID = -1

Const CEL_TIPO_CONTA = 0
Const CEL_TIPO_FORMULA = 1
Const CEL_TIPO_TITULO = 2
Const CEL_TIPO_VAZIO = 3

Const REL_OPERACAO_SOMA = 0
Const REL_OPERACAO_SUBTRAI = 1

Const CELULA_USADO_EM_FORMULA = 1

Const CELULA_FORMULA As String = "<Fórmula>"
Const CELULA_CONTA As String = "<Conta>"

Const CONTAS_EXERCICIO_ATUAL = 0
Const CONTAS_EXERCICIO_ANTERIOR = 1

Dim giEntraCelula As Integer
Dim giCelFixaCol As Integer
Dim giCelFixaLin As Integer
Dim giTipoAtual As Integer
Dim giLinhaAntiga As Integer
Dim giColunaAntiga As Integer



Public Function Trata_Parametros(sRelatorio As String, sTitulo As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colModelos As New Collection

On Error GoTo Erro_Trata_Parametros

    Me.Caption = sTitulo
    gsRelatorio = sRelatorio
    
    lErro = CF("RelDMPL_Le_Modelos_Distintos", gsRelatorio, colModelos)
    If lErro <> SUCESSO Then Error 60605

    For iIndice = 1 To colModelos.Count
        ComboModelos.AddItem colModelos.Item(iIndice)
    Next

    Trata_Parametros = SUCESSO

    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = Err
     
    Select Case Err
          
        Case 60605
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166675)
     
    End Select
     
    Exit Function

End Function

Private Sub BotaoContas_Click()

Dim lErro As Long
Dim objConta As Node

On Error GoTo Erro_BotaoContas_Click

    'Coloca o Frame de Contas visível e o de Fórmulas invisível
    TipoElemento(CEL_TIPO_CONTA).Visible = True
    TipoElemento(CEL_TIPO_FORMULA).Visible = False
    TipoElemento(CEL_TIPO_TITULO).Visible = False
    FrameExercicio.Visible = True

    lErro = Preenche_GridContas()
    If lErro <> SUCESSO Then Error 60529

    giTipoAtual = CEL_TIPO_CONTA

    Exit Sub

Erro_BotaoContas_Click:

    Select Case Err

        Case 60529

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166676)

     End Select

     Exit Sub

End Sub

Private Sub BotaoFormula_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFormula_Click

    TipoElemento(CEL_TIPO_CONTA).Visible = False
    TipoElemento(CEL_TIPO_FORMULA).Visible = True
    TipoElemento(CEL_TIPO_TITULO).Visible = False
    FrameExercicio.Visible = False

    Call Obtem_ListaFormulas

    lErro = Preenche_GridFormulas()
    If lErro <> SUCESSO Then Error 60532

    giTipoAtual = CEL_TIPO_FORMULA

    Exit Sub

Erro_BotaoFormula_Click:

    Select Case Err

        Case 60532

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166677)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTitulo_Click()
    
    TipoElemento(CEL_TIPO_CONTA).Visible = False
    TipoElemento(CEL_TIPO_FORMULA).Visible = False
    TipoElemento(CEL_TIPO_TITULO).Visible = True
    FrameExercicio.Visible = False
    giTipoAtual = CEL_TIPO_TITULO
        
    Call Preenche_Titulo
    
End Sub

Private Sub BotaoVazio_Click()
    
    TipoElemento(CEL_TIPO_CONTA).Visible = False
    TipoElemento(CEL_TIPO_FORMULA).Visible = False
    TipoElemento(CEL_TIPO_TITULO).Visible = False
    FrameExercicio.Visible = False
    giTipoAtual = CEL_TIPO_VAZIO

End Sub

Private Function Preenche_GridContas() As Long
'Preenche o GridContas com os dados da coleção colRelConta

Dim lErro As Long
Dim iLinhas As Integer
Dim objRelConta As ClassRelDMPLConta
Dim objRel As New ClassRelDMPL
Dim sContaMascarada As String
Dim iLinhaAtual As Integer
Dim iColunaAtual As Integer

On Error GoTo Erro_Preenche_GridContas

    'Limpa o grid
    Call Grid_Limpa(objGridContas)

    If giCelFixaCol = -1 Then
    
        iLinhaAtual = GridRelatorio.Row
        iColunaAtual = GridRelatorio.Col
    Else
        iLinhaAtual = giCelFixaLin
        iColunaAtual = giCelFixaCol
    End If

    For Each objRel In gcolRel

        If objRel.iLinha = iLinhaAtual And objRel.iColuna = iColunaAtual Then

            For Each objRelConta In gcolRelConta
        
                If objRel.iLinha = objRelConta.iLinha And objRel.iColuna = objRelConta.iColuna Then
        
                    iLinhas = iLinhas + 1
        
                    If Len(objRelConta.sContaInicial) > 0 Then
        
                        'mascara a conta
                        sContaMascarada = String(STRING_CONTA, 0)
                
                        lErro = Mascara_RetornaContaEnxuta(objRelConta.sContaInicial, sContaMascarada)
                        If lErro <> SUCESSO Then Error 60533
                
                        ContaInicio.PromptInclude = False
                        ContaInicio.Text = sContaMascarada
                        ContaInicio.PromptInclude = True
                        
                        GridContas.TextMatrix(iLinhas, iGrid_ContaInicio_Col) = ContaInicio.Text
        
                    End If
        
                    If Len(objRelConta.sContaFinal) > 0 Then
        
                        'mascara a conta
                        sContaMascarada = String(STRING_CONTA, 0)
                
                        lErro = Mascara_RetornaContaEnxuta(objRelConta.sContaFinal, sContaMascarada)
                        If lErro <> SUCESSO Then Error 60534
                
                        ContaFim.PromptInclude = False
                        ContaFim.Text = sContaMascarada
                        ContaFim.PromptInclude = True
        
                        GridContas.TextMatrix(iLinhas, iGrid_ContaFinal_Col) = ContaFim.Text
        
                    End If
        
                End If
        
            Next
    
            objGridContas.iLinhasExistentes = iLinhas
            
            If objRel.iExercicio = CONTAS_EXERCICIO_ATUAL Then
                BotaoExercAtual.Value = MARCADO
            Else
                BotaoExercAnt.Value = MARCADO
            End If
            
            Exit For
        
        End If
        
    Next

    Preenche_GridContas = SUCESSO

    Exit Function

Erro_Preenche_GridContas:

    Preenche_GridContas = Err

    Select Case Err

        Case 60533
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objRelConta.sContaInicial)
        
        Case 60534
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objRelConta.sContaFinal)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166678)

    End Select

    Exit Function

End Function

Private Sub Obtem_ListaFormulas()
'obtem as células que estão posicionados acima da celula em questao e coloca seus titulos na lista de formulas

Dim objNodePai As Node
Dim objNodeIrmao As Node
Dim lErro As Long
Dim objRel As ClassRelDMPL
Dim iLinhaAtual As Integer
Dim iColunaAtual As Integer

On Error GoTo Erro_Obtem_ListaFormulas

    ListaFormula.Clear

    If giCelFixaCol = -1 Then
    
        iLinhaAtual = GridRelatorio.Row
        iColunaAtual = GridRelatorio.Col
    Else
        iLinhaAtual = giCelFixaLin
        iColunaAtual = giCelFixaCol
    End If

    For Each objRel In gcolRel
    
        If objRel.iLinha < iLinhaAtual Or (objRel.iLinha = iLinhaAtual And objRel.iColuna < iColunaAtual) And (objRel.iTipo = CEL_TIPO_CONTA Or objRel.iTipo = CEL_TIPO_FORMULA) Then
        
            If objRel.iTipo = CEL_TIPO_CONTA Or objRel.iTipo = CEL_TIPO_FORMULA Then
                ListaFormula.AddItem "L" & CStr(objRel.iLinha) & "C" & CStr(objRel.iColuna)
            End If
        
        End If

    Next

    Exit Sub

Erro_Obtem_ListaFormulas:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166679)

    End Select

    Exit Sub

End Sub

Private Function Preenche_GridFormulas() As Long
'Preenche o GridFormulas com o conteúdo da coleção colRelFormula

Dim lErro As Long
Dim iIndice As Integer
Dim objRelFormula As ClassRelDMPLFormula
Dim objRel As ClassRelDMPL
Dim iLinha As Integer
Dim iLinhaAtual As Integer
Dim iColunaAtual As Integer

On Error GoTo Erro_Preenche_GridFormulas

    'Limpa o grid
    Call Grid_Limpa(objGridFormulas)

    If giCelFixaCol = -1 Then
    
        iLinhaAtual = GridRelatorio.Row
        iColunaAtual = GridRelatorio.Col
    Else
        iLinhaAtual = giCelFixaLin
        iColunaAtual = giCelFixaCol
    End If

    For Each objRel In gcolRel

        If objRel.iLinha = iLinhaAtual And objRel.iColuna = iColunaAtual Then

            For Each objRelFormula In gcolRelFormula
        
                If objRel.iLinha = objRelFormula.iLinha And objRel.iColuna = objRelFormula.iColuna Then
        
                    iLinha = iLinha + 1
        
                    If objRelFormula.iOperacao = REL_OPERACAO_SOMA Then
        
                        GridFormulas.TextMatrix(iLinha, iGrid_Operacao_Col) = SomaSubtrai.List(REL_OPERACAO_SOMA)
        
                    Else
        
                        GridFormulas.TextMatrix(iLinha, iGrid_Operacao_Col) = SomaSubtrai.List(REL_OPERACAO_SUBTRAI)
        
                    End If
        
                    GridFormulas.TextMatrix(iLinha, iGrid_Formula_Col) = "L" & CStr(objRelFormula.iLinhaFormula) & "C" & CStr(objRelFormula.iColunaFormula)
        
                End If
        
            Next
    
        Exit For
    
        End If

    Next

    objGridFormulas.iLinhasExistentes = iLinha

    Preenche_GridFormulas = SUCESSO

    Exit Function

Erro_Preenche_GridFormulas:

    Preenche_GridFormulas = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166680)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim sModelo As String
Dim vbMsgRes As Long
Dim iIndice As Integer

On Error GoTo Erro_BotaoExcluir_Click

    If Len(ComboModelos.Text) = 0 Then Error 60535

    'Envia Mensagem pedindo confirmação da Exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_MODELORELDRE")

    If vbMsgRes = vbYes Then

        sModelo = ComboModelos.Text

        'Exclui o modelo
        lErro = CF("RelDMPL_Exclui1", gsRelatorio, sModelo)
        If lErro <> SUCESSO Then Error 60536

        For iIndice = 0 To ComboModelos.ListCount - 1
            If ComboModelos.List(iIndice) = sModelo Then
                ComboModelos.RemoveItem iIndice
                Exit For
            End If
        Next
            
        Call Limpa_Tela_RelDMPLConfig

        iAlterado = 0

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 60535
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_INFORMADO", Err)

        Case 60536

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166681)

    End Select

    Exit Sub

End Sub

Function Limpa_Tela_RelDMPLConfig() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_RelDMPLConfig

    ComboModelos.Text = ""
    BotaoVazio.Value = True
    ListaFormula.Clear
    Call Grid_Limpa(objGridContas)
    Call Grid_Limpa(objGridFormulas)
    GridRelatorio.Clear
    Set gcolRel = New Collection
    Set gcolRelConta = New Collection
    Set gcolRelFormula = New Collection

    'Inicializa Grid Relatorio
    lErro = Inicializa_Grid_Relatorio()
    If lErro <> SUCESSO Then Error 60695
    
    Limpa_Tela_RelDMPLConfig = SUCESSO

    Exit Function

Erro_Limpa_Tela_RelDMPLConfig:

    Limpa_Tela_RelDMPLConfig = Err

    Select Case Err

        Case 60695

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166682)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 60562

    Call Limpa_Tela_RelDMPLConfig

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 60562

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166683)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iAchou As Integer
Dim iLinhaAtual As Integer
Dim iColunaAtual As Integer

On Error GoTo Erro_Gravar_Registro

    gsModelo = ComboModelos.Text
    
    If Len(gsModelo) = 0 Then Error 60563

    If giCelFixaCol = -1 Then
    
        iLinhaAtual = GridRelatorio.Row
        iColunaAtual = GridRelatorio.Col
    Else
        iLinhaAtual = giCelFixaLin
        iColunaAtual = giCelFixaCol
    End If

    'Salva dados da celula corrente do grid
    lErro = Move_Tela_Memoria(iLinhaAtual, iColunaAtual)
    If lErro <> SUCESSO Then Error 60630

    lErro = CF("RelDMPL_Grava", gsRelatorio, gsModelo, gcolRel, gcolRelConta, gcolRelFormula)
    If lErro <> SUCESSO Then Error 60631
    
    'verifica se o nome do modelo já está na combo
    For iIndice = 0 To ComboModelos.ListCount - 1
        If ComboModelos.List(iIndice) = gsModelo Then
            iAchou = 1
            Exit For
        End If
    Next
    
    'se não tiver, coloca-a.
    If iAchou = 0 Then ComboModelos.AddItem gsModelo

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    Select Case Err
    
        Case 60563
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_INFORMADO", Err)

        Case 60630, 60631

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166684)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(iLinha As Integer, iColuna As Integer) As Long

Dim lErro As Long
Dim iTipo As Integer

On Error GoTo Erro_Move_Tela_Memoria

    If BotaoContas.Value = True Then
        iTipo = CEL_TIPO_CONTA
    ElseIf BotaoFormula.Value = True Then
        iTipo = CEL_TIPO_FORMULA
    ElseIf BotaoTitulo.Value = True Then
        iTipo = CEL_TIPO_TITULO
    ElseIf BotaoVazio.Value = True Then
        iTipo = CEL_TIPO_VAZIO
    End If

    lErro = Move_Grid_Memoria(iTipo, iLinha, iColuna)
    If lErro <> SUCESSO Then Error 60564

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 60564
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166685)

    End Select

    Exit Function

End Function

Private Function Move_Grid_Memoria(iTipo As Integer, iLinha As Integer, iColuna As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Grid_Memoria

    Select Case iTipo

        Case CEL_TIPO_CONTA

            lErro = Move_Contas_Memoria(iLinha, iColuna)
            If lErro <> SUCESSO Then Error 60565

        Case CEL_TIPO_FORMULA

            lErro = Move_Formulas_Memoria(iLinha, iColuna)
            If lErro <> SUCESSO Then Error 60566
            
        Case CEL_TIPO_TITULO
        
            lErro = Move_Titulo_Memoria(iLinha, iColuna)
            If lErro <> SUCESSO Then Error 60567

        Case CEL_TIPO_VAZIO
            lErro = Move_Vazio_Memoria(iLinha, iColuna)
            If lErro <> SUCESSO Then Error 60575

    End Select

    Move_Grid_Memoria = SUCESSO

    Exit Function

Erro_Move_Grid_Memoria:

    Move_Grid_Memoria = Err

    Select Case Err

        Case 60565, 60566, 60567, 60575
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166686)

    End Select

    Exit Function

End Function

Private Function Move_Contas_Memoria(iLinha As Integer, iColuna As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sContaInicio As String
Dim sContaFim As String
Dim sContaFormatada As String, sContaFormatada1 As String
Dim iContaPreenchida As Integer
Dim objRelConta As ClassRelDMPLConta
Dim colRelContaNovo As New Collection
Dim objRelFormula As ClassRelDMPLFormula
Dim colRelFormulaNovo As New Collection
Dim objRel As ClassRelDMPL
Dim iAchou As Integer

On Error GoTo Erro_Move_Contas_Memoria

    'Remove todos os componentes antigos do elemento tipo formula
    'Pois este elemento poderia ter sido uma formula e ter sido transformado em conta
    For Each objRelFormula In gcolRelFormula

        If objRelFormula.iLinha <> iLinha Or objRelFormula.iColuna <> iColuna Then colRelFormulaNovo.Add objRelFormula

    Next

    Set gcolRelFormula = colRelFormulaNovo

    'Remove todos os componentes antigos do elemento tipo conta
    For Each objRelConta In gcolRelConta

        If objRelConta.iLinha <> iLinha Or objRelConta.iColuna <> iColuna Then colRelContaNovo.Add objRelConta

    Next

    Set gcolRelConta = colRelContaNovo

    'Adiciona os novos elementos
    For iIndice = 1 To objGridContas.iLinhasExistentes

        sContaInicio = GridContas.TextMatrix(iIndice, iGrid_ContaInicio_Col)
        sContaFim = GridContas.TextMatrix(iIndice, iGrid_ContaFinal_Col)
    
        lErro = CF("Conta_Formata", sContaInicio, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 60568
        
        If iContaPreenchida = CONTA_VAZIA Then Error 60569
    
        lErro = CF("Conta_Formata", sContaFim, sContaFormatada1, iContaPreenchida)
        If lErro <> SUCESSO Then Error 60570
    
        If iContaPreenchida = CONTA_VAZIA Then Error 60571
    
        Set objRelConta = New ClassRelDMPLConta
    
        objRelConta.iLinha = iLinha
        objRelConta.iColuna = iColuna
        objRelConta.sContaInicial = sContaFormatada
        objRelConta.sContaFinal = sContaFormatada1
        objRelConta.iItem = iIndice
    
        If objRelConta.sContaInicial > objRelConta.sContaFinal Then Error 60694
    
        gcolRelConta.Add objRelConta

    Next

    iAchou = 0

    For Each objRel In gcolRel
        If objRel.iLinha = iLinha And objRel.iColuna = iColuna Then
            iAchou = 1
            Exit For
        End If
    Next

    'Se não encontrou o elemento na colecao
    If iAchou = 0 Then
        'adiciona um novo elemento a coleção
        Set objRel = New ClassRelDMPL
        gcolRel.Add objRel
        
        objRel.iLinha = iLinha
        objRel.iColuna = iColuna
    End If
        
    objRel.iTipo = CEL_TIPO_CONTA
    If BotaoExercAnt.Value = True Then
        objRel.iExercicio = CONTAS_EXERCICIO_ANTERIOR
    Else
        objRel.iExercicio = CONTAS_EXERCICIO_ATUAL
    End If

    GridRelatorio.TextMatrix(iLinha, iColuna) = CELULA_CONTA

    Move_Contas_Memoria = SUCESSO

    Exit Function

Erro_Move_Contas_Memoria:

    Move_Contas_Memoria = Err

    Select Case Err

        Case 60568, 60570
        
        Case 60569
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INICIO_NAO_PREENCHIDA", Err, iIndice)
            
        Case 60571
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_FIM_NAO_PREENCHIDA", Err, iIndice)
            
        Case 60694
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTAFIM_MENOR_CONTAINICIO", Err, iIndice, objRelConta.sContaFinal, objRelConta.sContaInicial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166687)

    End Select

    Exit Function

End Function

Private Function Move_Formulas_Memoria(iLinha As Integer, iColuna As Integer) As Long

Dim lErro As Long, iAchou As Integer, iOperacao As Integer
Dim iIndice As Integer, iIndice1 As Integer, iIndice2 As Integer
Dim sFormula As String, iFormula As Integer
Dim sOperacao As String, iOperacaoGrid As Integer
Dim objRelFormula As ClassRelDMPLFormula
Dim colRelFormulaNovo As New Collection
Dim objRelConta As ClassRelDMPLConta
Dim colRelContaNovo As New Collection
Dim objRel As ClassRelDMPL

On Error GoTo Erro_Move_Formulas_Memoria

    'Remove todos os componentes antigos do elemento tipo conta
    'Pois este elemento poderia ter sido uma conta e ter sido transformado em formula
    For Each objRelFormula In gcolRelFormula

        If objRelFormula.iLinha <> iLinha Or objRelFormula.iColuna <> iColuna Then colRelFormulaNovo.Add objRelFormula

    Next

    Set gcolRelFormula = colRelFormulaNovo

    'Remove todos os componentes antigos do elemento tipo conta
    For Each objRelConta In gcolRelConta

        If objRelConta.iLinha <> iLinha Or objRelConta.iColuna <> iColuna Then colRelContaNovo.Add objRelConta

    Next

    Set gcolRelConta = colRelContaNovo

    For iIndice = 1 To objGridFormulas.iLinhasExistentes

        sFormula = GridFormulas.TextMatrix(iIndice, iGrid_Formula_Col)
        sOperacao = GridFormulas.TextMatrix(iIndice, iGrid_Operacao_Col)

        If Len(Trim(sFormula)) = 0 Then Error 60572

        If Len(Trim(sOperacao)) = 0 And iIndice < objGridFormulas.iLinhasExistentes Then Error 60573


        For iIndice1 = 0 To SomaSubtrai.ListCount - 1
            If SomaSubtrai.List(iIndice1) = sOperacao Then
                iOperacaoGrid = SomaSubtrai.ItemData(iIndice1)
                Exit For
            End If
        Next

        Set objRelFormula = New ClassRelDMPLFormula

        objRelFormula.iLinha = iLinha
        objRelFormula.iColuna = iColuna
        objRelFormula.iOperacao = iOperacaoGrid
        objRelFormula.iItem = iIndice
        objRelFormula.iLinhaFormula = CInt(Mid(sFormula, 2, InStr(sFormula, "C") - 2))
        objRelFormula.iColunaFormula = CInt(Mid(sFormula, InStr(sFormula, "C") + 1))

        gcolRelFormula.Add objRelFormula

    Next

    iAchou = 0

    For Each objRel In gcolRel
        If objRel.iLinha = iLinha And objRel.iColuna = iColuna Then
            iAchou = 1
            Exit For
        End If
    Next

    'Se não encontrou o elemento na colecao
    If iAchou = 0 Then
        'adiciona um novo elemento a coleção
        Set objRel = New ClassRelDMPL
        gcolRel.Add objRel
        
        objRel.iLinha = iLinha
        objRel.iColuna = iColuna
    End If
        
    objRel.iTipo = CEL_TIPO_FORMULA

    GridRelatorio.TextMatrix(iLinha, iColuna) = CELULA_FORMULA
    
    Move_Formulas_Memoria = SUCESSO

    Exit Function

Erro_Move_Formulas_Memoria:

    Move_Formulas_Memoria = Err

    Select Case Err

        Case 60572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMULA_NAO_PREENCHIDA", Err, iIndice)

        Case 60573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_NAO_PREENCHIDO", Err, iIndice)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166688)

    End Select

    Exit Function

End Function

Private Function Move_Titulo_Memoria(iLinha As Integer, iColuna As Integer) As Long

Dim lErro As Long
Dim objRelConta As ClassRelDMPLConta
Dim colRelContaNovo As New Collection
Dim objRelFormula As ClassRelDMPLFormula
Dim colRelFormulaNovo As New Collection
Dim objRel As ClassRelDMPL
Dim iAchou As Integer

On Error GoTo Erro_Move_Titulo_Memoria

    'Remove todos os componentes antigos do elemento tipo formula
    'Pois este elemento poderia ter sido uma formula e ter sido transformado em titulo
    For Each objRelFormula In gcolRelFormula

        If objRelFormula.iLinha <> iLinha Or objRelFormula.iColuna <> iColuna Then colRelFormulaNovo.Add objRelFormula

    Next

    Set gcolRelFormula = colRelFormulaNovo

    'Remove todos os componentes antigos do elemento tipo conta
    For Each objRelConta In gcolRelConta

        If objRelConta.iLinha <> iLinha Or objRelConta.iColuna <> iColuna Then colRelContaNovo.Add objRelConta

    Next

    Set gcolRelConta = colRelContaNovo

    iAchou = 0

    For Each objRel In gcolRel
        If objRel.iLinha = iLinha And objRel.iColuna = iColuna Then
            iAchou = 1
            Exit For
        End If
    Next

    'Se não encontrou o elemento na colecao
    If iAchou = 0 Then
        'adiciona um novo elemento a coleção
        Set objRel = New ClassRelDMPL
        gcolRel.Add objRel
        
        objRel.iLinha = iLinha
        objRel.iColuna = iColuna
    End If
        
    objRel.sTitulo = Titulo.Text
    objRel.iTipo = CEL_TIPO_TITULO
    
    GridRelatorio.TextMatrix(iLinha, iColuna) = Titulo.Text
    
    Move_Titulo_Memoria = SUCESSO

    Exit Function

Erro_Move_Titulo_Memoria:

    Move_Titulo_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166689)

    End Select

    Exit Function

End Function

Private Function Move_Vazio_Memoria(iLinha As Integer, iColuna As Integer) As Long

Dim lErro As Long
Dim objRelConta As ClassRelDMPLConta
Dim colRelContaNovo As New Collection
Dim objRelFormula As ClassRelDMPLFormula
Dim colRelFormulaNovo As New Collection
Dim iIndice As Integer
Dim objRel As ClassRelDMPL

On Error GoTo Erro_Move_Vazio_Memoria

    'Remove todos os componentes antigos do elemento tipo formula
    'Pois este elemento poderia ter sido uma formula e ter sido transformado em titulo
    For Each objRelFormula In gcolRelFormula

        If objRelFormula.iLinha <> iLinha Or objRelFormula.iColuna <> iColuna Then colRelFormulaNovo.Add objRelFormula

    Next

    Set gcolRelFormula = colRelFormulaNovo

    'Remove todos os componentes antigos do elemento tipo conta
    For Each objRelConta In gcolRelConta

        If objRelConta.iLinha <> iLinha Or objRelConta.iColuna <> iColuna Then colRelContaNovo.Add objRelConta

    Next

    Set gcolRelConta = colRelContaNovo

    For iIndice = 1 To gcolRel.Count
        Set objRel = gcolRel.Item(iIndice)
        If objRel.iLinha = iLinha And objRel.iColuna = iColuna Then
            gcolRel.Remove (iIndice)
            Exit For
        End If
    Next

    GridRelatorio.TextMatrix(iLinha, iColuna) = ""

    Move_Vazio_Memoria = SUCESSO

    Exit Function

Erro_Move_Vazio_Memoria:

    Move_Vazio_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166690)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 60576

    Call Limpa_Tela_RelDMPLConfig

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 60576

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166691)

    End Select

    Exit Sub

End Sub

Private Function Testa_Uso_Celula(iLinha As Integer, iColuna As Integer, iStatus As Integer) As Long
'verifica se a celula representada pela linha/coluna em questão está sendo usada em alguma formula. Se usar ==> erro

Dim lErro As Long
Dim objRelFormula As ClassRelDMPLFormula
Dim objRel As ClassRelDMPL
Dim objRel1 As ClassRelDMPL

On Error GoTo Erro_Testa_Uso_Celula

    For Each objRelFormula In gcolRelFormula
    
        If objRelFormula.iLinhaFormula = iLinha And objRelFormula.iColunaFormula = iColuna Then Error 60577

    Next

    Testa_Uso_Celula = SUCESSO
    
    Exit Function

Erro_Testa_Uso_Celula:

    Testa_Uso_Celula = Err

    Select Case Err

        Case 60577
            Call Rotina_Erro(vbOKOnly, "ERRO_CEL_UTILIZA_CEL_EM_FORMULA", Err, objRel.iLinha, objRel.iColuna, iLinha, iColuna)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166692)

    End Select
    
    Exit Function
    
End Function

Private Sub ComboModelos_Change()

        iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboModelos_Click()

Dim lErro As Long

On Error GoTo Erro_ComboModelos_Click

    If ComboModelos.ListIndex = -1 Then Exit Sub

    'verifica se existe a necessidade de salvar o modelo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 60596

    gsModelo = ComboModelos.Text

    Set gcolRel = New Collection
    Set gcolRelConta = New Collection
    Set gcolRelFormula = New Collection

    lErro = CF("RelDMPL_Le_Modelo", gsRelatorio, gsModelo, gcolRel)
    If lErro <> SUCESSO Then Error 60593

    lErro = CF("RelDMPLConta_Le_Modelo", gsRelatorio, gsModelo, gcolRelConta)
    If lErro <> SUCESSO Then Error 60594

    lErro = CF("RelDMPLFormula_Le_Modelo", gsRelatorio, gsModelo, gcolRelFormula)
    If lErro <> SUCESSO Then Error 60595

    lErro = Preenche_Modelo_Tela(gcolRel)
    If lErro <> SUCESSO Then Error 60597

    iAlterado = 0

    Exit Sub

Erro_ComboModelos_Click:

    Select Case Err

        Case 60593, 60594, 60595, 60596, 60597

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166693)

    End Select

    Exit Sub

End Sub

Private Function Preenche_Modelo_Tela(colRel As Collection) As Long
'carrega o grid

Dim lErro As Long
Dim objRel As ClassRelDMPL
Dim objRelAnterior As ClassRelDMPL

On Error GoTo Erro_Preenche_Modelo_Tela

    GridRelatorio.Clear

    'Inicializa Grid Relatorio
    lErro = Inicializa_Grid_Relatorio()
    If lErro <> SUCESSO Then Error 60696


    For Each objRel In colRel

        If objRel.iTipo = CEL_TIPO_FORMULA Then GridRelatorio.TextMatrix(objRel.iLinha, objRel.iColuna) = CELULA_FORMULA
        If objRel.iTipo = CEL_TIPO_CONTA Then GridRelatorio.TextMatrix(objRel.iLinha, objRel.iColuna) = CELULA_CONTA
        If objRel.iTipo = CEL_TIPO_TITULO Then GridRelatorio.TextMatrix(objRel.iLinha, objRel.iColuna) = objRel.sTitulo

    Next

    Call GridRelatorio_EnterCell

    Preenche_Modelo_Tela = SUCESSO

    Exit Function

Erro_Preenche_Modelo_Tela:

    Preenche_Modelo_Tela = Err

    Select Case Err

        Case 60696

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166694)

    End Select

    Exit Function

End Function

Private Sub ComboModelos_Validate(Cancel As Boolean)
'Trata saida da combo de modelos

Dim lErro As Long
Dim iIndice As Integer, iCodigo As Integer
Dim iAchou As Integer

On Error GoTo Erro_ComboModelos_Validate

    If Len(gsModelo) = 0 Then Exit Sub
    
    If ComboModelos.Text <> gsModelo Then

        gsModelo = ComboModelos.Text

    End If

    Exit Sub

Erro_ComboModelos_Validate:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166695)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set gcolRel = New Collection
    Set gcolRelConta = New Collection
    Set gcolRelFormula = New Collection
    Set objGridFormulas = New AdmGrid
    Set objGridContas = New AdmGrid

    'Inicializa Grid Relatorio
    lErro = Inicializa_Grid_Relatorio()
    If lErro <> SUCESSO Then Error 60693

    'Inicializa Grid de Formulas
    lErro = Inicializa_Grid_Formulas(objGridFormulas)
    If lErro <> SUCESSO Then Error 60600

    'Inicializa Grid de Contas
    lErro = Inicializa_Grid_Contas(objGridContas)
    If lErro <> SUCESSO Then Error 60601

   'inicializa a mascara de conta
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaInicio)
    If lErro <> SUCESSO Then Error 60602

    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaFim)
    If lErro <> SUCESSO Then Error 60603

'    'Inicializa a Lista de Plano de Contas
'    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
'    If lErro <> SUCESSO Then Error 60604

    Set objEventoConta = New AdmEvento

    BotaoVazio.Value = MARCADO
    giCelFixaCol = -1
    giCelFixaLin = -1
    giTipoAtual = -1
    iAlterado = 0
    giLinhaAntiga = -1
    giColunaAntiga = -1
    giEntraCelula = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = SUCESSO

    Select Case Err

        Case 60600, 60601, 60602, 60603, 60604, 60693

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166696)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_Relatorio() As Long

Dim iIndice As Integer

    'Largura da primeira coluna
    GridRelatorio.ColWidth(0) = 500
    
    For iIndice = 0 To GridRelatorio.Cols - 1
        GridRelatorio.ColAlignment(iIndice) = flexAlignCenterCenter
    Next
    
    For iIndice = 1 To GridRelatorio.Rows - 1
        GridRelatorio.TextMatrix(iIndice, 0) = "L" & CStr(iIndice)
    Next
    
    For iIndice = 1 To GridRelatorio.Cols - 1
        GridRelatorio.TextMatrix(0, iIndice) = "C" & CStr(iIndice)
    Next

    GridRelatorio.FocusRect = flexFocusNone
    GridRelatorio.ForeColorSel = vbHighlightText

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridFormulas.Name

                lErro = Saida_Celula_Formulas(objGridInt)
                If lErro <> SUCESSO Then Error 60611

            'Se for o GridDescontos
            Case GridContas.Name

                lErro = Saida_Celula_Contas(objGridInt)
                If lErro <> SUCESSO Then Error 60612

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 60613

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 60611, 60612

        Case 60613
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166697)

    End Select

    Exit Function

End Function

'********************************
'Funções relativas ao GridFormulas
'********************************

Private Function Inicializa_Grid_Formulas(objGridInt As AdmGrid) As Long
'Inicializa o Grid de fórmulas
Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Formulas

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Formula")
    objGridInt.colColuna.Add ("Soma/Subtrai")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Formula.Name)
    objGridInt.colCampo.Add (SomaSubtrai.Name)

    'Colunas do Grid
    iGrid_Formula_Col = 1
    iGrid_Operacao_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridFormulas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_FORMULAS_DRE + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridFormulas.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Formulas = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Formulas:

    Inicializa_Grid_Formulas = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166698)

    End Select

    Exit Function

End Function

Private Sub GridFormulas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridFormulas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGridFormulas, iAlterado)

    End If

End Sub

Private Sub GridFormulas_EnterCell()

    Call Grid_Entrada_Celula(objGridFormulas, iAlterado)

End Sub

Private Sub GridFormulas_GotFocus()

    Call Grid_Recebe_Foco(objGridFormulas)

End Sub

Private Sub GridFormulas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFormulas)

End Sub

Private Sub GridFormulas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFormulas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGridFormulas, iAlterado)

    End If

End Sub

Private Sub GridFormulas_LeaveCell()

    Call Saida_Celula(objGridFormulas)

End Sub

Private Sub GridFormulas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridFormulas)

End Sub

Private Sub GridFormulas_RowColChange()

    Call Grid_RowColChange(objGridFormulas)

End Sub

Private Sub GridFormulas_Scroll()

    Call Grid_Scroll(objGridFormulas)

End Sub

Public Function Saida_Celula_Formulas(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid de fórmulas que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Formulas

    If objGridInt.objGrid Is GridFormulas Then

        Select Case GridFormulas.Col

            Case iGrid_Formula_Col

                lErro = Saida_Celula_Formula(objGridInt)
                If lErro <> SUCESSO Then Error 60612

            Case iGrid_Operacao_Col

                lErro = Saida_Celula_SomaSubtrai(objGridInt)
                If lErro <> SUCESSO Then Error 60613

        End Select

    End If

    Saida_Celula_Formulas = SUCESSO

    Exit Function

Erro_Saida_Celula_Formulas:

    Saida_Celula_Formulas = Err

    Select Case Err

        Case 60612, 60613

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166699)

    End Select

    Exit Function

End Function

Private Sub Formula_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Formula_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridFormulas)
End Sub

Private Sub Formula_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFormulas)
End Sub

Private Sub Formula_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFormulas.objControle = Formula
    lErro = Grid_Campo_Libera_Foco(objGridFormulas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub SomaSubtrai_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SomaSubtrai_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridFormulas)
End Sub

Private Sub SomaSubtrai_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFormulas)
End Sub

Private Sub SomaSubtrai_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFormulas.objControle = SomaSubtrai
    lErro = Grid_Campo_Libera_Foco(objGridFormulas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Function Saida_Celula_Formula(objGridInt As AdmGrid) As Long
'faz a critica da celula de formula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim sFormula As String
Dim iAchou As Integer

On Error GoTo Erro_Saida_Celula_Formula

    Set objGridInt.objControle = Formula

    If Len(Trim(Formula.Text)) > 0 Then

        sFormula = Formula.Text
        iAchou = 0

        For iIndice = 0 To ListaFormula.ListCount - 1
            If sFormula = ListaFormula.List(iIndice) Then
                iAchou = 1
                Exit For
            End If
        Next

        If iAchou = 0 Then Error 60614

        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridFormulas.Row - GridFormulas.FixedRows) = objGridFormulas.iLinhasExistentes Then
            objGridFormulas.iLinhasExistentes = objGridFormulas.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 60615

    Saida_Celula_Formula = SUCESSO

    Exit Function

Erro_Saida_Celula_Formula:

    Saida_Celula_Formula = Err

    Select Case Err

        Case 60614
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMULA_INVALIDA1", Err, sFormula)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 60615
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166700)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_SomaSubtrai(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_SomaSubtrai

    Set objGridInt.objControle = SomaSubtrai

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 60616

    Saida_Celula_SomaSubtrai = SUCESSO

    Exit Function

Erro_Saida_Celula_SomaSubtrai:

    Saida_Celula_SomaSubtrai = Err

    Select Case Err

        Case 60616
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166701)

    End Select

    Exit Function

End Function
'********************************
' fim do tratamento do GridFormulas
'********************************

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

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ContaInicio.Name)
    objGridInt.colCampo.Add (ContaFim.Name)

    'Colunas do Grid
    iGrid_ContaInicio_Col = 1
    iGrid_ContaFinal_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridContas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_CONTAS_DRE + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridContas.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Contas = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Contas:

    Inicializa_Grid_Contas = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166702)

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
                If lErro <> SUCESSO Then Error 60617

            Case iGrid_ContaFinal_Col

                lErro = Saida_Celula_ContaFim(objGridInt)
                If lErro <> SUCESSO Then Error 60618

        End Select

    End If

    Saida_Celula_Contas = SUCESSO

    Exit Function

Erro_Saida_Celula_Contas:

    Saida_Celula_Contas = Err

    Select Case Err

        Case 60617, 60618

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166703)

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

On Error GoTo Erro_Saida_Celula_ContaInicio

    Set objGridInt.objControle = ContaInicio

    If Len(Trim(ContaInicio.ClipText)) > 0 Then

        sContaFormatada = String(STRING_CONTA, 0)
        
        lErro = CF("Conta_Formata", ContaInicio.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 60619
    
        If iContaPreenchida = CONTA_PREENCHIDA Then
    
            'verifica se a Conta Final existe
            lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
            If lErro <> SUCESSO And lErro <> 6030 Then Error 60620

            If lErro = 6030 Then Error 60621

        End If

        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridContas.Row - GridContas.FixedRows) = objGridContas.iLinhasExistentes Then
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 60622

    Saida_Celula_ContaInicio = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaInicio:

    Saida_Celula_ContaInicio = Err

    Select Case Err

        Case 60619, 60620, 60622
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 60621
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, ContaInicio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166704)

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
        If lErro <> SUCESSO Then Error 60623
    
        If iContaPreenchida = CONTA_PREENCHIDA Then
    
            'verifica se a Conta Final existe
            lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
            If lErro <> SUCESSO And lErro <> 6030 Then Error 60624

            If lErro = 6030 Then Error 60625

        End If

        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridContas.Row - GridContas.FixedRows) = objGridContas.iLinhasExistentes Then
            objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 60626

    Saida_Celula_ContaFim = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaFim:

    Saida_Celula_ContaFim = Err

    Select Case Err

        Case 60623, 60624, 60626
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 60625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, ContaFim.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166705)

    End Select

    Exit Function

End Function
'********************************
' fim do tratamento do GridContas
'********************************

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gcolRel = Nothing
    Set gcolRelConta = Nothing
    Set gcolRelFormula = Nothing
    Set objGridFormulas = Nothing
    Set objGridContas = Nothing
    
    Set objEventoConta = Nothing

End Sub

Private Sub ListaFormula_DblClick()

Dim lErro As Long

On Error GoTo Erro_ListaFormula_DblClick

    If ListaFormula.ListIndex = -1 Then Exit Sub

    If GridFormulas.Col = iGrid_Formula_Col Then

        Formula.Text = ListaFormula.List(ListaFormula.ListIndex)

        GridFormulas.TextMatrix(GridFormulas.Row, iGrid_Formula_Col) = Formula.Text

        If objGridFormulas.objGrid.Row - objGridFormulas.objGrid.FixedRows = objGridFormulas.iLinhasExistentes Then
            objGridFormulas.iLinhasExistentes = objGridFormulas.iLinhasExistentes + 1
        End If

    End If

    Exit Sub

Erro_ListaFormula_DblClick:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166706)

    End Select

    Exit Sub

End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then

        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 60627

    End If

    Exit Sub

Erro_TvwContas_Expand:

    Select Case Err

        Case 60627

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166707)

    End Select

    Exit Sub

End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sConta As String
Dim sCaracterInicial As String
Dim lErro As Long
Dim sContaEnxuta As String
Dim sContaMascarada As String
Dim iLinha As Integer

On Error GoTo Erro_TvwContas_NodeClick

    sCaracterInicial = left(Node.Key, 1)

    sConta = right(Node.Key, Len(Node.Key) - 1)

    sContaEnxuta = String(STRING_CONTA, 0)

    'volta mascarado apenas os caracteres preenchidos
    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then Error 60628

    If GridContas.Col = iGrid_ContaInicio_Col Then

        ContaInicio.PromptInclude = False
        ContaInicio.Text = sContaEnxuta
        ContaInicio.PromptInclude = True

        GridContas.TextMatrix(GridContas.Row, iGrid_ContaInicio_Col) = ContaInicio.Text

    ElseIf GridContas.Col = iGrid_ContaFinal_Col Then

        ContaFim.PromptInclude = False
        ContaFim.Text = sContaEnxuta
        ContaFim.PromptInclude = True

        GridContas.TextMatrix(GridContas.Row, iGrid_ContaFinal_Col) = ContaFim.Text

    End If

    If objGridContas.objGrid.Row - objGridContas.objGrid.FixedRows = objGridContas.iLinhasExistentes Then
        objGridContas.iLinhasExistentes = objGridContas.iLinhasExistentes + 1
    End If

    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err

        Case 60628
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166708)

    End Select

    Exit Sub

End Sub

Private Sub GridRelatorio_LeaveCell()

Dim lErro As Long

On Error GoTo Erro_GridRelatorio_LeaveCell

    If giLinhaAntiga = -1 And giColunaAntiga = -1 Then

        If giCelFixaCol = -1 Then

            GridRelatorio.CellBackColor = vbWindowBackground
            GridRelatorio.CellForeColor = vbWindowText

            lErro = Move_Tela_Memoria(GridRelatorio.Row, GridRelatorio.Col)
            If lErro <> SUCESSO Then Error 60629
        
        End If
    
    End If
    
    Exit Sub

Erro_GridRelatorio_LeaveCell:

    Select Case Err

        Case 60629
            giLinhaAntiga = GridRelatorio.Row
            giColunaAntiga = GridRelatorio.Col
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166709)

    End Select

    Exit Sub

End Sub

Private Sub GridRelatorio_EnterCell()

Dim lErro As Long
Dim objRel As ClassRelDMPL
Dim iAchou As Integer

On Error GoTo Erro_GridRelatorio_EnterCell
        
    If giEntraCelula = 1 Then Exit Sub
        
    If giLinhaAntiga <> -1 Then
        giEntraCelula = 1
        GridRelatorio.Row = giLinhaAntiga
        GridRelatorio.Col = giColunaAntiga
        giLinhaAntiga = -1
        giColunaAntiga = -1
        giEntraCelula = 0
    Else
    
        If giCelFixaCol = -1 Then
        
            'limpa grids
            Call Grid_Limpa(objGridContas)
            Call Grid_Limpa(objGridFormulas)
            Titulo.Text = ""
        
            For Each objRel In gcolRel
    
                If objRel.iLinha = GridRelatorio.Row And objRel.iColuna = GridRelatorio.Col Then
                    iAchou = 1
                    Exit For
                End If
                
            Next
        
            If iAchou = 1 Then
        
                If objRel.iTipo = CEL_TIPO_CONTA Then
                    BotaoContas.Value = False
                    BotaoContas.Value = True
                ElseIf objRel.iTipo = CEL_TIPO_FORMULA Then
                    BotaoFormula.Value = False
                    BotaoFormula.Value = True
                ElseIf objRel.iTipo = CEL_TIPO_TITULO Then
                    BotaoTitulo.Value = False
                    BotaoTitulo.Value = True
                Else
                    BotaoVazio.Value = False
                    BotaoVazio.Value = True
                End If
                
            Else
                BotaoVazio.Value = False
                BotaoVazio.Value = True
            End If
    
        End If
    
    End If
    
    Exit Sub

Erro_GridRelatorio_EnterCell:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166710)

    End Select

    Exit Sub

End Sub

Private Sub Seta_Tipo(iTipo As Integer)

    Select Case iTipo
    
        Case CEL_TIPO_CONTA
            BotaoContas.Value = True
        Case CEL_TIPO_FORMULA
            BotaoFormula.Value = True
        Case CEL_TIPO_TITULO
            BotaoTitulo.Value = True
        Case CEL_TIPO_VAZIO
            BotaoVazio.Value = True

    End Select

End Sub

Private Function Preenche_Titulo() As Long
'Preenche o titulo

Dim lErro As Long
Dim iIndice As Integer
Dim objRelFormula As ClassRelDMPLFormula
Dim objRel As ClassRelDMPL
Dim iLinha As Integer
Dim iLinhaAtual As Integer
Dim iColunaAtual As Integer

On Error GoTo Erro_Preenche_Titulo

    'Limpa o grid
    Call Grid_Limpa(objGridFormulas)

    If giCelFixaCol = -1 Then
    
        iLinhaAtual = GridRelatorio.Row
        iColunaAtual = GridRelatorio.Col
    Else
        iLinhaAtual = giCelFixaLin
        iColunaAtual = giCelFixaCol
    End If

    For Each objRel In gcolRel

        If objRel.iLinha = iLinhaAtual And objRel.iColuna = iColunaAtual Then

            Titulo.Text = objRel.sTitulo
    
            Exit For
    
        End If

    Next

    Preenche_Titulo = SUCESSO

    Exit Function

Erro_Preenche_Titulo:

    Preenche_Titulo = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166711)

    End Select

    Exit Function

End Function



'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_REL_DMPL_CONFIG
    Set Form_Load_Ocx = Me
    Caption = "Configuração do Demonstrativo de Mutação do Patrimônio Líquido"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelDMPLConfig"
    
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




Private Sub LabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContas, Source, X, Y)
End Sub

Private Sub LabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContas, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub BotaoConta_Click()

Dim lErro As Long
Dim sContaOrigem As String
Dim iContaPreenchida As Integer
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoConta_Click

    If GridContas.Col = iGrid_ContaInicio_Col Then
    
        If Len(Trim(ContaInicio.ClipText)) > 0 Then
        
            lErro = CF("Conta_Formata", ContaInicio.Text, sContaOrigem, iContaPreenchida)
            If lErro <> SUCESSO Then gError 197943
    
            If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sContaOrigem
        Else
            objPlanoConta.sConta = ""
        End If
        
        'Chama a tela que lista os vendedores
        Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoConta)
    
    ElseIf GridContas.Col = iGrid_ContaFinal_Col Then
    
        If Len(Trim(ContaFim.ClipText)) > 0 Then
        
            lErro = CF("Conta_Formata", ContaFim.Text, sContaOrigem, iContaPreenchida)
            If lErro <> SUCESSO Then gError 197943
    
            If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sContaOrigem
        Else
            objPlanoConta.sConta = ""
        End If
    
        'Chama a tela que lista os vendedores
        Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoConta)
        
    End If

    Exit Sub
    
Erro_BotaoConta_Click:

    Select Case gErr
        
        Case 197943
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197945)
            
    End Select

    Exit Sub
    
End Sub

Private Sub objEventoConta_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoConta_evSelecao
    
    Set objPlanoConta = obj1
    
    sContaEnxuta = String(STRING_CONTA, 0)

    lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
    If lErro <> SUCESSO Then gError 197919
    
    If GridContas.Col = iGrid_ContaInicio_Col Then
        ContaInicio.PromptInclude = False
        ContaInicio.Text = sContaEnxuta
        ContaInicio.PromptInclude = True
        If Not (Me.ActiveControl Is ContaInicio) Then GridContas.TextMatrix(GridContas.Row, iGrid_ContaInicio_Col) = ContaInicio.Text
    ElseIf GridContas.Col = iGrid_ContaFinal_Col Then
        ContaFim.PromptInclude = False
        ContaFim.Text = sContaEnxuta
        ContaFim.PromptInclude = True
        If Not (Me.ActiveControl Is ContaFim) Then GridContas.TextMatrix(GridContas.Row, iGrid_ContaFinal_Col) = ContaFim.Text
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoConta_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197915)
        
    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ContaInicio Then Call BotaoConta_Click
        If Me.ActiveControl Is ContaFim Then Call BotaoConta_Click
    
    End If
    
End Sub

