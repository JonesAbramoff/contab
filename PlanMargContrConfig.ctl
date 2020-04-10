VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl PlanMargContrConfigOcx 
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9180
   ScaleHeight     =   6045
   ScaleWidth      =   9180
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4905
      Index           =   1
      Left            =   270
      TabIndex        =   2
      Top             =   795
      Width           =   8775
      Begin VB.Frame Frame2 
         Caption         =   "Configuração das Linhas"
         Height          =   2310
         Index           =   1
         Left            =   75
         TabIndex        =   5
         Top             =   2505
         Width           =   8580
         Begin VB.ComboBox Formato 
            Height          =   315
            ItemData        =   "PlanMargContrConfig.ctx":0000
            Left            =   480
            List            =   "PlanMargContrConfig.ctx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox Formula8 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6375
            MaxLength       =   255
            TabIndex        =   23
            Top             =   1500
            Width           =   1890
         End
         Begin VB.TextBox Formula7 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   4440
            MaxLength       =   255
            TabIndex        =   22
            Top             =   1500
            Width           =   1890
         End
         Begin VB.TextBox Formula6 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   2475
            MaxLength       =   255
            TabIndex        =   21
            Top             =   1545
            Width           =   1890
         End
         Begin VB.TextBox Formula5 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   570
            MaxLength       =   255
            TabIndex        =   20
            Top             =   1530
            Width           =   1890
         End
         Begin VB.TextBox Formula4 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6285
            MaxLength       =   255
            TabIndex        =   19
            Top             =   1095
            Width           =   1890
         End
         Begin VB.CheckBox GrupoL1 
            Caption         =   "Check1"
            Height          =   225
            Left            =   5550
            TabIndex        =   18
            Top             =   435
            Width           =   1170
         End
         Begin VB.TextBox FormulaL1 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6180
            MaxLength       =   255
            TabIndex        =   17
            Top             =   510
            Width           =   2430
         End
         Begin VB.TextBox Formula3 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   4320
            MaxLength       =   255
            TabIndex        =   16
            Top             =   1080
            Width           =   1890
         End
         Begin VB.TextBox Formula2 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   2355
            MaxLength       =   255
            TabIndex        =   15
            Top             =   1125
            Width           =   1890
         End
         Begin VB.TextBox Formula1 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   360
            MaxLength       =   255
            TabIndex        =   14
            Top             =   1200
            Width           =   1890
         End
         Begin VB.TextBox FormulaGeral 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3600
            MaxLength       =   255
            TabIndex        =   13
            Top             =   480
            Width           =   1890
         End
         Begin VB.TextBox AnaliseLinDescricao 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   240
            MaxLength       =   255
            TabIndex        =   12
            Top             =   480
            Width           =   3360
         End
         Begin MSFlexGridLib.MSFlexGrid GridAnaliseLin 
            Height          =   1950
            Left            =   165
            TabIndex        =   7
            Top             =   240
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   3440
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Configuração das Colunas"
         Height          =   2310
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   8580
         Begin VB.TextBox AnaliseColTitulo 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1320
            MaxLength       =   255
            TabIndex        =   10
            Top             =   990
            Width           =   1260
         End
         Begin VB.TextBox AnaliseColDescricao 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   2685
            MaxLength       =   255
            TabIndex        =   11
            Top             =   990
            Width           =   3360
         End
         Begin MSFlexGridLib.MSFlexGrid GridAnaliseCol 
            Height          =   1905
            Left            =   135
            TabIndex        =   6
            Top             =   360
            Width           =   8280
            _ExtentX        =   14605
            _ExtentY        =   3360
            _Version        =   393216
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4845
      Index           =   2
      Left            =   255
      TabIndex        =   3
      Top             =   870
      Visible         =   0   'False
      Width           =   8745
      Begin VB.Frame Frame3 
         Caption         =   "Configuração"
         Height          =   4695
         Left            =   45
         TabIndex        =   8
         Top             =   75
         Width           =   8670
         Begin VB.TextBox DVVDescricao 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   240
            MaxLength       =   255
            TabIndex        =   24
            Top             =   840
            Width           =   3360
         End
         Begin VB.TextBox DVVFormulaSimulacao 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   5535
            MaxLength       =   255
            TabIndex        =   27
            Top             =   1845
            Width           =   1890
         End
         Begin VB.TextBox DVVFormulaCliente 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   4620
            MaxLength       =   255
            TabIndex        =   26
            Top             =   1320
            Width           =   1890
         End
         Begin VB.TextBox DVVFormulaPadrao 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3660
            MaxLength       =   255
            TabIndex        =   25
            Top             =   840
            Width           =   1890
         End
         Begin MSFlexGridLib.MSFlexGrid GridDVV 
            Height          =   4215
            Left            =   165
            TabIndex        =   9
            Top             =   315
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   7435
            _Version        =   393216
         End
      End
   End
   Begin VB.PictureBox Picture1 
      DrawStyle       =   1  'Dash
      Height          =   555
      Left            =   6900
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1560
         Picture         =   "PlanMargContrConfig.ctx":0022
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1080
         Picture         =   "PlanMargContrConfig.ctx":01A0
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "PlanMargContrConfig.ctx":06D2
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "PlanMargContrConfig.ctx":085C
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5355
      Left            =   165
      TabIndex        =   1
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9446
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Planilha de Análise"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Planilha de Despesas Variáveis de Venda"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "PlanMargContrConfigOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Numero maximo de linhas nos grids ANALISECOL, ANALISELIN, DVV respectivamente
Const NUM_MAXIMO_LINHAS_ANALISELIN = 200
Const NUM_MAXIMO_LINHAS_ANALISECOL = 8 + 1
Const NUM_MAXIMO_LINHAS_DVV = 200

'Constantes para TabStrips
Const TAB_Analise = 1
Const TAB_DVV = 2

'Grid Colunas
Public iGrid_NumeroColuna_Col As Integer
Public iGrid_Titulo_Col As Integer
Public iGrid_Descricao_Col As Integer

'Grid Linhas
Public iGrid_Descricao1_Col As Integer
Public iGrid_FormulaGeral_Col As Integer
Public iGrid_GrupoL1_Col As Integer
Public iGrid_FormulaL1_Col As Integer
Public iGrid_Formato_Col As Integer
Public iGrid_FormulaColuna1_Col As Integer
Public iGrid_FormulaColuna2_Col As Integer
Public iGrid_FormulaColuna3_Col As Integer
Public iGrid_FormulaColuna4_Col As Integer
Public iGrid_FormulaColuna5_Col As Integer
Public iGrid_FormulaColuna6_Col As Integer
Public iGrid_FormulaColuna7_Col As Integer
Public iGrid_FormulaColuna8_Col As Integer

'Grid DVV
Public iGrid_Descricao2_Col As Integer
Public iGrid_FormulaPadrao_Col As Integer
Public iGrid_FormulaCliente_Col As Integer
Public iGrid_FormulaSimulacao_Col As Integer

Dim giFrameAtual As Integer
Public iAlterado As Integer

'Objetos dos Grids
Dim objGridAnaliseCol As AdmGrid
Dim objGridAnaliseLin As AdmGrid
Dim objGridDVV As AdmGrid

'Property Variables:
Dim m_Caption As String
Dim iFrameAtual As Integer

Event Unload()

Public Sub Form_Load()

Dim lErro As Long
Dim objMargContr As New ClassMargContr

On Error GoTo Erro_Form_Load

    giFrameAtual = 1

    'inicializa Objetos dos Grids
    Set objGridAnaliseCol = New AdmGrid
    Set objGridAnaliseLin = New AdmGrid
    Set objGridDVV = New AdmGrid
    
    'Faz as Inicializações dos Grids
    lErro = Inicializa_Grid_AnaliseCol(objGridAnaliseCol)
    If lErro <> SUCESSO Then gError 116961
    
    lErro = Inicializa_Grid_AnaliseLin(objGridAnaliseLin)
    If lErro <> SUCESSO Then gError 116962
    
    lErro = Inicializa_Grid_DVV(objGridDVV)
    If lErro <> SUCESSO Then gError 116963
    
    'carrega dados das tabelas usadas pelos grids Config. de Colunas e Config. de Linhas
    lErro = CF("MargContr_Le_Analise", objMargContr)
    If lErro <> SUCESSO Then gError 116964
    
    'Preenche dados no grid Configuração de Linhas
    lErro = Preenche_Grid_AnaliseCol(objMargContr)
    If lErro <> SUCESSO Then gError 116965
            
    'Preenche dados no grid Configuração de Linhas
    lErro = Preenche_Grid_AnaliseLin(objMargContr)
    If lErro <> SUCESSO Then gError 116966
            
    'Carrega dados das tabelas usadas pelos GRIDDvv
    lErro = CF("MargContr_Le_DVV", objMargContr)
    If lErro <> SUCESSO Then gError 116967
    
    'Preenche dados no grid DVV
    lErro = Preenche_Grid_DVV(objMargContr)
    If lErro <> SUCESSO Then gError 116968
               
    iAlterado = 0
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 116961 To 116968

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164889)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Preenche_Grid_AnaliseCol(ByVal objMargContr As ClassMargContr) As Long
'Prenche Grid Configuração de Colunas

Dim lErro As Long
Dim objPlanMargContrCol As ClassPlanMargContrCol

On Error GoTo Erro_Preenche_Grid_AnaliseCol

    'Preenche o Grid AnaliseCol com dados da tabela PlanMargContrCol
    For Each objPlanMargContrCol In objMargContr.colPlanMargContrCol
        
        GridAnaliseCol.TextMatrix(objPlanMargContrCol.iColuna, iGrid_Titulo_Col) = objPlanMargContrCol.sTitulo
        GridAnaliseCol.TextMatrix(objPlanMargContrCol.iColuna, iGrid_Descricao_Col) = objPlanMargContrCol.sDescricao

    Next
        
    objGridAnaliseCol.iLinhasExistentes = objMargContr.colPlanMargContrCol.Count
    
    Preenche_Grid_AnaliseCol = SUCESSO
    
    Exit Function
    
Erro_Preenche_Grid_AnaliseCol:

    Preenche_Grid_AnaliseCol = gErr

    Select Case gErr

       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164890)

    End Select

    Exit Function
    
End Function

Private Function Preenche_Grid_AnaliseLin(ByVal objMargContr As ClassMargContr) As Long
'Preenche Grid de configuração de Linhas

Dim lErro As Long
Dim objPlanMargContrLin As ClassPlanMargContrLin
Dim objPlanMargContrLinCol As ClassPlanMargContrLinCol
Dim iLinha As Integer, iColuna As Integer

On Error GoTo Erro_Preenche_Grid_AnaliseLin

    'Preenche o Grid Configuração de Linhas com dados da tabela PlanMargContrLin
    For Each objPlanMargContrLin In objMargContr.colPlanMargContrLin

        iLinha = objPlanMargContrLin.iLinha
        
        GridAnaliseLin.TextMatrix(iLinha, iGrid_Descricao1_Col) = objPlanMargContrLin.sDescricao
        GridAnaliseLin.TextMatrix(iLinha, iGrid_FormulaGeral_Col) = objPlanMargContrLin.sFormulaGeral
        
        'Verifica se FormulaL1 está preenchida, se estiver, coloco Grupo L1 como MARCADO
        If objPlanMargContrLin.iEditavel <> 0 Then GridAnaliseLin.TextMatrix(objPlanMargContrLin.iLinha, iGrid_GrupoL1_Col) = MARCADO
        
        If objPlanMargContrLin.iFormato = GRID_FORMATO_MOEDA Then
            GridAnaliseLin.TextMatrix(iLinha, iGrid_Formato_Col) = GRID_FORMATO_MOEDA_STRING
        Else
            GridAnaliseLin.TextMatrix(iLinha, iGrid_Formato_Col) = GRID_FORMATO_PERCENTAGEM_STRING
        End If
        
        GridAnaliseLin.TextMatrix(iLinha, iGrid_FormulaL1_Col) = objPlanMargContrLin.sFormulaL1

        'Preenche o Grid Configuração de Linhas com dados da tabela PlanMargContrLinCol
        'cada objPlanMargContrLinCol, armazenada informações sobre uma coluna Fórmula1 ou Fórmula2 ou...ou Fórmula8
        For iColuna = 1 To objMargContr.colPlanMargContrCol.Count
        
            Set objPlanMargContrLinCol = objMargContr.colPlanMargContrLinCol(objMargContr.IndAnalise(iLinha, iColuna))
            GridAnaliseLin.TextMatrix(objPlanMargContrLinCol.iLinha, objPlanMargContrLinCol.iColuna + COLUNAS_GERAIS_ANALISE_LIN) = objPlanMargContrLinCol.sFormula
            
        Next
        
    Next
    
    objGridAnaliseLin.iLinhasExistentes = objMargContr.colPlanMargContrLin.Count

    Call Grid_Refresh_Checkbox(objGridAnaliseLin)
    
    Preenche_Grid_AnaliseLin = SUCESSO
    
    Exit Function

Erro_Preenche_Grid_AnaliseLin:

    Preenche_Grid_AnaliseLin = gErr

    Select Case gErr

       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164891)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_DVV(ByVal objMargContr As ClassMargContr) As Long
'Preenche Grid de Despesas Variaveis de Venda

Dim lErro As Long
Dim objDVVLinCol As ClassDVVLinCol
Dim objDVVLin As ClassDVVLin
Dim iLinha As Integer, iColuna As Integer

On Error GoTo Erro_Preenche_grid_DVV

    'Preenche o Grid DVV com dados da tabela DVVLin
    For Each objDVVLin In objMargContr.colDVVLin

        iLinha = objDVVLin.iLinha
        GridDVV.TextMatrix(iLinha, iGrid_Descricao2_Col) = objDVVLin.sDescricao

        'Preenche o Grid DVV com dados da tabela DVVLinCol
        For iColuna = 1 To MAX_NUM_FORMULAS_DVV

            'cada objDVVLinCol, armazenada informações sobre uma coluna Padrão ou cliente ou simulação
            Set objDVVLinCol = objMargContr.colDVVLinCol(objMargContr.IndDVV(iLinha, iColuna))
            GridDVV.TextMatrix(objDVVLinCol.iLinha, objDVVLinCol.iColuna + COLUNAS_GERAIS_DVV) = objDVVLinCol.sFormula
            
        Next
        
    Next
    
    objGridDVV.iLinhasExistentes = objMargContr.colDVVLin.Count
    
    Preenche_Grid_DVV = SUCESSO
    
    Exit Function

Erro_Preenche_grid_DVV:

    Preenche_Grid_DVV = gErr

    Select Case gErr

       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164892)

    End Select

    Exit Function

End Function

Private Sub AnaliseLinDescricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AnaliseLinDescricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseLin)
End Sub

Private Sub AnaliseLinDescricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseLin)
End Sub

Private Sub AnaliseLinDescricao_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseLin.objControle = AnaliseLinDescricao
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseLin)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Function Valida_Grid_AnaliseCol() As Long
'Faz a Validação dos dados do grid Configuração de Colunas

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Valida_Grid_AnaliseCol

    'Verifica se há itens no grid
    If objGridAnaliseCol.iLinhasExistentes = 0 Then gError 116969

    'para cada item do grid
    For iIndice = 1 To objGridAnaliseCol.iLinhasExistentes

        'Verifica se o Campo Titulo foi Preenchido
        If Len(Trim(GridAnaliseCol.TextMatrix(iIndice, iGrid_Titulo_Col))) = 0 Then gError 116970
        
    Next

    Valida_Grid_AnaliseCol = SUCESSO

    Exit Function

Erro_Valida_Grid_AnaliseCol:

    Valida_Grid_AnaliseCol = gErr

    Select Case gErr
    
       Case 116969
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_ANALISECOL_NAO_PREENCHIDO", gErr)
        
       Case 116970
            Call Rotina_Erro(vbOKOnly, "ERRO_ANALISECOL_TITULO_NAO_PREENCHIDO", gErr, iIndice)
       
       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164893)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_AnaliseLin() As Long
'Faz a Validação dos dados do grid Configuração de Linhas

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Valida_Grid_AnaliseLin

    'Verifica se há itens no grid
    If objGridAnaliseLin.iLinhasExistentes = 0 Then gError 116971

    'para cada item do grid
    For iIndice = 1 To objGridAnaliseLin.iLinhasExistentes

        'Verifica se Campo Descrição foi preenchido
        If Len(Trim(GridAnaliseLin.TextMatrix(iIndice, iGrid_Descricao1_Col))) = 0 Then gError 116972

        'Verifica se o GrupoL1 foi selecionado
        If StrParaInt(GridAnaliseLin.TextMatrix(iIndice, iGrid_GrupoL1_Col)) <> MARCADO Then
            
            'Se GrupoL1 não foi selecionado então FormulaL1 não pode ser preenchida
            If Len(Trim(GridAnaliseLin.TextMatrix(iIndice, iGrid_FormulaL1_Col))) > 0 Then gError 121045
        
        End If
        
    Next
    
    Valida_Grid_AnaliseLin = SUCESSO

    Exit Function

Erro_Valida_Grid_AnaliseLin:

    Valida_Grid_AnaliseLin = gErr

    Select Case gErr
    
       Case 116971
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_ANALISELIN_NAO_PREENCHIDO", gErr)
        
       Case 116972
            Call Rotina_Erro(vbOKOnly, "ERRO_ANALISELIN_DESCRICAO_NAO_PREENCHIDO", gErr, iIndice)

       Case 121044
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMULAL1_NAO_PREENCHIDA", gErr, iIndice)
       
       Case 121045
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMULAL1_PREENCHIDA", gErr, iIndice)
       
       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164894)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Dvv() As Long
'Faz a Validação dos dados do grid Despesas Variaveis de Venda

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Valida_Grid_Dvv

    'Verifica se há itens no grid
    If objGridDVV.iLinhasExistentes = 0 Then gError 116973

    'para cada item do grid
    For iIndice = 1 To objGridDVV.iLinhasExistentes
        
        'Verifica se Campo Descrição foi preenchido
        If Len(Trim(GridDVV.TextMatrix(iIndice, iGrid_Descricao2_Col))) = 0 Then gError 116974

    Next

    Valida_Grid_Dvv = SUCESSO

    Exit Function

Erro_Valida_Grid_Dvv:

    Valida_Grid_Dvv = gErr

    Select Case gErr
    
       Case 116973
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_DVV_NAO_PREENCHIDO", gErr)
        
       Case 116974
            Call Rotina_Erro(vbOKOnly, "ERRO_DVV_DESCRICAO_NAO_PREENCHIDO", gErr, iIndice)

       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164895)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long
'Aciona Rotinas de Validação de Grids e de Gravação

Dim lErro As Long
Dim objMargContr As New ClassMargContr

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Valida Grid Configuração de Colunas
    lErro = Valida_Grid_AnaliseCol
    If lErro <> SUCESSO Then gError 116975
           
    'Valida Grid de Configuração de Linhas
    lErro = Valida_Grid_AnaliseLin
    If lErro <> SUCESSO Then gError 116976

    'Valida Grid de Despesas Variaveis de Venda
    lErro = Valida_Grid_Dvv
    If lErro <> SUCESSO Then gError 116977

    'Preenche OBJs referentes ao Grid de Configuração de Colunas que serão usados na gravação nas tabelas
    lErro = Move_Tela_Memoria(objMargContr)
    If lErro <> SUCESSO Then gError 116978

    'Preenche OBJs referentes ao Grid de Configuração de Linhas que serão usados na gravação nas tabelas
    lErro = Move_Tela_Memoria1(objMargContr)
    If lErro <> SUCESSO Then gError 121083
    
    'Preenche OBJs referentes ao Grid de Despesas Variáveis de Venda que serão usados na gravação nas tabelas
    lErro = Move_Tela_Memoria2(objMargContr)
    If lErro <> SUCESSO Then gError 121084
    
    'Aciona Rotina de Gravação no BD
    lErro = CF("PlanMargContrConfig_Grava", objMargContr)
    If lErro <> SUCESSO Then gError 116979

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 116975 To 116979, 121083, 121084
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164896)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'Aciona Rotinas de exclusão de dados no BD

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Confirma exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_MARGCONTR")
    If vbMsgRes = vbNo Then Exit Sub

    'Aciona Rotina de Exclusão de dados
    lErro = CF("PlanMargContrConfig_Exclui")
    If lErro <> SUCESSO Then gError 121066
    
    'Limpa a Tela PLanMargContrConfig
    Call Limpa_PlanMargContrConfig
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 121066
              
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164897)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 116980

    'Informa Sucesso da Gravação
    vbResult = Rotina_Aviso(vbOKOnly, "AVISO_GRAVACAO_SUCESSO")
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116980

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 164898)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_BotaoLimpar_Click

    'Confirma se quer realmente limpar Grids
    vbResult = Rotina_Aviso(vbYesNo, "AVISO_LIMPA_TELA_MARGCONTR")
    If vbResult = vbNo Then Exit Sub
        
    'Limpa a Tela
    Call Limpa_PlanMargContrConfig
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 116981

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164899)

    End Select

    Exit Sub

End Sub

Sub GridDVV_Click()
    
Dim iExecutaEntradaCelula As Integer
Dim iAlterado As Integer

    Call Grid_Click(objGridDVV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGridDVV, iAlterado)

    End If
    
End Sub

Sub GridDVV_GotFocus()
    Call Grid_Recebe_Foco(objGridDVV)
End Sub

Sub GridDVV_EnterCell()

Dim iAlterado As Integer

    Call Grid_Entrada_Celula(objGridDVV, iAlterado)

End Sub

Sub GridDVV_LeaveCell()
    Call Saida_Celula(objGridDVV)
End Sub

Private Sub GridDVV_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridDVV)

End Sub

Private Sub GridDVV_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer
Dim iAlterado As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDVV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDVV, iAlterado)
    End If

End Sub

Sub GridDVV_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridDVV)
End Sub

Sub GridDVV_RowColChange()
    Call Grid_RowColChange(objGridDVV)
End Sub

Sub GridDVV_Scroll()
    Call Grid_Scroll(objGridDVV)
End Sub


Sub GridAnaliseLin_Click()
    
Dim iExecutaEntradaCelula As Integer
Dim iAlterado As Integer

    Call Grid_Click(objGridAnaliseLin, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGridAnaliseLin, iAlterado)

    End If
    
End Sub

Sub GridAnaliseLin_GotFocus()
    Call Grid_Recebe_Foco(objGridAnaliseLin)
End Sub

Sub GridAnaliseLin_EnterCell()

Dim iAlterado As Integer

    Call Grid_Entrada_Celula(objGridAnaliseLin, iAlterado)

End Sub

Sub GridAnaliseLin_LeaveCell()
    Call Saida_Celula(objGridAnaliseLin)
End Sub

Private Sub GridAnaliseLin_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridAnaliseLin)

End Sub

Private Sub GridAnaliseLin_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer
Dim iAlterado As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAnaliseLin, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAnaliseLin, iAlterado)
    End If

End Sub

Sub GridAnaliseLin_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridAnaliseLin)
End Sub

Sub GridAnaliseLin_RowColChange()
    Call Grid_RowColChange(objGridAnaliseLin)
End Sub

Sub GridAnaliseLin_Scroll()
    Call Grid_Scroll(objGridAnaliseLin)
End Sub


Sub GridAnaliseCol_Click()
    
Dim iExecutaEntradaCelula As Integer
Dim iAlterado As Integer

    Call Grid_Click(objGridAnaliseCol, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGridAnaliseCol, iAlterado)

    End If
    
End Sub

Sub GridAnaliseCol_GotFocus()
    Call Grid_Recebe_Foco(objGridAnaliseCol)
End Sub

Sub GridAnaliseCol_EnterCell()

Dim iAlterado As Integer

    Call Grid_Entrada_Celula(objGridAnaliseCol, iAlterado)

End Sub

Sub GridAnaliseCol_LeaveCell()
    Call Saida_Celula(objGridAnaliseCol)
End Sub

Private Sub GridAnaliseCol_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridAnaliseCol)

End Sub

Private Sub GridAnaliseCol_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer
Dim iAlterado As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAnaliseCol, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAnaliseCol, iAlterado)
    End If

End Sub

Sub GridAnaliseCol_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridAnaliseCol)
End Sub

Sub GridAnaliseCol_RowColChange()
    Call Grid_RowColChange(objGridAnaliseCol)
End Sub

Sub GridAnaliseCol_Scroll()
    Call Grid_Scroll(objGridAnaliseCol)
End Sub

Private Sub AnaliseColTitulo_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AnaliseColTitulo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseCol)
End Sub

Private Sub AnaliseColTitulo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseCol)
End Sub

Private Sub AnaliseColTitulo_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseCol.objControle = AnaliseColTitulo
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseCol)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub AnaliseColDescricao_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AnaliseColDescricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseCol)
End Sub

Private Sub AnaliseColDescricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseCol)
End Sub

Private Sub AnaliseColDescricao_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseCol.objControle = AnaliseColDescricao
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseCol)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub DVVdescricao_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DVVdescricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDVV)
End Sub

Private Sub DVVdescricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDVV)
End Sub

Private Sub DVVdescricao_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridDVV.objControle = DVVDescricao
    lErro = Grid_Campo_Libera_Foco(objGridDVV)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub DVVFormulapadrao_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DVVFormulapadrao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDVV)
End Sub

Private Sub DVVFormulapadrao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDVV)
End Sub

Private Sub DVVFormulapadrao_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridDVV.objControle = DVVFormulaPadrao
    lErro = Grid_Campo_Libera_Foco(objGridDVV)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub DVVFormulaCliente_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DVVFormulaCliente_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDVV)
End Sub

Private Sub DVVFormulaCliente_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDVV)
End Sub

Private Sub DVVFormulaCliente_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridDVV.objControle = DVVFormulaCliente
    lErro = Grid_Campo_Libera_Foco(objGridDVV)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub DVVFormulaSimulacao_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DVVFormulaSimulacao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDVV)
End Sub

Private Sub DVVFormulaSimulacao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDVV)
End Sub

Private Sub DVVFormulaSimulacao_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridDVV.objControle = DVVFormulaSimulacao
    lErro = Grid_Campo_Libera_Foco(objGridDVV)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GrupoL1_Click()

    'Limpa a Célula Fórmula L1, caso o campo Grupo L1, tenha sido desmarcado
    If GridAnaliseLin.TextMatrix(GridAnaliseLin.Row, iGrid_GrupoL1_Col) = DESMARCADO Then GridAnaliseLin.TextMatrix(GridAnaliseLin.Row, iGrid_FormulaL1_Col) = ""
    
End Sub

Private Sub GrupoL1_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseLin)
End Sub

Private Sub GrupoL1_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseLin)
End Sub

Private Sub GrupoL1_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseLin.objControle = GrupoL1
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseLin)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub FormulaGeral_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FormulaGeral_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseLin)
End Sub

Private Sub FormulaGeral_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseLin)
End Sub

Private Sub FormulaGeral_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseLin.objControle = FormulaGeral
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseLin)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub FormulaL1_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FormulaL1_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseLin)
End Sub

Private Sub FormulaL1_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseLin)
End Sub

Private Sub FormulaL1_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseLin.objControle = FormulaL1
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseLin)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub
Private Sub Formato_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Formato_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseLin)
End Sub

Private Sub Formato_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseLin)
End Sub

Private Sub Formato_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseLin.objControle = Formato
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseLin)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Formula1_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Formula1_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseLin)
End Sub

Private Sub Formula1_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseLin)
End Sub

Private Sub Formula1_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseLin.objControle = Formula1
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseLin)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Formula2_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Formula2_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseLin)
End Sub

Private Sub Formula2_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseLin)
End Sub

Private Sub Formula2_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseLin.objControle = Formula2
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseLin)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Formula3_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Formula3_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseLin)
End Sub

Private Sub Formula3_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseLin)
End Sub

Private Sub Formula3_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseLin.objControle = Formula3
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseLin)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub
Private Sub Formula4_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Formula4_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseLin)
End Sub

Private Sub Formula4_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseLin)
End Sub

Private Sub Formula4_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseLin.objControle = Formula4
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseLin)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Formula5_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Formula5_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseLin)
End Sub

Private Sub Formula5_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseLin)
End Sub

Private Sub Formula5_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseLin.objControle = Formula5
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseLin)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Formula6_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Formula6_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseLin)
End Sub

Private Sub Formula6_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseLin)
End Sub

Private Sub Formula6_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseLin.objControle = Formula6
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseLin)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Formula7_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Formula7_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseLin)
End Sub

Private Sub Formula7_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseLin)
End Sub

Private Sub Formula7_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseLin.objControle = Formula7
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseLin)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Formula8_Change()
 iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Formula8_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnaliseLin)
End Sub

Private Sub Formula8_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnaliseLin)
End Sub

Private Sub Formula8_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridAnaliseLin.objControle = Formula8
    lErro = Grid_Campo_Libera_Foco(objGridAnaliseLin)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame1(giFrameAtual).Visible = False
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        
        'Armazena novo valor de giFrameAtual
        giFrameAtual = TabStrip1.SelectedItem.Index
       
    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164900)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_AnaliseCol(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Nº.")
    objGridInt.colColuna.Add ("Título")
    objGridInt.colColuna.Add ("Descrição")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (AnaliseColTitulo.Name)
    objGridInt.colCampo.Add (AnaliseColDescricao.Name)
    
    'Colunas do Grid
    iGrid_NumeroColuna_Col = 0
    iGrid_Titulo_Col = 1
    iGrid_Descricao_Col = 2
    
    'Grid do GridInterno
    objGridInt.objGrid = GridAnaliseCol

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_LINHAS_ANALISECOL

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridAnaliseCol.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    objGridInt.iProibidoIncluirNoMeioGrid = 0

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_Grid_AnaliseCol = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_AnaliseLin(objGridInt As AdmGrid) As Long
'Inicializa o Grid Linhas

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Nº.")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Fórmula Geral")
    objGridInt.colColuna.Add ("Editável")
    objGridInt.colColuna.Add ("Fórmula L1")
    objGridInt.colColuna.Add ("Formato")
    objGridInt.colColuna.Add ("Fórmula Coluna 1")
    objGridInt.colColuna.Add ("Fórmula Coluna 2")
    objGridInt.colColuna.Add ("Fórmula Coluna 3")
    objGridInt.colColuna.Add ("Fórmula Coluna 4")
    objGridInt.colColuna.Add ("Fórmula Coluna 5")
    objGridInt.colColuna.Add ("Fórmula Coluna 6")
    objGridInt.colColuna.Add ("Fórmula Coluna 7")
    objGridInt.colColuna.Add ("Fórmula Coluna 8")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (AnaliseLinDescricao.Name)
    objGridInt.colCampo.Add (FormulaGeral.Name)
    objGridInt.colCampo.Add (GrupoL1.Name)
    objGridInt.colCampo.Add (FormulaL1.Name)
    objGridInt.colCampo.Add (Formato.Name)
    objGridInt.colCampo.Add (Formula1.Name)
    objGridInt.colCampo.Add (Formula2.Name)
    objGridInt.colCampo.Add (Formula3.Name)
    objGridInt.colCampo.Add (Formula4.Name)
    objGridInt.colCampo.Add (Formula5.Name)
    objGridInt.colCampo.Add (Formula6.Name)
    objGridInt.colCampo.Add (Formula7.Name)
    objGridInt.colCampo.Add (Formula8.Name)
    
    'Controles que participam do Grid
    iGrid_NumeroColuna_Col = 0
    iGrid_Descricao1_Col = 1
    iGrid_FormulaGeral_Col = 2
    iGrid_GrupoL1_Col = 3
    iGrid_FormulaL1_Col = 4
    iGrid_Formato_Col = 5
    iGrid_FormulaColuna1_Col = 6
    iGrid_FormulaColuna2_Col = 7
    iGrid_FormulaColuna3_Col = 8
    iGrid_FormulaColuna4_Col = 9
    iGrid_FormulaColuna5_Col = 10
    iGrid_FormulaColuna6_Col = 11
    iGrid_FormulaColuna7_Col = 12
    iGrid_FormulaColuna8_Col = 13
    
    'Grid do GridInterno
    objGridInt.objGrid = GridAnaliseLin

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_LINHAS_ANALISELIN

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridAnaliseLin.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridInt.iProibidoIncluirNoMeioGrid = 0
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_AnaliseLin = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_DVV(objGridInt As AdmGrid) As Long
'Inicializa o Grid Linhas

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Nº.")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Padrão")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Simulação")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (DVVDescricao.Name)
    objGridInt.colCampo.Add (DVVFormulaPadrao.Name)
    objGridInt.colCampo.Add (DVVFormulaCliente.Name)
    objGridInt.colCampo.Add (DVVFormulaSimulacao.Name)
    
    'Controles que participam do Grid
    iGrid_NumeroColuna_Col = 0
    iGrid_Descricao2_Col = 1
    iGrid_FormulaPadrao_Col = 2
    iGrid_FormulaCliente_Col = 3
    iGrid_FormulaSimulacao_Col = 4
    
    'Grid do GridInterno
    objGridInt.objGrid = GridDVV

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_LINHAS_DVV

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 15

    'Largura da primeira coluna
    GridDVV.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridInt.iProibidoIncluirNoMeioGrid = 0

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_DVV = SUCESSO

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

On Error GoTo Erro_Rotina_Grid_Enable

    Select Case objControl.Name
       
        Case AnaliseColDescricao.Name
            
            'Se o Descrição estiver preenchida habilita
            If Len(Trim(GridAnaliseCol.TextMatrix(iLinha, iGrid_Titulo_Col))) > 0 Then
                AnaliseColDescricao.Enabled = True
            Else
                AnaliseColDescricao.Enabled = False
            End If
         
        Case FormulaGeral.Name, GrupoL1.Name, Formato.Name, Formula1.Name, Formula2.Name, Formula3.Name, Formula4.Name, Formula5.Name, Formula6.Name, Formula7.Name, Formula8.Name
            'Se a Descrição estiver preenchido habilita
            If Len(Trim(GridAnaliseLin.TextMatrix(iLinha, iGrid_Descricao1_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
         Case FormulaL1.Name
            'Se a Descrição estiver preenchido habilita
            If Len(Trim(GridAnaliseLin.TextMatrix(iLinha, iGrid_Descricao1_Col))) > 0 And GridAnaliseLin.TextMatrix(iLinha, iGrid_GrupoL1_Col) = MARCADO Then
           
                FormulaL1.Enabled = True
            Else
                FormulaL1.Enabled = False
            End If
                       
        Case DVVFormulaPadrao.Name, DVVFormulaCliente.Name, DVVFormulaSimulacao.Name
           'Se a Descrição estiver preenchido desabilita
           If Len(Trim(GridDVV.TextMatrix(iLinha, iGrid_Descricao2_Col))) > 0 Then
               objControl.Enabled = True
           Else
               objControl.Enabled = False
           End If
        
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164901)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula
    
    ' aciona rotina que inicializa a saida da celula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridAnaliseCol
            Case GridAnaliseCol.Name

                lErro = Saida_Celula_GridAnaliseCol(objGridInt)
                If lErro <> SUCESSO Then gError 116982

            'Se for o GridAnaliseLin
            Case GridAnaliseLin.Name

                lErro = Saida_Celula_GridAnaliseLin(objGridInt)
                If lErro <> SUCESSO Then gError 116983

            'Se for o GridDVV
            Case GridDVV.Name

                lErro = Saida_Celula_GridDVV(objGridInt)
                If lErro <> SUCESSO Then gError 116984


        End Select

        'Finaliza Saída da Celula
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 116985

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 116982 To 116985

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164902)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridAnaliseCol(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridAnaliseCol

    'Verifica qual a coluna atual do Grid Connfiguração de Colunas
    Select Case objGridInt.objGrid.Col

        'Se for a de Titulo
        Case iGrid_Titulo_Col
            lErro = Saida_Celula_Titulo(objGridInt)
            If lErro <> SUCESSO Then gError 116986

        'Se for a de Descricao
        Case iGrid_Descricao_Col
            lErro = Saida_Celula_Descricao(objGridInt)
            If lErro <> SUCESSO Then gError 116987

    End Select

    Saida_Celula_GridAnaliseCol = SUCESSO

    Exit Function

Erro_Saida_Celula_GridAnaliseCol:

    Saida_Celula_GridAnaliseCol = gErr

    Select Case gErr

        Case 116986, 116987

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164903)

    End Select

    Exit Function

End Function

Function Saida_Celula_Descricao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Descrição que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Descricao

    Set objGridInt.objControle = AnaliseColDescricao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116988
       
    Saida_Celula_Descricao = SUCESSO

    Exit Function

Erro_Saida_Celula_Descricao:

    Saida_Celula_Descricao = gErr

    Select Case gErr

        Case 116988
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164904)

    End Select

    Exit Function

End Function

Function Saida_Celula_Descricao1(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Descrição do Grid AnaliseLin que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Descricao1

    Set objGridInt.objControle = AnaliseLinDescricao

     'verifica se célula foi prenchida
     If Len(Trim(AnaliseLinDescricao.Text)) > 0 Then
    
        'Acrescenta uma linha no Grid se for o caso
        If GridAnaliseLin.Row - GridAnaliseLin.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116989
    
    'Preenche célula Formato com valor Default
    GridAnaliseLin.TextMatrix(objGridInt.iLinhasExistentes, iGrid_Formato_Col) = GRID_FORMATO_MOEDA_STRING
    
    Saida_Celula_Descricao1 = SUCESSO

    Exit Function

Erro_Saida_Celula_Descricao1:

    Saida_Celula_Descricao1 = gErr

    Select Case gErr

        Case 116989
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164905)

    End Select

    Exit Function

End Function

Function Saida_Celula_Descricao2(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Descrição do grid DVV que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Descricao2

    Set objGridInt.objControle = DVVDescricao

     If Len(Trim(DVVDescricao.Text)) > 0 Then
     
        'Acrescenta uma linha no Grid se for o caso
        If GridDVV.Row - GridDVV.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116990
    
    Saida_Celula_Descricao2 = SUCESSO

    Exit Function

Erro_Saida_Celula_Descricao2:

    Saida_Celula_Descricao2 = gErr

    Select Case gErr

        Case 116990
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164906)

    End Select

    Exit Function

End Function

Function Saida_Celula_FormulaPadrao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Formula Padrao que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_FormulaPadrao

    Set objGridInt.objControle = DVVFormulaPadrao
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116991
    
    Saida_Celula_FormulaPadrao = SUCESSO

    Exit Function

Erro_Saida_Celula_FormulaPadrao:

    Saida_Celula_FormulaPadrao = gErr

    Select Case gErr

        Case 116991
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164907)

    End Select

    Exit Function

End Function

Function Saida_Celula_FormulaCliente(objGridInt As AdmGrid) As Long
'Faz a crítica da célula FormulaCliente que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_FormulaCliente

    Set objGridInt.objControle = DVVFormulaCliente

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116992
    
    Saida_Celula_FormulaCliente = SUCESSO

    Exit Function

Erro_Saida_Celula_FormulaCliente:

    Saida_Celula_FormulaCliente = gErr

    Select Case gErr

        Case 116992
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164908)

    End Select

    Exit Function

End Function

Function Saida_Celula_FormulaSimulacao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Formula Simulação que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_FormulaSimulacao

    Set objGridInt.objControle = DVVFormulaSimulacao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116993
    
    Saida_Celula_FormulaSimulacao = SUCESSO

    Exit Function

Erro_Saida_Celula_FormulaSimulacao:

    Saida_Celula_FormulaSimulacao = gErr

    Select Case gErr

        Case 116993
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164909)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridAnaliseLin(objGridInt As AdmGrid) As Long
'Verifica qual célula do Grid AnaliseLin está deixando de ser a corrente e aciona a respectiva função saida_celula

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridAnaliseLin

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Descrição
        Case iGrid_Descricao1_Col
            lErro = Saida_Celula_Descricao1(objGridInt)
            If lErro <> SUCESSO Then gError 121066

        'Se for a de Fórmula Geral
        Case iGrid_FormulaGeral_Col
            lErro = Saida_Celula_FormulaGeral(objGridInt)
            If lErro <> SUCESSO Then gError 121067

        'Se for a de Grupo L1
        Case iGrid_GrupoL1_Col
            lErro = Saida_Celula_GrupoL1(objGridInt)
            If lErro <> SUCESSO Then gError 121068

        'Se for a de Fórmula L1
        Case iGrid_FormulaL1_Col
            lErro = Saida_Celula_FormulaL1(objGridInt)
            If lErro <> SUCESSO Then gError 121069

        'Se for a de Formato
        Case iGrid_Formato_Col
            lErro = Saida_Celula_Formato(objGridInt)
            If lErro <> SUCESSO Then gError 121070
            
        'Se for a de Fórmula Coluna 1
        Case iGrid_FormulaColuna1_Col
            lErro = Saida_Celula_Formula1(objGridInt)
            If lErro <> SUCESSO Then gError 121071

        'Se for a de Fórmula Coluna 2
        Case iGrid_FormulaColuna2_Col
            lErro = Saida_Celula_Formula2(objGridInt)
            If lErro <> SUCESSO Then gError 121072

        'Se for a de Fórmula Coluna 3
        Case iGrid_FormulaColuna3_Col
            lErro = Saida_Celula_Formula3(objGridInt)
            If lErro <> SUCESSO Then gError 121073
        
        'Se for a de Fórmula Coluna 4
        Case iGrid_FormulaColuna4_Col
            lErro = Saida_Celula_Formula4(objGridInt)
            If lErro <> SUCESSO Then gError 121074

        'Se for a de Fórmula Coluna 5
        Case iGrid_FormulaColuna5_Col
            lErro = Saida_Celula_Formula5(objGridInt)
            If lErro <> SUCESSO Then gError 121075
            
        'Se for a de Fórmula Coluna 6
        Case iGrid_FormulaColuna6_Col
            lErro = Saida_Celula_Formula6(objGridInt)
            If lErro <> SUCESSO Then gError 121076

        'Se for a de Fórmula Coluna 7
        Case iGrid_FormulaColuna7_Col
            lErro = Saida_Celula_Formula7(objGridInt)
            If lErro <> SUCESSO Then gError 121077
            
        'Se for a de Fórmula Coluna 8
        Case iGrid_FormulaColuna8_Col
            lErro = Saida_Celula_Formula8(objGridInt)
            If lErro <> SUCESSO Then gError 121078

    End Select

    Saida_Celula_GridAnaliseLin = SUCESSO

    Exit Function

Erro_Saida_Celula_GridAnaliseLin:

    Saida_Celula_GridAnaliseLin = gErr

    Select Case gErr

        Case 121066 To 121078

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164910)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridDVV(objGridInt As AdmGrid) As Long
'Verifica qual célula do Grid DVV está deixando de ser a corrente e aciona a respectiva função saida_celula

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridDVV

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Descrição
        Case iGrid_Descricao2_Col
            lErro = Saida_Celula_Descricao2(objGridInt)
            If lErro <> SUCESSO Then gError 121079

        'Se for a de Fórmula Padrão
        Case iGrid_FormulaPadrao_Col
            lErro = Saida_Celula_FormulaPadrao(objGridInt)
            If lErro <> SUCESSO Then gError 121080

        'Se for a de Fórmula Cliente
        Case iGrid_FormulaCliente_Col
            lErro = Saida_Celula_FormulaCliente(objGridInt)
            If lErro <> SUCESSO Then gError 121081

        'Se for a de Fórmula Simulação
        Case iGrid_FormulaSimulacao_Col
            lErro = Saida_Celula_FormulaSimulacao(objGridInt)
            If lErro <> SUCESSO Then gError 121082
    
    End Select

    Saida_Celula_GridDVV = SUCESSO

    Exit Function

Erro_Saida_Celula_GridDVV:

    Saida_Celula_GridDVV = gErr

    Select Case gErr

        Case 121079 To 121082

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164911)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Titulo(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Titulo que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Titulo

    Set objGridInt.objControle = AnaliseColTitulo

    If Len(Trim(AnaliseColTitulo.Text)) > 0 Then

        'Acrescenta uma linha no Grid se for o caso
        If GridAnaliseCol.Row - GridAnaliseCol.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116994
    
    Saida_Celula_Titulo = SUCESSO

    Exit Function

Erro_Saida_Celula_Titulo:

    Saida_Celula_Titulo = gErr

    Select Case gErr

        Case 116994
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164912)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FormulaGeral(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Fórmula Geral que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_FormulaGeral

    Set objGridInt.objControle = FormulaGeral
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116995
    
    Saida_Celula_FormulaGeral = SUCESSO

    Exit Function

Erro_Saida_Celula_FormulaGeral:

    Saida_Celula_FormulaGeral = gErr

    Select Case gErr

        Case 116995
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164913)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GrupoL1(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Grupo L1 que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GrupoL1

    Set objGridInt.objControle = GrupoL1
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116996

    Saida_Celula_GrupoL1 = SUCESSO

    Exit Function

Erro_Saida_Celula_GrupoL1:

    Saida_Celula_GrupoL1 = gErr

    Select Case gErr

        Case 116996
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164914)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FormulaL1(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Fórmula L1 que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_FormulaL1

    Set objGridInt.objControle = FormulaL1
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116997
    
    Saida_Celula_FormulaL1 = SUCESSO

    Exit Function

Erro_Saida_Celula_FormulaL1:

    Saida_Celula_FormulaL1 = gErr

    Select Case gErr

        Case 116997
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164915)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Formula1(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Fórmula 1 que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Formula1

    Set objGridInt.objControle = Formula1
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116998
    
    Saida_Celula_Formula1 = SUCESSO

    Exit Function

Erro_Saida_Celula_Formula1:

    Saida_Celula_Formula1 = gErr

    Select Case gErr

        Case 116998
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164916)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Formato(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Formato que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Formato

    Set objGridInt.objControle = Formato
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116999
    
    Saida_Celula_Formato = SUCESSO

    Exit Function

Erro_Saida_Celula_Formato:

    Saida_Celula_Formato = gErr

    Select Case gErr

      Case 116999
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
      
      Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164917)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Formula2(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Formula 2 que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Formula2

    Set objGridInt.objControle = Formula2

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 121000
        
    Saida_Celula_Formula2 = SUCESSO

    Exit Function

Erro_Saida_Celula_Formula2:

    Saida_Celula_Formula2 = gErr

    Select Case gErr

        Case 121000
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164918)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Formula3(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Formula3 que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Formula3

    Set objGridInt.objControle = Formula3

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 121001
    
    Saida_Celula_Formula3 = SUCESSO

    Exit Function

Erro_Saida_Celula_Formula3:

    Saida_Celula_Formula3 = gErr

    Select Case gErr

        Case 121001
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164919)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Formula4(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Formula4 que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Formula4

    Set objGridInt.objControle = Formula4

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 121002
    
    Saida_Celula_Formula4 = SUCESSO

    Exit Function

Erro_Saida_Celula_Formula4:

    Saida_Celula_Formula4 = gErr

    Select Case gErr

        Case 121002
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164920)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Formula5(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Formula5 que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Formula5

    Set objGridInt.objControle = Formula5

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 121003
    
    Saida_Celula_Formula5 = SUCESSO

    Exit Function

Erro_Saida_Celula_Formula5:

    Saida_Celula_Formula5 = gErr

    Select Case gErr

        Case 121003
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164921)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Formula6(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Formula6 que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Formula6

    Set objGridInt.objControle = Formula6

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 121004
    
    Saida_Celula_Formula6 = SUCESSO

    Exit Function

Erro_Saida_Celula_Formula6:

    Saida_Celula_Formula6 = gErr

    Select Case gErr

        Case 121004
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164922)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Formula7(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Formula7 que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Formula7

    Set objGridInt.objControle = Formula7

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 121005
    
    Saida_Celula_Formula7 = SUCESSO

    Exit Function

Erro_Saida_Celula_Formula7:

    Saida_Celula_Formula7 = gErr

    Select Case gErr

        Case 121005
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164923)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Formula8(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Formula8 que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Formula8

    Set objGridInt.objControle = Formula8

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 121006
    
    Saida_Celula_Formula8 = SUCESSO

    Exit Function

Erro_Saida_Celula_Formula8:

    Saida_Celula_Formula8 = gErr

    Select Case gErr

        Case 121006
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164924)

    End Select

    Exit Function

End Function

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub Limpa_PlanMargContrConfig()
'Limpa os campos da tela sem fechar o sistema de setas

    Call Limpa_Tela(Me)

    Call Grid_Limpa(objGridAnaliseCol)
    Call Grid_Limpa(objGridAnaliseLin)
    Call Grid_Limpa(objGridDVV)
    
    iAlterado = 0
        
    Exit Sub

End Sub

Public Function Move_Tela_Memoria(ByVal objMargContr As ClassMargContr) As Long
'Preenche objMargContr com dados que no Grid de Configuração de Linhas

Dim objPlanMargContrCol As ClassPlanMargContrCol
Dim iIndice As Integer

On Error GoTo Erro_Move_Tela_Memoria
    
    'preenche o objPlanMargContrCol com os dados lidos do grid configuração de Colunas
    For iIndice = 1 To objGridAnaliseCol.iLinhasExistentes

        'Reinstancia o objPlanMargContrCol
        Set objPlanMargContrCol = New ClassPlanMargContrCol
        
        objPlanMargContrCol.iColuna = GridAnaliseCol.TextMatrix(iIndice, iGrid_NumeroColuna_Col)
        objPlanMargContrCol.sDescricao = GridAnaliseCol.TextMatrix(iIndice, iGrid_Descricao_Col)
        objPlanMargContrCol.sTitulo = GridAnaliseCol.TextMatrix(iIndice, iGrid_Titulo_Col)
        
        'Adiciona o ObjPlanMargContrCol na collection
        objMargContr.colPlanMargContrCol.Add objPlanMargContrCol
            
    Next
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164925)

    End Select

    Exit Function

End Function

Public Function Move_Tela_Memoria1(ByVal objMargContr As ClassMargContr) As Long
'Preenche objMargContr com dados que do Grid Configuração de Linhas

Dim objPlanMargContrLin As ClassPlanMargContrLin
Dim objPlanMargContrLinCol As ClassPlanMargContrLinCol
Dim iIndice As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_Move_Tela_Memoria1
    
    'preenche o objPlanMargContrLin com os dados lidos
    For iIndice = 1 To objGridAnaliseLin.iLinhasExistentes

        'Instancia o ObjPlanMargContrLin
        Set objPlanMargContrLin = New ClassPlanMargContrLin
                    
        'carrega objPlanMargContrLin com dados do Grid Configuração de Linhas
        objPlanMargContrLin.iLinha = iIndice
        objPlanMargContrLin.sDescricao = GridAnaliseLin.TextMatrix(iIndice, iGrid_Descricao1_Col)
        objPlanMargContrLin.sFormulaGeral = GridAnaliseLin.TextMatrix(iIndice, iGrid_FormulaGeral_Col)
        objPlanMargContrLin.sFormulaL1 = GridAnaliseLin.TextMatrix(iIndice, iGrid_FormulaL1_Col)
        objPlanMargContrLin.iEditavel = IIf(StrParaInt(GridAnaliseLin.TextMatrix(iIndice, iGrid_GrupoL1_Col)) = MARCADO, 1, 0)
        
        If GridAnaliseLin.TextMatrix(iIndice, iGrid_Formato_Col) = GRID_FORMATO_MOEDA_STRING Then
            objPlanMargContrLin.iFormato = GRID_FORMATO_MOEDA
        Else
            objPlanMargContrLin.iFormato = GRID_FORMATO_PERCENTAGEM
        End If
            
        'Carrega ObjPlanMargContrLinCol com dados de Fórmula1 até Fórmula8 do Grid de Configuração de Linhas
        For iIndice2 = 1 To MAX_NUM_FORMULAS_ANALISELIN
        
            'Instancia o objPLanMargContrLinCol
            Set objPlanMargContrLinCol = New ClassPlanMargContrLinCol
                                    
            objPlanMargContrLinCol.iLinha = iIndice
            objPlanMargContrLinCol.iColuna = iIndice2
            objPlanMargContrLinCol.sFormula = GridAnaliseLin.TextMatrix(iIndice, iGrid_FormulaColuna1_Col - 1 + iIndice2)
    
            'Adiciona o ObjPlanMargContrLinCol na collection
            objMargContr.colPlanMargContrLinCol.Add objPlanMargContrLinCol
            
        Next
        
        'Adiciona o objPlanMargContrLin na collection
        objMargContr.colPlanMargContrLin.Add objPlanMargContrLin
            
    Next
    
    Move_Tela_Memoria1 = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria1:

    Move_Tela_Memoria1 = gErr

    Select Case gErr

       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164926)

    End Select

    Exit Function

End Function

Public Function Move_Tela_Memoria2(ByVal objMargContr As ClassMargContr) As Long
'Preenche objMargContr com dados que estão no Grid DVV

Dim objDVVLin As ClassDVVLin
Dim objDVVLinCol As ClassDVVLinCol
Dim iIndice As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_Move_Tela_Memoria2
    
    'preenche o objdvv com os dados lidos
    For iIndice = 1 To objGridDVV.iLinhasExistentes

        'Instancia obj
        Set objDVVLin = New ClassDVVLin
                    
        'carrega objDVVLin
        objDVVLin.iLinha = iIndice
        objDVVLin.sDescricao = GridDVV.TextMatrix(iIndice, iGrid_Descricao2_Col)
        
        'Carrega ObjDVVLinCol com campos Padrão, Cliente e Simulação do Grid DVV
        For iIndice2 = 1 To MAX_NUM_FORMULAS_DVV
        
            'Instancia o ObjDVVLinCol
            Set objDVVLinCol = New ClassDVVLinCol
                    
            objDVVLinCol.iLinha = iIndice
            objDVVLinCol.iColuna = iIndice2
            objDVVLinCol.sFormula = GridDVV.TextMatrix(iIndice, iGrid_FormulaPadrao_Col - 1 + iIndice2)
            
            'Adiciona objDVVLinCol na collection
            objMargContr.colDVVLinCol.Add objDVVLinCol
            
        Next
        
        'Adiciona o objDVVLin na collection
        objMargContr.colDVVLin.Add objDVVLin
            
    Next

    Move_Tela_Memoria2 = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria2:

    Move_Tela_Memoria2 = gErr

    Select Case gErr

       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164927)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'Parent.HelpContextID = IDH_TIPOS_BLOQUEIO
    Set Form_Load_Ocx = Me
    Caption = "Configuração da Análise de Margem de Contribuição"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PlanMargContrConfig"
    
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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Public Sub Form_Unload(Cancel As Integer)

    Set objGridAnaliseCol = Nothing
    Set objGridAnaliseLin = Nothing
    Set objGridDVV = Nothing

End Sub
