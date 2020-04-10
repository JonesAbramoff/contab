VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl CclTelaOcx 
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   KeyPreview      =   -1  'True
   ScaleHeight     =   1560
   ScaleWidth      =   6630
   Begin VB.TextBox CclDescricao 
      Height          =   315
      Left            =   1110
      MaxLength       =   100
      TabIndex        =   1
      Top             =   1020
      Width           =   2940
   End
   Begin VB.CheckBox ativo 
      Caption         =   "Ativo"
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
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CclTelaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CclTelaOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CclTelaOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "CclTelaOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox TipoCcl 
      Height          =   315
      ItemData        =   "CclTelaOcx.ctx":0994
      Left            =   4785
      List            =   "CclTelaOcx.ctx":099E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1020
      Width           =   1635
   End
   Begin MSMask.MaskEdBox Ccl 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   315
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   " "
   End
   Begin MSComctlLib.TreeView TvwCcl 
      Height          =   1905
      Left            =   150
      TabIndex        =   3
      Top             =   1845
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   3360
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label CclLabel 
      AutoSize        =   -1  'True
      Caption         =   "Centro de Custo/Lucro:"
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
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   10
      Top             =   360
      Width           =   2025
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4230
      TabIndex        =   11
      Top             =   1080
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Centros de Custo / Lucro"
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
      Left            =   180
      TabIndex        =   12
      Top             =   1605
      Width           =   2175
   End
End
Attribute VB_Name = "CclTelaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1

Dim iAlterado As Integer

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim vbMsgRet As VbMsgBoxResult
Dim iTemFilho As Integer
Dim sCcl As String
Dim iCclPreenchido As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'critica o formato do centro de custo/lucro
    lErro = CF("Ccl_Formata", Ccl.Text, sCcl, iCclPreenchido)
    If lErro <> SUCESSO Then Error 10726

    If iCclPreenchido = CCL_VAZIA Then Error 8025

    objCcl.sCcl = sCcl
    
    'Verifica se o centro de custo/lucro existe
    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO And lErro <> 5599 Then Error 8002
        
    'Se o centro de custo não está cadastrado ==> erro
    If lErro = 5599 Then Error 8003
    
    'se o centro de custo for analítico
    If objCcl.iTipoCcl = CCL_ANALITICA Then
    
        'verifica se foi utilizado em movimento de estoque
        lErro = CF("MovimentoEstoque_Le_Ccl", objCcl.sCcl)
        If lErro <> SUCESSO And lErro <> 60868 Then Error 60873
        
        'se encontrou algum movimento de estoque utilizando o centro de custo
        If lErro = SUCESSO Then Error 60874
    
        'Verificar se existem associacoes com contas.
        lErro = CF("ContaCcl_Le_Ccl", objCcl.sCcl)
        If lErro <> SUCESSO And lErro <> 5603 Then Error 8006
    
        'Se tem associação com conta
        If lErro = SUCESSO Then
    
            'Verificar se existem documentos automáticos para o ccl.
            lErro = CF("DocAuto_Le_Ccl", objCcl.sCcl)
            If lErro <> SUCESSO And lErro <> 8097 Then Error 8007
        
            'Se existem documentos automaticos
            If lErro = SUCESSO Then
            
                'Existem associacoes com contas e documentos automaticos ==> avisar.
                vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CCL_COM_ASSOC_DOCAUTO")
            
            Else
            
                'Apenas associacoes com contas, avisar.
                vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CCL_COM_ASSOCIACOES")
                
            End If
        
        Else
            'Não existem associacões com contas
            vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CCL")
    
        End If
        
    Else
    
        'se o centro de custo não for analítico
        
        'verifica se o centro de custo tem filhos
        lErro = CF("Ccl_Tem_Filho", objCcl.sCcl, iTemFilho)
        If lErro <> SUCESSO Then Error 5594
        
        'se tiver filhos
        If iTemFilho = CCL_TEM_FILHOS Then
            'avisa que vai excluir o centro de custo e seus filhos
            vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CCL_SINTETICA_COM_FILHOS")
        Else
            'avisa que vai excluir o centro de custo
            vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CCL_SINTETICA")
        End If
        
    End If
    
    'se o usuário confirmar a exclusão
    If vbMsgRet = vbYes Then
        
        'exclui o centro de custo/lucro
        lErro = CF("Ccl_Exclui", objCcl.sCcl)
        If lErro <> SUCESSO Then Error 8026
        
'        'exclui o centro de custo/lucro da árvore
'        lErro = Excluir_Arvore_Ccl(TvwCcl.Nodes, objCcl)
'        If lErro <> SUCESSO Then Error 12245
    
        'limpar a tela
        Call Limpa_Tela(Me)
        
        iAlterado = 0
    
    End If
      
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 8002, 8006, 8007, 8026, 10726, 12245, 60873
        
        Case 8003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, objCcl.sCcl)
            Ccl.SetFocus
                    
        Case 8025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_INFORMADO", Err)
            Ccl.SetFocus
            
        Case 60874
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_SINTETICA_USADA_EM_MOVESTOQUE", Err, objCcl.sCcl)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 144314)
        
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chamar a rotina de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 9934
    
    'limpa a tela
    Call Limpa_Tela(Me)
    
    ativo.Value = 1
    
    iAlterado = 0
    
    Exit Sub
       
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 9934
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144315)

     End Select
        
     Exit Sub
       
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCcl As String
Dim iCclPreenchido As Integer
Dim vbMsgRet As VbMsgBoxResult
Dim iTemFilho As Integer

On Error GoTo Erro_Gravar_Registro
        
    GL_objMDIForm.MousePointer = vbHourglass
        
    sCcl = String(STRING_CCL, 0)

    'critica o formato do centro de custo/lucro
    lErro = CF("Ccl_Formata", Ccl.Text, sCcl, iCclPreenchido)
    If lErro <> SUCESSO Then Error 9612

    'testa se o centro de custo/lucro está preenchido
    If iCclPreenchido <> CCL_PREENCHIDA Then Error 9613
    
    objCcl.sCcl = sCcl

    'verifica se o centro de custo/lucro possui um centro de custo/lucro pai
    lErro = CF("Ccl_Critica_CclPai", sCcl)
    If lErro <> SUCESSO Then Error 10463

    'le o ccl para ver se ele esta cadastrado
    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO And lErro <> 5599 Then Error 12246

    objCcl.sDescCcl = CclDescricao.Text
    objCcl.iTipoCcl = TipoCcl.ItemData(TipoCcl.ListIndex)
    
    If ativo.Value = vbChecked Then
        objCcl.iAtivo = 1
    Else
        objCcl.iAtivo = 0
    End If

    'se o centro de custo está cadastrado
    If lErro = SUCESSO Then
    
        If objCcl.iAtivo = 0 Then
        
            'verifica se o centro de custo tem filhos
            lErro = CF("Ccl_Tem_Filho", objCcl.sCcl, iTemFilho)
            If lErro <> SUCESSO Then Error 32295
        
            'se tiver filhos
            If iTemFilho = CCL_TEM_FILHOS Then
            
                'avisa que vai desativar os filhos também
                vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_ATUALIZAR_CCL_COM_FILHOS")
                If vbMsgRet <> vbYes Then Error 32296
                
            End If
            
        End If

        lErro = Trata_Alteracao(objCcl, objCcl.sCcl)
        If lErro <> SUCESSO Then Error 32298

        'chamar a rotina para atualizar o centro de custo
        lErro = Atualizar_Ccl(objCcl)
        If lErro <> SUCESSO Then Error 9610
    
    Else
        
        'se o centro de custo não está cadastrado
        'chamar a rotina para inserir o centro de custo
        lErro = Inserir_Ccl(objCcl)
        If lErro <> SUCESSO Then Error 9611
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:
    
    Gravar_Registro = Err
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 9610, 9611, 9612, 10463, 12246
            
        Case 9613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_INFORMADO", Err)
    
        Case 32295, 32296, 32298
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144316)

     End Select
        
     Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'verifica se o usuário deseja salvar as informações da tela antes de apagá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 9624

    'limpar a tela
    Call Limpa_Tela(Me)
    
    ativo.Value = 1
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0

    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case Err
    
        Case 9624
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144317)
        
    End Select
        
    Exit Sub

End Sub

Private Sub Ccl_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Ccl_Validate(Cancel As Boolean)
     
Dim lErro As Long
Dim sCclFormatado As String
Dim iCclPreenchido As Integer

On Error GoTo Erro_Ccl_Validate

    If Len(Ccl.ClipText) > 0 Then

        sCclFormatado = String(STRING_CCL, 0)

        'critica o formato do centro de custo
        lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatado, iCclPreenchido)
        If lErro <> SUCESSO Then Error 9621
    
    End If
    
    Exit Sub
    
Erro_Ccl_Validate:

    Cancel = True


    Select Case Err
    
        Case 9621
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144318)
        
    End Select

    Exit Sub
    
End Sub

Private Sub CclDescricao_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim iIndice As Integer

On Error GoTo Erro_Ccl_Form_Load

    Set objEventoCcl = New AdmEvento

    'inicializa a mascara de centro de custo/lucro
    lErro = Inicializa_Mascara_Ccl()
    If lErro <> SUCESSO Then Error 9583
    
    'selecionar "Sintética" como opção inicial de TipoCcl
    For iIndice = 0 To TipoCcl.ListCount - 1
        If TipoCcl.ItemData(iIndice) = CCL_SINTETICA Then
            TipoCcl.ListIndex = iIndice
            Exit For
        End If
    Next
    
'    'Inicializa a Lista de Centro de Custos e Lucros
'    lErro = Carga_Arvore_Ccl(TvwCcl.Nodes)
'    If lErro <> SUCESSO Then Error 12239

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Ccl_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 9583, 12239
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144319)

    End Select
    
    iAlterado = 0
    
End Sub

Function Trata_Parametros(Optional objCcl As ClassCcl) As Long

Dim lErro As Long
Dim iLote As Integer

On Error GoTo Erro_Trata_Parametros

    'Se foi passado um centro de custo/lucro como parametro, exibir seus dados
    If Not (objCcl Is Nothing) Then
    
        'traz os dados do centro de custo passado como parametro para a tela.
        lErro = Traz_Ccl_Tela(objCcl.sCcl)
        If lErro <> SUCESSO Then Error 10477
        
    End If
            
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 10477
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144320)
    
    End Select
    
    Exit Function

End Function

Private Function Traz_Ccl_Tela(sCcl As String) As Long
'traz os dados do centro de custo passado como parametro para a tela.

Dim objCcl As New ClassCcl
Dim lErro As Long
Dim iIndice As Integer
Dim sCclEnxuta As String

On Error GoTo Erro_Traz_Ccl_Tela

    'limpa a tela
    Call Limpa_Tela(Me)
    
    sCclEnxuta = String(STRING_CCL, 0)
    
    'colocar a mascara no centro de custo
    lErro = Mascara_RetornaCclEnxuta(sCcl, sCclEnxuta)
    If lErro <> SUCESSO Then Error 10476
            
    'colocar o centro de custo na tela
    Ccl.PromptInclude = False
    Ccl.Text = sCclEnxuta
    Ccl.PromptInclude = True

    objCcl.sCcl = sCcl
    
    'Verifica se o centro de custo/lucro existe
    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO And lErro <> 5599 Then Error 8030
    
    If objCcl.iAtivo = 1 Then
        ativo.Value = vbChecked
    Else
        ativo.Value = vbUnchecked
    End If
    
    'se o centro de custo estiver cadastrado
    If lErro = SUCESSO Then
    
        'traz seus dados para a tela
        CclDescricao.Text = objCcl.sDescCcl
    
        For iIndice = 0 To TipoCcl.ListCount - 1
            If TipoCcl.ItemData(iIndice) = objCcl.iTipoCcl Then
                TipoCcl.ListIndex = iIndice
                Exit For
            End If
        Next
        
    End If
            
    iAlterado = 0
    
    Traz_Ccl_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Ccl_Tela:

    Traz_Ccl_Tela = Err
    
    Select Case Err
    
        Case 8030
    
        Case 10476
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", Err, sCcl)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144321)
    
    End Select
    
    iAlterado = 0
    
    Exit Function
    
End Function

Private Function Carga_Arvore_Ccl(colNodes As Nodes) As Long
'move os dados de centro de custo/lucro do banco de dados para a arvore colNodes.

Dim objNode As Node
Dim colCcl As New Collection
Dim objCcl As ClassCcl
Dim lErro As Long
Dim sCclMascarado As String
Dim sCcl As String
Dim sCclPai As String
    
On Error GoTo Erro_Carga_Arvore_Ccl
    
    'leitura dos centro de custo/lucro no BD
    lErro = CF("Ccl_Le_Todos", colCcl)
    If lErro <> SUCESSO Then Error 12240
    
    'para cada centro de custo encontrado no bd
    For Each objCcl In colCcl
        
        sCclMascarado = String(STRING_CCL, 0)

        'coloca a mascara no centro de custo
        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 12241

        sCcl = "C" & objCcl.sCcl

        sCclPai = String(STRING_CCL, 0)
        
        'retorna o centro de custo/lucro "pai" da centro de custo/lucro em questão, se houver
        lErro = Mascara_RetornaCclPai(objCcl.sCcl, sCclPai)
        If lErro <> SUCESSO Then Error 10368
        
        'se o centro de custo/lucro possui um centro de custo/lucro "pai"
        If Len(Trim(sCclPai)) > 0 Then

            sCclPai = "C" & sCclPai
            
            'adiciona o centro de custo como filho do centro de custo pai
            Set objNode = colNodes.Add(colNodes.Item(sCclPai), tvwChild, sCcl)

        Else
        
            'se o centro de custo/lucro não possui centro de custo/lucro "pai", adiciona na árvore sem pai
            Set objNode = colNodes.Add(, tvwLast, sCcl)
            
        End If
        
        'coloca o texto do nó que acabou de ser inserido
        objNode.Text = sCclMascarado & SEPARADOR & objCcl.sDescCcl
        
    Next
    
    Carga_Arvore_Ccl = SUCESSO

    Exit Function

Erro_Carga_Arvore_Ccl:

    Carga_Arvore_Ccl = Err

    Select Case Err

        Case 10368
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclPai", Err, objCcl.sCcl)

        Case 12240

        Case 12241
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144322)

    End Select
    
    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    Set objEventoCcl = Nothing
    
   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub TipoCcl_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TipoCcl_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iNivel As Integer
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl

On Error GoTo Erro_TipoCcl_Validate

    'se o centro de custo não estiver preenchido ==> não continua a execução
    If Len(Ccl.Text) = 0 Then Exit Sub
        
    'transforma o centro de custo no formato do banco de dados
    lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then Error 10472
    
    objCcl.sCcl = sCclFormatada
    
    'verifica se o centro de custo/lucro já está cadastrado
    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO And lErro <> 5599 Then Error 10473
        
    'se nao encontrou o centro de custo/lucro ==> não continua a crítica
    If lErro = 5599 Then Exit Sub
    
    objCcl.iTipoCcl = TipoCcl.ItemData(TipoCcl.ListIndex)
    
    'critica o tipo do centro de custo/lucro
    lErro = CF("Ccl_Critica_Tipo", objCcl)
    If lErro <> SUCESSO Then Error 10474
    
    Exit Sub
        
Erro_TipoCcl_Validate:

    Cancel = True


    Select Case Err
    
        Case 10472, 10473, 10474
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144323)
        
    End Select

    Exit Sub

End Sub

Private Sub TvwCcl_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim sCcl As String
Dim lErro As Long
    
On Error GoTo Erro_TvwCcl_NodeClick
    
    'pega o centro de custo seleconado na árvore
    sCcl = right(Node.Key, Len(Node.Key) - 1)

    'traz seus dados para a tela
    lErro = Traz_Ccl_Tela(sCcl)
    If lErro <> SUCESSO Then Error 10478
        
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_TvwCcl_NodeClick:

    Select Case Err
    
        Case 10478
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144324)
            
    End Select
        
    Exit Sub
    
End Sub

Private Function Alterar_Arvore_Ccl(colNode As Nodes, objCcl As ClassCcl) As Long
'Atualizar a TreeView com o objCcl

Dim objNode As Node
Dim lErro As Long
Dim sCclMascarado As String
Dim iAchou As Integer

On Error GoTo Error_Alterar_Arvore_Ccl

    iAchou = 0

    For Each objNode In colNode
                               
        'verifica pela chave
        If objNode.Key = "C" & objCcl.sCcl Then
            
            sCclMascarado = String(STRING_CCL, 0)
            
            lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then Error 12244
            
            objNode.Text = sCclMascarado & SEPARADOR & objCcl.sDescCcl
            
            iAchou = 1
            
            Exit For
        
        End If
    
    Next
    
    If iAchou = 0 Then Error 9616
    
    Alterar_Arvore_Ccl = SUCESSO

    Exit Function

Error_Alterar_Arvore_Ccl:

    Alterar_Arvore_Ccl = Err
 
    Select Case Err

        Case 9616
        
        Case 12244
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144325)
    
    End Select
    
    Exit Function
     
End Function

Private Function Inserir_Arvore_Ccl(colNodes As Nodes, objCcl As ClassCcl) As Long
'insere o ccl na arvore

Dim objNode As Node
Dim lErro As Long
Dim sCclMascarado As String
Dim sCclPai As String
    
On Error GoTo Erro_Inserir_Arvore_Ccl
    
    sCclPai = String(STRING_CCL, 0)
        
    sCclMascarado = String(STRING_CCL, 0)
            
    'colocar a mascara no centro de custo
    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then Error 12248
        
    'retorna o ccl "pai" em questão
    lErro = Mascara_RetornaCclPai(objCcl.sCcl, sCclPai)
    If lErro <> SUCESSO Then Error 20419
        
    'se o ccl possui um ccl "pai"
    If Len(Trim(sCclPai)) > 0 Then

        sCclPai = "C" & sCclPai
            
        'insere o centro de custo na arvore como filho do nó pai
        Set objNode = colNodes.Add(colNodes.Item(sCclPai), tvwChild, "C" & objCcl.sCcl, sCclMascarado & SEPARADOR & objCcl.sDescCcl)
        colNodes.Item(sCclPai).Sorted = True

    Else
    
        'se o ccl não possui ccl "pai" ==> insere o nó na arvore sem pai
        Set objNode = colNodes.Add(, , "C" & objCcl.sCcl, sCclMascarado & SEPARADOR & objCcl.sDescCcl)
        TvwCcl.Sorted = True
    
    End If
    
    Inserir_Arvore_Ccl = SUCESSO

    Exit Function

Erro_Inserir_Arvore_Ccl:

    Inserir_Arvore_Ccl = Err

    Select Case Err

        Case 12248
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)
                  
        Case 20419
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclPai", Err, objCcl.sCcl)
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144326)

    End Select
    
    Exit Function

End Function

Private Function Excluir_Arvore_Ccl(colNodes As Nodes, objCcl As ClassCcl) As Long
'Exclui o Ccl da Arvore

Dim objNode As Node
    
    For Each objNode In colNodes
        If objNode.Key = "C" & objCcl.sCcl Then
            colNodes.Remove (objNode.Index)
            Exit For
        End If
    Next
    
    Excluir_Arvore_Ccl = SUCESSO

End Function

Private Function Inicializa_Mascara_Ccl() As Long
'inicializa a mascara de centro de custo/lucro /m

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Ccl

    'Inicializa a máscara de Centro de custo/lucro
    sMascaraCcl = String(STRING_CCL, 0)
    
    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then Error 9582
    
    'coloca a mascara na tela.
    Ccl.Mask = sMascaraCcl
    
    Inicializa_Mascara_Ccl = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Ccl:

    Inicializa_Mascara_Ccl = Err
    
    Select Case Err
    
        Case 9582
            lErro = Rotina_Erro(vbOKOnly, "Erro_MascaraCcl", Err)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144327)
        
    End Select

    Exit Function

End Function

Private Function Atualizar_Ccl(objCcl As ClassCcl) As Long
'/m atualiza os dados do centro de custo

Dim lErro As Long

On Error GoTo Erro_Atualizar_Ccl

    'altera os dados do centro de custo no bd
    lErro = CF("Ccl_Altera", objCcl)
    If lErro <> SUCESSO Then Error 9614
    
    'alterar a definição do centro de custo na arvore de centros de custo
'    lErro = Alterar_Arvore_Ccl(TvwCcl.Nodes, objCcl)
'    If lErro <> SUCESSO And lErro <> 9616 Then Error 9615
'
'    'Se o centro de custo/lucro não estava cadastrado na arvore
'    If lErro = 9616 Then
'
'        'inserir o centro de custo na arvore
'        lErro = Inserir_Arvore_Ccl(TvwCcl.Nodes, objCcl)
'        If lErro <> SUCESSO Then Error 9617
'
'    End If
    
    Atualizar_Ccl = SUCESSO
    
    Exit Function
    
Erro_Atualizar_Ccl:

    Atualizar_Ccl = Err
    
    Select Case Err
    
        Case 9614, 9615, 9617
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144328)
        
    End Select

    Exit Function
        
End Function

Private Function Inserir_Ccl(objCcl As ClassCcl) As Long

Dim lErro As Long

On Error GoTo Erro_Inserir_Ccl

    'insere o centro de custo/lucro no banco de dados
    lErro = CF("Ccl_Insere", objCcl)
    If lErro <> SUCESSO Then Error 9618
    
'    'exclui da arvore de centros de custo se estiver cadastrado
'    Call Excluir_Arvore_Ccl(TvwCcl.Nodes, objCcl)
'
'    'inserir o centro de custo na arvore
'    lErro = Inserir_Arvore_Ccl(TvwCcl.Nodes, objCcl)
'    If lErro <> SUCESSO Then Error 9619
    
    Inserir_Ccl = SUCESSO
    
    Exit Function
    
Erro_Inserir_Ccl:

    Inserir_Ccl = Err
    
    Select Case Err
    
        Case 9618, 9619
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144329)
        
    End Select

    Exit Function
        
End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCcl As String
Dim iCclPreenchido As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Ccl"

    sCcl = String(STRING_CCL, 0)
    
    If Len(Trim(Ccl.ClipText)) > 0 Then
    
        'critica o formato do centro de custo/lucro
        lErro = CF("Ccl_Formata", Ccl.Text, sCcl, iCclPreenchido)
        If lErro <> SUCESSO Then Error 14953

        objCcl.sCcl = sCcl
    
    Else
        objCcl.sCcl = ""
    
    End If
    
    objCcl.sDescCcl = CclDescricao.Text
    objCcl.iTipoCcl = TipoCcl.ItemData(TipoCcl.ListIndex)
    
    If ativo.Value = vbChecked Then
        objCcl.iAtivo = 1
    Else
        objCcl.iAtivo = 0
    End If
    

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Ccl", objCcl.sCcl, STRING_CCL, "Ccl"
    colCampoValor.Add "DescCcl", objCcl.sDescCcl, STRING_CCL_DESCRICAO, "DescCcl"
    colCampoValor.Add "TipoCcl", objCcl.iTipoCcl, 0, "TipoCcl"
    colCampoValor.Add "AtivoCcl", objCcl.iTipoCcl, 0, "AtivoCcl"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 14953
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144330)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCcl As New ClassCcl

On Error GoTo Erro_Tela_Preenche

    objCcl.sCcl = colCampoValor.Item("Ccl").vValor

    If objCcl.sCcl <> 0 Then

        lErro = Traz_Ccl_Tela(objCcl.sCcl)
        If lErro <> SUCESSO Then Error 14955

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 14955

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144331)

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

    Parent.HelpContextID = IDH_CENTRO_CUSTO_CENTRO_LUCRO
    Set Form_Load_Ocx = Me
    Caption = "Centro de Custo/Centro de Lucro"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CclTela"
    
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



Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub CclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclLabel, Source, X, Y)
End Sub

Private Sub CclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclLabel, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub CclLabel_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Ccl.Text) > 0 Then objCcl.sCcl = Ccl.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("CclTodosLista", colSelecao, objCcl, objEventoCcl)

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objCcl As ClassCcl
Dim sCcl As String

On Error GoTo Erro_objEventoCcl_evSelecao
    
    Set objCcl = obj1

    sCcl = objCcl.sCcl

    'traz seus dados para a tela
    lErro = Traz_Ccl_Tela(sCcl)
    If lErro <> SUCESSO Then gError 197924

    Me.Show
    
    Exit Sub
    
Erro_objEventoCcl_evSelecao:

    Select Case gErr
    
        Case 197924

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197925)
        
    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Ccl Then
            Call CclLabel_Click
        End If
    
    End If

End Sub

