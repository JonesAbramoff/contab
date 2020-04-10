VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl MaoDeObraOcx 
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7875
   KeyPreview      =   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   7875
   Begin VB.CheckBox Ativo 
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
      Left            =   3600
      TabIndex        =   1
      Top             =   255
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5580
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "MaoDeObraOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "MaoDeObraOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "MaoDeObraOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "MaoDeObraOcx.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox Observacao 
      Height          =   315
      Left            =   1710
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1515
      Width           =   6015
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2565
      Picture         =   "MaoDeObraOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Numeração Automática"
      Top             =   255
      Width           =   300
   End
   Begin VB.TextBox Nome 
      Height          =   315
      Left            =   1710
      TabIndex        =   2
      Top             =   675
      Width           =   3720
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1710
      TabIndex        =   0
      Top             =   240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzido 
      Height          =   315
      Left            =   1710
      TabIndex        =   3
      Top             =   1095
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TipoMO 
      Height          =   315
      Left            =   1710
      TabIndex        =   15
      Top             =   1980
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin VB.Label LabelDescTipoMO 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2235
      TabIndex        =   17
      Top             =   1980
      Width           =   4530
   End
   Begin VB.Label LabelTipo 
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
      Left            =   1200
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   16
      Top             =   2025
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Observação:"
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
      Left            =   555
      TabIndex        =   13
      Top             =   1590
      Width           =   1095
   End
   Begin VB.Label LabelNomeRed 
      AutoSize        =   -1  'True
      Caption         =   "Nome Reduzido:"
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
      Left            =   240
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   12
      Top             =   1140
      Width           =   1410
   End
   Begin VB.Label LabelNome 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
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
      Left            =   1095
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   735
      Width           =   555
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   990
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   10
      Top             =   315
      Width           =   660
   End
End
Attribute VB_Name = "MaoDeObraOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Private WithEvents objEventoTipoDeMaodeObra As AdmEvento
Attribute objEventoTipoDeMaodeObra.VB_VarHelpID = -1
Private WithEvents objEventoMO As AdmEvento
Attribute objEventoMO.VB_VarHelpID = -1

Dim iAlterado As Integer

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Form_Load
    
    Ativo.Value = MARCADO
    
    'Inicializa variávies AdmEvento
    Set objEventoTipoDeMaodeObra = New AdmEvento
    Set objEventoMO = New AdmEvento
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 193857
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193858)

    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático do próximo cliente
    lErro = CF("MO_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 193859

    'Exibe código na Tela
    Codigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 193859
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 193860)
    
    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objMO As New ClassMO
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 193858

    objMO.lCodigo = CLng(Codigo.Text)

    'Lê os dados da mao de obra a ser excluido
    lErro = CF("MO_Le", objMO)
    If lErro <> SUCESSO And lErro <> 193817 Then gError 193859

    'Verifica se mao de obra não está cadastrado
    If lErro <> SUCESSO Then gError 193860

    'Envia aviso perguntando se realmente deseja excluir cliente e suas filiais
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_MAODEOBRA", objMO.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a Mao de Obra
        lErro = CF("MO_Exclui", objMO)
        If lErro <> SUCESSO Then gError 193861

        'Limpa a Tela
        lErro = Limpa_Tela_MO()
        If lErro <> SUCESSO Then gError 193862

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 193858
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_MO_NAO_PREENCHIDO", gErr)

        Case 193859, 193861, 193862
        
        Case 193860
            Call Rotina_Erro(vbOKOnly, "ERRO_MO_INEXISTENTE", gErr, objMO.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193863)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Ativo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoTipoDeMaodeObra_Click()

End Sub

Private Sub LabelTipo_Click()

Dim colSelecao As New Collection
Dim objTipoMO As New ClassTiposDeMaodeObra
    
On Error GoTo Erro_LabelTipo_Click
    
    objTipoMO.iCodigo = StrParaInt(TipoMO.Text)
        
    Call Chama_Tela("TiposDeMaodeObraLista", colSelecao, objTipoMO, objEventoTipoDeMaodeObra)

    Exit Sub

Erro_LabelTipo_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193913)

    End Select

    Exit Sub

End Sub

Public Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'Critica se é um Long
    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then gError 193864

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 193864

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193865)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o Cliente
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 193866

    'Limpa a Tela
    lErro = Limpa_Tela_MO()
    If lErro <> SUCESSO Then gError 193867
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 193866, 193867

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193868)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Verifica se dados de Cliente necessários foram preenchidos
'Grava Cliente no BD
'Atualiza ListBox de Clientes

Dim lErro As Long
Dim iIndice As Integer
Dim objMO As New ClassMO

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi preenchido o Código
    If Len(Trim(Codigo.Text)) = 0 Then gError 193869

    'Verifica se foi preenchido o Nome
    If Len(Trim(Nome.Text)) = 0 Then gError 193870

    'Verifica se foi preenchido o Nome Reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 193871

    If Len(Trim(TipoMO.Text)) = 0 Then gError 195579

    'Lê os dados da Tela relacionados ao Cliente
    lErro = Move_Tela_Memoria(objMO)
    If lErro <> SUCESSO Then gError 193873

    lErro = Trata_Alteracao(objMO, objMO.lCodigo)
    If lErro <> SUCESSO Then gError 193874
    
    'Grava o Cliente no BD
    lErro = CF("MO_Grava", objMO)
    If lErro <> SUCESSO Then gError 193875

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 193869
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_MO_NAO_PREENCHIDO", gErr)

        Case 193870
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_MO_NAO_PREENCHIDO", gErr)

        Case 193871
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMERED_MO_NAO_PREENCHIDO", gErr)

        Case 193873 To 193875

        Case 195579
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193876)

    End Select

    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 193877

    'Limpa a Tela
    lErro = Limpa_Tela_MO()
    If lErro <> SUCESSO Then gError 193878
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 193877, 193878

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193879)

    End Select

End Sub


Public Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NomeReduzido_Validate(Cancel As Boolean)

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate

    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then

        If Not IniciaLetra(NomeReduzido.Text) Then gError 193880

    End If

    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True

    Select Case Err

        Case 193880
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", gErr, NomeReduzido.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193881)

    End Select

End Sub

Public Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoTipoDeMaodeObra_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTiposDeMaodeObra As ClassTiposDeMaodeObra
Dim iLinha As Integer

On Error GoTo Erro_objEventoTipoDeMaodeObra_evSelecao

    Set objTiposDeMaodeObra = obj1

    TipoMO.PromptInclude = False
    TipoMO.Text = CStr(objTiposDeMaodeObra.iCodigo)
    TipoMO.PromptInclude = True
    
    LabelDescTipoMO.Caption = objTiposDeMaodeObra.sDescricao
    
    iAlterado = REGISTRO_ALTERADO
    
    Me.Show

    Exit Sub

Erro_objEventoTipoDeMaodeObra_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193922)

    End Select

    Exit Sub

End Sub

Function Limpa_Tela_MO() As Long

Dim iIndice As Integer
Dim iIndice2 As Integer
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_MO

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    'Limpa as TextBox e as MaskedEditBox
    Call Limpa_Tela(Me)
    
    LabelDescTipoMO.Caption = ""

    Ativo.Value = MARCADO
    
    Limpa_Tela_MO = SUCESSO
    
    Exit Function
    
Erro_Limpa_Tela_MO:

    Limpa_Tela_MO = Err
    
    Select Case Err
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193891)

    End Select
    
    Exit Function
        
End Function

Function Trata_Parametros(Optional objMO As ClassMO) As Long

Dim lErro As Long
Dim objClienteEstatistica As New ClassFilialClienteEst

On Error GoTo Erro_Trata_Parametros

    'Se houver MO passado como parâmetro, exibe seus dados
    If Not (objMO Is Nothing) Then

        'Se Codigo é positivo
        If objMO.lCodigo > 0 Then

            'Lê MaodeOBra no BD a partir do código
            lErro = CF("MO_Le", objMO)
            If lErro <> SUCESSO And lErro <> 193817 Then gError 193892

            'Se não encontrou a MaodeObra no BD
            If lErro <> SUCESSO Then

                'Limpa a Tela e exibe apenas o código
                lErro = Limpa_Tela_MO()
                If lErro <> SUCESSO Then gError 193893
                
                Codigo.Text = CStr(objMO.lCodigo)

            Else  'Encontrou Cliente no BD

                'Exibe os dados da MO
                lErro = Traz_MO_Tela(objMO)
                If lErro <> SUCESSO Then gError 193894

            End If

        'se Nome Reduzido está preenchido
        ElseIf Len(Trim(objMO.sNomeReduzido)) > 0 Then

            'Lê a MO no BD a partir do Nome Reduzido
            lErro = CF("MO_Le_NomeRed", objMO)
            If lErro <> SUCESSO And lErro <> 193855 Then gError 193895

            'Se não encontrou a MO no BD
            If lErro <> SUCESSO Then

                'Limpa a Tela e exibe apenas o NomeReduzido
                lErro = Limpa_Tela_MO()
                If lErro <> SUCESSO Then gError 193896
                
                NomeReduzido.Text = CStr(objMO.sNomeReduzido)

            Else  'Encontrou MO no BD

                'Exibe os dados da MO
                lErro = Traz_MO_Tela(objMO)
                If lErro <> SUCESSO Then gError 193897

            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 193892 To 193897

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193898)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Traz_MO_Tela(objMO As ClassMO) As Long
'Exibe os dados de MO na tela

Dim lErro As Long
Dim bCancel As Boolean
Dim objMOTipo As ClassMOTipo
Dim iIndice As Integer
Dim objTiposDeMaodeObra As New ClassTiposDeMaodeObra

On Error GoTo Erro_Traz_MO_Tela

    lErro = Limpa_Tela_MO()
    If lErro <> SUCESSO Then gError 195372

    lErro = CF("MO_Le", objMO)
    If lErro <> SUCESSO And lErro <> 193817 Then gError 193899

    'Verifica se MO não está cadastrado
    If lErro <> SUCESSO Then gError 193900
    
    objTiposDeMaodeObra.iCodigo = objMO.iTipo

    lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
    If lErro <> SUCESSO And lErro <> 137598 Then gError 195578

    Ativo.Value = objMO.iAtivo
    Codigo.Text = CStr(objMO.lCodigo)
    Nome.Text = objMO.sNome
    NomeReduzido.Text = objMO.sNomeReduzido
    Observacao.Text = objMO.sObservacao
    LabelDescTipoMO.Caption = objTiposDeMaodeObra.sDescricao
    TipoMO.PromptInclude = False
    TipoMO.Text = CStr(objMO.iTipo)
    TipoMO.PromptInclude = True

    iAlterado = 0

    Traz_MO_Tela = SUCESSO

    Exit Function

Erro_Traz_MO_Tela:
    
    Traz_MO_Tela = gErr

    Select Case gErr
        
        Case 193899, 195372, 195578
        
        Case 193900
            Call Rotina_Erro(vbOKOnly, "ERRO_MO_INEXISTENTE", gErr, objMO.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193901)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objMO As ClassMO) As Long
'Lê os dados que estão na tela de MaoDeObra e coloca em objMO

Dim lErro As Long
Dim iIndice As Integer
Dim objMOTipo As ClassMOTipo

On Error GoTo Erro_Move_Tela_Memoria

    If Len(Trim(Codigo.Text)) > 0 Then objMO.lCodigo = CLng(Codigo.Text)

    objMO.sNome = Trim(Nome.Text)
    objMO.sNomeReduzido = Trim(NomeReduzido.Text)

    objMO.sObservacao = Trim(Observacao.Text)
    
    objMO.iAtivo = Ativo.Value
    
    objMO.iTipo = StrParaInt(TipoMO.Text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193902)

    End Select

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objMO As New ClassMO

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "MaoDeObra"

    'Lê os dados da Tela MO
    lErro = Move_Tela_Memoria(objMO)
    If lErro <> SUCESSO Then gError 193903

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objMO.lCodigo, 0, "Codigo"
    colCampoValor.Add "Nome", objMO.sNome, STRING_MO_NOME, "Nome"
    colCampoValor.Add "NomeReduzido", objMO.sNomeReduzido, STRING_MO_NOMERED, "NomeReduzido"
    colCampoValor.Add "Observacao", objMO.sObservacao, STRING_MO_OBS, "Observacao"
    colCampoValor.Add "Ativo", objMO.iAtivo, 0, "Ativo"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 193903

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193904)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objMO As New ClassMO

On Error GoTo Erro_Tela_Preenche

    objMO.lCodigo = colCampoValor.Item("Codigo").vValor

    If objMO.lCodigo <> 0 Then

        'Carrega objMO com os dados passados em colCampoValor
        objMO.sNome = colCampoValor.Item("Nome").vValor
        objMO.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
        objMO.sObservacao = colCampoValor.Item("Observacao").vValor
        objMO.iAtivo = colCampoValor.Item("Ativo").vValor

        'Exibe a mao de obra na Tela
        lErro = Traz_MO_Tela(objMO)
        If lErro <> SUCESSO Then gError 193905

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 193905

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193906)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    Set objEventoTipoDeMaodeObra = Nothing
    Set objEventoMO = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub LabelCodigo_Click()

Dim objMO As New ClassMO
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(Codigo.Text)) > 0 Then objMO.lCodigo = CDbl(Codigo.Text)

    'Chama Tela ClienteLista
    Call Chama_Tela("MaoDeObraLista", colSelecao, objMO, objEventoMO)

End Sub

Public Sub LabelNomeRed_Click()

Dim objMO As New ClassMO
Dim colSelecao As Collection

On Error GoTo Erro_LabelNomeRed_Click

    objMO.sNomeReduzido = NomeReduzido.Text

    'Chama Tela ClienteLista
    Call Chama_Tela("MaoDeObraLista", colSelecao, objMO, objEventoMO)

    Exit Sub

Erro_LabelNomeRed_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193907)

    End Select

    Exit Sub


End Sub

Private Sub objEventoMO_evSelecao(obj1 As Object)

Dim objMO As ClassMO
Dim bCancel As Boolean

    Set objMO = obj1

    'Executa o Validate
    Call MO_Traz_Tela(objMO.lCodigo)

    Me.Show

    Exit Sub

End Sub

Public Sub MO_Traz_Tela(ByVal lCodigo As Long)

Dim lErro As Long
Dim objMO As New ClassMO

On Error GoTo Erro_MO_Traz_Tela

    'Guarda o valor do código da mao de obra
    objMO.lCodigo = lCodigo

    'Lê a mao de obra no BD
    lErro = CF("MO_Le", objMO)
    If lErro <> SUCESSO And lErro <> 193817 Then gError 193908

    'Se mao de obra não está cadastrado, erro
    If lErro <> SUCESSO Then gError 193909

    'Exibe a mao de obra na Tela
    lErro = Traz_MO_Tela(objMO)
    If lErro <> SUCESSO Then gError 193910

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_MO_Traz_Tela:

    Select Case gErr

        Case 193908, 193910

        Case 193909
            Call Rotina_Erro(vbOKOnly, "ERRO_MO_INEXISTENTE", gErr, objMO.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193911)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = 0
    Set Form_Load_Ocx = Me
    Caption = "Mão de Obra"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "MaoDeObra"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
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

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
'''    m_Caption = New_Caption
End Property

Private Sub TipoMO_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

'***** fim do trecho a ser copiado ******

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is TipoMO Then
            Call LabelTipo_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is NomeReduzido Then
            Call LabelNomeRed_Click
        End If
    
    End If

End Sub

Private Sub TipoMO_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTiposDeMaodeObra As New ClassTiposDeMaodeObra

On Error GoTo Erro_TipoMO_Validate

    If Len(TipoMO.Text) > 0 Then
    
        lErro = Inteiro_Critica(TipoMO.Text)
        If lErro <> SUCESSO Then gError 195370
    
        objTiposDeMaodeObra.iCodigo = StrParaInt(TipoMO.Text)
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 193918
        
        If lErro <> SUCESSO Then gError 193919
        
    End If

    LabelDescTipoMO.Caption = objTiposDeMaodeObra.sDescricao

    Exit Sub

Erro_TipoMO_Validate:

    Cancel = True

    Select Case gErr

        Case 193918, 195370

        Case 193919
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_CADASTRADO", gErr, objTiposDeMaodeObra.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 193920)

    End Select

    Exit Sub

End Sub

