VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl LoteEst 
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   5085
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2280
      Picture         =   "Lote.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   360
      Width           =   300
   End
   Begin VB.Frame Frame1 
      Caption         =   "Valores Atuais"
      Height          =   1785
      Left            =   135
      TabIndex        =   14
      Top             =   1920
      Width           =   4800
      Begin VB.CommandButton BotaoItens 
         Caption         =   "Itens do Lote"
         Height          =   810
         Left            =   2595
         Picture         =   "Lote.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   795
         Width           =   1380
      End
      Begin VB.CommandButton BotaoRecalcular 
         Caption         =   "  Recalcular Totais do Lote"
         Height          =   810
         Left            =   885
         Picture         =   "Lote.ctx":052C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   780
         Width           =   1305
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Número de Itens:"
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
         Left            =   1065
         TabIndex        =   16
         Top             =   330
         Width           =   1470
      End
      Begin VB.Label NumItensAtual 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   300
         Left            =   2595
         TabIndex        =   15
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   2790
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Lote.ctx":069E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Lote.ctx":07F8
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Lote.ctx":0982
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Lote.ctx":0EB4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Lote 
      Height          =   315
      Left            =   1665
      TabIndex        =   0
      Top             =   345
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NumItensInf 
      Height          =   315
      Left            =   1665
      TabIndex        =   3
      Top             =   1395
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0"
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1665
      TabIndex        =   2
      Top             =   870
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin VB.Label LoteLbl 
      AutoSize        =   -1  'True
      Caption         =   "Lote:"
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
      Left            =   1170
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   13
      Top             =   390
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Número de Itens:"
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
      Left            =   165
      TabIndex        =   12
      Top             =   1440
      Width           =   1470
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   705
      TabIndex        =   11
      Top             =   930
      Width           =   915
   End
End
Attribute VB_Name = "LoteEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1

Private Sub BotaoItens_Click()

Dim colSelecao As New Collection
Dim objInventario As New ClassInventario
Dim objInvLote As New ClassInvLote
Dim lErro As Long

On Error GoTo Erro_BotaoItens_Click

    If Len(Trim(Lote.ClipText)) = 0 Then Error 59627
                
    objInvLote.iLote = CInt(Lote.Text)
    objInvLote.iFilialEmpresa = giFilialEmpresa
    
    'le para verificar se o lote está cadastrado
    lErro = CF("InvLotePendente_Le",objInvLote)
    If lErro <> SUCESSO And lErro <> 41181 Then Error 41266
    
    If lErro = 41181 Then Error 52156
    
    colSelecao.Add objInvLote.iLote
    
    Call Chama_Tela("InventarioLoteLista_Lote", colSelecao, objInventario)
        
    Exit Sub
    
Erro_BotaoItens_Click:

    Select Case Err
        
        Case 59627
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_PREENCHIDO", Err)
            
        Case 52156
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTEINVPEN_NAO_CADASTRADO", Err, objInvLote.iLote)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162427)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lLote As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 41267

    Call Limpa_LoteTela

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 41267

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162428)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lLote As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo lote disponível
    lErro = CF("InvLote_Automatico",lLote)
    If lErro <> SUCESSO Then Error 57520

    Lote.Text = CStr(lLote)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57520
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162429)
    
    End Select

    Exit Sub

End Sub
    
Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Lote_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Lote_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Lote, iAlterado)

End Sub

Private Sub LoteLbl_Click()

Dim objInvLote As New ClassInvLote
Dim colSelecao As New Collection

    Call Move_Tela_Memoria(objInvLote)

    Call Chama_Tela("InvLotePendenteLista", colSelecao, objInvLote, objEventoLote)

End Sub

Private Sub NumItensInf_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumItensInf_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumItensInf, iAlterado)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)

Dim objInvLote As ClassInvLote
Dim lErro As Long

On Error GoTo Erro_objEventoLote_evSelecao

    Set objInvLote = obj1

    lErro = Traz_InvLote_Tela(objInvLote)
    If lErro <> SUCESSO Then Error 41238

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoLote_evSelecao:

    Select Case Err

        Case 41238

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162430)

    End Select

    Exit Sub

End Sub

Private Sub Lote_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objInvLote As New ClassInvLote
Dim iLoteAtualizado As Integer

On Error GoTo Erro_Lote_Validate

    'se o número do lote não foi fornecido, nao critica
    If Len(Trim(Lote.ClipText)) = 0 Then Exit Sub
    If CInt(Lote.Text) = 0 Then Exit Sub

    'carrega em memória os dados da tela
    objInvLote.iFilialEmpresa = giFilialEmpresa
    objInvLote.iLote = CInt(Lote.Text)

    'verifica se o lote  está atualizado
    lErro = CF("InvLote_Critica_Atualizado",objInvLote, iLoteAtualizado)
    If lErro <> SUCESSO Then Error 55461
    
    If iLoteAtualizado = LOTE_ATUALIZADO Then Error 55462

    Exit Sub

Erro_Lote_Validate:

    Cancel = True


    Select Case Err

        Case 55461
        
        Case 55462
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INVLOTE_ATUALIZADO", Err, objInvLote.iFilialEmpresa, objInvLote.iLote)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162431)

    End Select

    Exit Sub

End Sub

Private Sub BotaoRecalcular_Click()

Dim lErro As Long
Dim objInvLotePendente As New ClassInvLote
Dim iLote  As Integer
Dim iNumIguais As Integer
Dim iNumItensAtualizado As Integer
Dim vbMesRes As VbMsgBoxResult

On Error GoTo Erro_Botao_Recalcular_Click

    'Carrega em memória os dados da tela
    If Len(Trim(Lote.ClipText)) = 0 Then Error 41264

    objInvLotePendente.iLote = CInt(Lote.Text)
    objInvLotePendente.iNumItensAtual = CInt(NumItensAtual.Caption)
        
    lErro = CF("InventarioPendente_Critica_Lote",objInvLotePendente, iNumIguais)
    If lErro <> SUCESSO Then Error 41265
    
    Select Case iNumIguais
    
        Case IGUAL
                    
            'para totais iguais dá uma mensagem
            vbMesRes = Rotina_Aviso(vbOKOnly, "AVISO_IGUALDADE_TOTAIS2")
        
        Case DIFERENTE
        
            vbMesRes = Rotina_Aviso(vbYesNo, "AVISO_ATUALIZACAO_LOTE2", objInvLotePendente.iLote, objInvLotePendente.iNumItensAtual)
                    
            If vbMesRes = vbYes Then
                
                lErro = CF("InvLotePendente_Atualiza1",objInvLotePendente)
                If lErro <> SUCESSO Then Error 52177
                
                NumItensAtual.Caption = Format(objInvLotePendente.iNumItensAtual, "##,##0")
            
            End If
                        
    End Select
    
    Exit Sub

Erro_Botao_Recalcular_Click:

    Select Case Err

        Case 41264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", Err)
    
        Case 41265, 52177
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162432)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objInvLote As New ClassInvLote
Dim lLote As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'se o número do lote não foi fornecido ==> erro
    If Len(Trim(Lote.ClipText)) = 0 Then Error 41269

    'carrega em memória os dados da tela
    objInvLote.iFilialEmpresa = giFilialEmpresa
    objInvLote.iLote = CInt(Lote.Text)

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_LOTE")

    If vbMsgRes = vbYes Then

        'exclui o lote do banco de dados
        lErro = CF("InvLotePendente_Exclui",objInvLote)
        If lErro <> SUCESSO Then Error 41270

        'limpa o conteudo dos campos da tela
        Call Limpa_LoteTela

        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 41269
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", Err)
            Lote.SetFocus

        Case 41270
            Lote.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162433)

    End Select

    Exit Sub

End Sub

Private Sub Move_Tela_Memoria(objInvLote As ClassInvLote)
'move os dados da tela para objInvLote

    'carrega em memória os dados da tela
    If Len(Trim(Lote.ClipText)) = 0 Then
        objInvLote.iLote = 0
    Else
        objInvLote.iLote = CInt(Lote.Text)
    End If

    objInvLote.iFilialEmpresa = giFilialEmpresa
    objInvLote.sDescricao = Descricao.Text
    If Len(Trim(NumItensInf)) > 0 Then
        objInvLote.iNumItensInf = CInt(NumItensInf.Text)
    Else
        objInvLote.iNumItensInf = 0
    End If

End Sub

Public Function Trata_Parametros(Optional objInvLote As ClassInvLote) As Long

Dim lErro As Long
Dim lLote As Long

On Error GoTo Erro_Trata_Paramentros

    'Se foi passado um lote como parametro, exibir seus dados
    If Not (objInvLote Is Nothing) Then
        
        lErro = Traz_InvLote_Tela(objInvLote)
        If lErro <> SUCESSO Then Error 41193
    
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Paramentros:

    Trata_Parametros = Err

    Select Case Err

         Case 41193

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162434)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Sub Limpa_LoteTela()

Dim lErro As Long

    Call Limpa_Tela(Me)

    NumItensAtual.Caption = "0"
    
    Lote.Text = ""
    
End Sub

Function Traz_InvLote_Tela(objInvLote As ClassInvLote) As Long

Dim iLoteAtualizado As Integer
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Traz_InvLote_Tela

    Call Limpa_LoteTela

    Lote.Text = CStr(objInvLote.iLote)

    'verifica se o lote  está atualizado
    lErro = CF("InvLote_Critica_Atualizado",objInvLote, iLoteAtualizado)
    If lErro <> SUCESSO Then Error 41237

    'Se é um lote que já foi contabilizado, não pode sofrer alteração
    If iLoteAtualizado = LOTE_ATUALIZADO Then Error 41263

    lErro = CF("InvLotePendente_Le",objInvLote)
    If lErro <> SUCESSO And lErro <> 41181 Then Error 41266

    'se o lote está cadastrado, coloca o restante das informações na tela
    If lErro = SUCESSO Then

        Descricao.Text = objInvLote.sDescricao
        NumItensInf.Text = CStr(objInvLote.iNumItensInf)
        NumItensAtual.Caption = Format(objInvLote.iNumItensAtual, "##,##0")

    End If

    Traz_InvLote_Tela = SUCESSO

    Exit Function

Erro_Traz_InvLote_Tela:

    Traz_InvLote_Tela = Err

    Select Case Err

        Case 41237, 41266

        Case 41263
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INVLOTE_ATUALIZADO", Err, objInvLote.iFilialEmpresa, objInvLote.iLote)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162435)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objInvLote As New ClassInvLote

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    Call Move_Tela_Memoria(objInvLote)
    
    If objInvLote.iLote = 0 Then Error 41296
    
    lErro = Trata_Alteracao(objInvLote, objInvLote.iFilialEmpresa, objInvLote.iLote)
    If lErro <> SUCESSO Then Error 32281
        
    lErro = CF("InvLotePendente_Grava",objInvLote)
    If lErro <> SUCESSO Then Error 41297
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 32281
    
        Case 41296
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", Err)
            
        Case 41297

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162436)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim lLote As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 41284

    'limpa o conteudo dos campos da tela
    Call Limpa_LoteTela

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 41284

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162437)

    End Select

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objInvLote As New ClassInvLote
Dim colLancamento_Detalhe As New Collection

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "InvLotePendente"

    If Len(Trim(Lote.ClipText)) > 0 Then

        objInvLote.iLote = Lote.Text
        objInvLote.iFilialEmpresa = giFilialEmpresa

    End If

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", objInvLote.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Lote", objInvLote.iLote, 0, "Lote"
    colCampoValor.Add "Descricao", objInvLote.sDescricao, STRING_INVLOTE_DESCRICAO, "Descricao"

    'Exemplo de Filtro para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162438)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objInvLote As New ClassInvLote

On Error GoTo Erro_Tela_Preenche

    objInvLote.iLote = colCampoValor.Item("Lote").vValor

    If objInvLote.iLote <> 0 Then

        objInvLote.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor

        lErro = Traz_InvLote_Tela(objInvLote)
        If lErro <> SUCESSO Then Error 41235

        iAlterado = 0

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 41235

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162439)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
     
    Set objEventoLote = Nothing

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim colExerciciosAbertos As New Collection
Dim objExercicio As ClassExercicio

On Error GoTo Erro_Form_Load

    Set objEventoLote = New AdmEvento
    
    NumItensAtual.Caption = "0"
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162440)

    End Select

    iAlterado = 0
    
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

    Parent.HelpContextID = IDH_LOTE_EST
    Set Form_Load_Ocx = Me
    Caption = "Lote de Inventário"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LoteEst"
    
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
        If Me.ActiveControl Is Lote Then
            Call LoteLbl_Click
        End If
    End If

End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub NumItensAtual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumItensAtual, Source, X, Y)
End Sub

Private Sub NumItensAtual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumItensAtual, Button, Shift, X, Y)
End Sub

Private Sub LoteLbl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LoteLbl, Source, X, Y)
End Sub

Private Sub LoteLbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LoteLbl, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

