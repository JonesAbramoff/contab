VERSION 5.00
Begin VB.UserControl ContaCcl2Ocx 
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8880
   LockControls    =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   8880
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7080
      ScaleHeight     =   495
      ScaleWidth      =   1605
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   1665
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "ContaCcl2Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "ContaCcl2Ocx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "ContaCcl2Ocx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoContasAssociadas 
      Caption         =   "Contas Associadas"
      Enabled         =   0   'False
      Height          =   900
      Left            =   7500
      Picture         =   "ContaCcl2Ocx.ctx":080A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   945
      Width           =   1155
   End
   Begin VB.CommandButton BotaoCclAssociados 
      Caption         =   "Centros C. Associados"
      Enabled         =   0   'False
      Height          =   900
      Left            =   3105
      Picture         =   "ContaCcl2Ocx.ctx":1274
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   945
      Width           =   1155
   End
   Begin VB.CommandButton BotaoMarcarCcl 
      Caption         =   "Marcar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   7500
      Picture         =   "ContaCcl2Ocx.ctx":1CDE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1965
      Width           =   1155
   End
   Begin VB.CommandButton BotaoMarTodosCcl 
      Caption         =   "Mar.Todos"
      Enabled         =   0   'False
      Height          =   555
      Left            =   7500
      Picture         =   "ContaCcl2Ocx.ctx":2408
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2670
      Width           =   1155
   End
   Begin VB.CommandButton BotaoDesmarcarCcl 
      Caption         =   "Desmarcar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   7500
      Picture         =   "ContaCcl2Ocx.ctx":3422
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3375
      Width           =   1155
   End
   Begin VB.CommandButton BotaoDesTodosCcl 
      Caption         =   "Des.Todos"
      Enabled         =   0   'False
      Height          =   555
      Left            =   7485
      Picture         =   "ContaCcl2Ocx.ctx":3B24
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4065
      Width           =   1155
   End
   Begin VB.ListBox Contas 
      Height          =   3660
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   900
      Width           =   2775
   End
   Begin VB.CommandButton BotaoDesTodosConta 
      Caption         =   "Des.Todos"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3105
      Picture         =   "ContaCcl2Ocx.ctx":4D06
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4065
      Width           =   1155
   End
   Begin VB.CommandButton BotaoDesmarcarConta 
      Caption         =   "Desmarcar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3105
      Picture         =   "ContaCcl2Ocx.ctx":5EE8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3375
      Width           =   1155
   End
   Begin VB.CommandButton BotaoMarTodosConta 
      Caption         =   "Mar.Todos"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3120
      Picture         =   "ContaCcl2Ocx.ctx":65EA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2670
      Width           =   1155
   End
   Begin VB.CommandButton BotaoMarcarConta 
      Caption         =   "Marcar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3120
      Picture         =   "ContaCcl2Ocx.ctx":7604
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1995
      Width           =   1155
   End
   Begin VB.ListBox Ccl 
      Height          =   3660
      ItemData        =   "ContaCcl2Ocx.ctx":7D2E
      Left            =   4440
      List            =   "ContaCcl2Ocx.ctx":7D35
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label2 
      Height          =   195
      Left            =   4560
      Top             =   720
      Width           =   1470
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
      Caption         =   "Centros de Custo"
   End
   Begin VB.Label Label7 
      Height          =   195
      Left            =   120
      Top             =   660
      Width           =   1500
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
      Caption         =   "Contas Sintéticas"
   End
End
Attribute VB_Name = "ContaCcl2Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Centro de Custo/Lucro Contábil

Public Sub Desmarca_ListBox_Ccl()

Dim iIndice As Integer

    For iIndice = 0 To Ccl.ListCount - 1
        Ccl.Selected(iIndice) = False
    Next
    
    Ccl.ListIndex = -1
    
End Sub

Public Sub Desmarca_ListBox_Contas()

Dim iIndice As Integer

    For iIndice = 0 To Contas.ListCount - 1
        Contas.Selected(iIndice) = False
    Next
    
    Contas.ListIndex = -1
    
End Sub

Private Sub BotaoCclAssociados_Click()
'mostra os centros de custo/lucro associados à conta selecionada /m

Dim lErro As Long
Dim iIndice As Integer
Dim colCcl As New Collection
Dim vCcl As Variant
Dim sCcl1 As String
Dim sConta As String
Dim iContaPreenchida As Integer
    
On Error GoTo Erro_BotaoCclAssociados_Click
    
    'verifica se alguma conta foi selecionada
    If Len(Contas.Text) = 0 Then Error 9877
    
    Call Desmarca_ListBox_Ccl
    
    lErro = CF("Conta_Formata",Left(Contas.Text, InStr(Contas.Text, SEPARADOR) - 2), sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 8264
    
    lErro = CF("ContaCcl_Associa_Ccl",colCcl, sConta)
    If lErro <> SUCESSO Then Error 9923
    
    For Each vCcl In colCcl
    
        sCcl1 = String(STRING_CCL, 0)
        
        lErro = Mascara_MascararCcl(CStr(vCcl), sCcl1)
        If lErro <> SUCESSO Then Error 8260
    
        For iIndice = 0 To Ccl.ListCount - 1
            If Left(Ccl.List(iIndice), InStr(Ccl.List(iIndice), SEPARADOR) - 2) = sCcl1 Then
                Ccl.Selected(iIndice) = True
                Exit For
            End If
        Next
        
    Next
    
    Ccl.ListIndex = -1
    
    Exit Sub
    
Erro_BotaoCclAssociados_Click:

    Select Case Err
    
        Case 8260
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARAR_CCL", Err, CStr(vCcl))

        Case 8264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATAR_CONTA", Err, Contas.Text)
            
        Case 9877
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_SELECIONADA", Err)
            
        Case 9923

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 154954)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoContasAssociadas_Click()
'mostra as contas associadas ao centro de custo selecionado /m

Dim lErro As Long
Dim iIndice As Integer
Dim colContas As New Collection
Dim vConta As Variant
Dim sCcl As String
Dim sConta1 As String
    
On Error GoTo Erro_BotaoContasAssociadas_Click
        
    'verifica se há algum centro de custo selecionado
    If Len(Ccl.Text) = 0 Then Error 9878

    Call Desmarca_ListBox_Contas
    
    lErro = CF("Ccl_Formata",Left(Ccl.Text, InStr(Ccl.Text, SEPARADOR) - 2), sCcl, 1)
    If lErro <> SUCESSO Then Error 8265
    
    lErro = CF("ContaCcl_Associa_Conta",colContas, sCcl)
    If lErro <> SUCESSO Then Error 9924
    
    For Each vConta In colContas
    
        sConta1 = String(STRING_CONTA, 0)
        
        lErro = Mascara_MascararConta(CStr(vConta), sConta1)
        If lErro <> SUCESSO Then Error 8261
    
        For iIndice = 0 To Contas.ListCount - 1
            If Left(Contas.List(iIndice), InStr(Contas.List(iIndice), SEPARADOR) - 2) = sConta1 Then
                Contas.Selected(iIndice) = True
                Exit For
            End If
        Next
        
    Next
    
    Contas.ListIndex = -1
    
    Exit Sub
    
Erro_BotaoContasAssociadas_Click:

    Select Case Err
    
        Case 8261
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARAR_CONTA", Err, CStr(vConta))

        Case 8265
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATAR_CCL", Err, Ccl.Text)

        Case 9878
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_SELECIONADA", Err)

        Case 9924
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 154955)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoDesmarcarCcl_Click()

Dim iIndex As Integer

    If Ccl.ListIndex = -1 Then Exit Sub

        iIndex = Ccl.ListIndex

        Ccl.Selected(iIndex) = False

End Sub

Private Sub BotaoDesmarcarConta_Click()

Dim iIndex As Integer

    If Contas.ListIndex = -1 Then Exit Sub
    
        iIndex = Contas.ListIndex

        Contas.Selected(iIndex) = False

End Sub

Private Sub BotaoDesTodosCcl_Click()

    Call Desmarca_ListBox_Ccl
    
End Sub

Private Sub BotaoDesTodosConta_Click()

    Call Desmarca_ListBox_Contas
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim iIndex As Integer
Dim vbMsgRet As VbMsgBoxResult
Dim colContas As New Collection
Dim colCcl As New Collection
Dim sConta As String
Dim sCcl As String
Dim iContaPreenchida As Integer

On Error GoTo Erro_BotaoGravar_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Coloca em colContas os elementos marcados na ListBox Contas
    For iIndex = 0 To Contas.ListCount - 1
        
        If Contas.Selected(iIndex) = True Then
        
            'coloca a conta selecionada no formato do bd
            lErro = CF("Conta_Formata",Left(Contas.List(iIndex), InStr(Contas.List(iIndex), SEPARADOR) - 2), sConta, iContaPreenchida)
            If lErro <> SUCESSO Then Error 9625
            
            colContas.Add sConta
            
        End If
               
    Next
    
    'se nenhuma conta estiver selecionada ==> erro
    If colContas.Count = 0 Then Error 9930

    'Coloca em colCcl os elementos marcados na ListBox Ccl
    For iIndex = 0 To Ccl.ListCount - 1
        
        If Ccl.Selected(iIndex) = True Then
        
            'coloca o centro de custo/lucro selecionado no formato do bd
            lErro = CF("Ccl_Formata",Left(Ccl.List(iIndex), InStr(Ccl.List(iIndex), SEPARADOR) - 2), sCcl, 1)
            If lErro <> SUCESSO Then Error 8263

            colCcl.Add sCcl
            
        End If
               
    Next
    
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_CONTACCL")
    
    If vbMsgRet = vbYes Then
    
        'atualiza as associacoes de conta com centro de custo/lucro no banco de dados
        lErro = CF("ContaCcl_Atualizacao_Contabil",colContas, colCcl)
        If lErro <> SUCESSO Then Error 8202
        
        'limpa a tela
        Desmarca_ListBox_Contas
        Desmarca_ListBox_Ccl
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoGravar_Click:
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 8202
        
        Case 8263
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATAR_CCL", Err, Ccl.List(iIndex))
    
        Case 9625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATAR_CONTA", Err, Contas.List(iIndex))
            
        Case 9930
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_SELECIONADA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 154956)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()

    Call Desmarca_ListBox_Contas
    
    Call Desmarca_ListBox_Ccl

End Sub

Private Sub BotaoMarcarCcl_Click()

Dim iIndex As Integer

    If Ccl.ListIndex = -1 Then Exit Sub
    
        iIndex = Ccl.ListIndex

        Ccl.Selected(iIndex) = True

End Sub

Private Sub BotaoMarcarConta_Click()

Dim iIndex As Integer

    If Contas.ListIndex = -1 Then Exit Sub
    
        iIndex = Contas.ListIndex

        Contas.Selected(iIndex) = True

End Sub

Private Sub BotaoMarTodosCcl_Click()

Dim iIndice As Integer

    For iIndice = 0 To Ccl.ListCount - 1
        Ccl.Selected(iIndice) = True
    Next

    Ccl.ListIndex = -1
    
End Sub

Private Sub BotaoMarTodosConta_Click()

Dim iIndice As Integer

    For iIndice = 0 To Contas.ListCount - 1
        Contas.Selected(iIndice) = True
    Next
    
    Contas.ListIndex = -1

End Sub

Private Sub Ccl_DblClick()

    Call BotaoContasAssociadas_Click

End Sub

Private Sub Contas_DblClick()

    Call BotaoCclAssociados_Click

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colContas As New Collection
Dim colCcl As New Collection
Dim objPlanoConta As ClassPlanoConta
Dim sCcl1 As String
Dim iIndice As Integer
Dim sConta1 As String
Dim objCcl As ClassCcl

On Error GoTo Erro_Form_Load

    Contas.Clear
   
    lErro = CF("ContaCcl_Le_Todos_Conta",colContas, CONTA_SINTETICA)
    If lErro <> SUCESSO And lErro <> 8127 Then Error 8102
        
    If lErro = 8127 Then Error 9920
        
    Label7.Caption = "Contas Sintéticas"

    'Ativar botões para conta
    BotaoMarcarConta.Enabled = True
    BotaoMarTodosConta.Enabled = True
    BotaoCclAssociados.Enabled = True
    BotaoDesmarcarConta.Enabled = True
    BotaoDesTodosConta.Enabled = True
        
    For Each objPlanoConta In colContas
    
        sConta1 = String(STRING_CONTA, 0)
        
        lErro = Mascara_MascararConta(objPlanoConta.sConta, sConta1)
        If lErro <> SUCESSO Then Error 8256
        
        Contas.AddItem sConta1 & " " & SEPARADOR & " " & objPlanoConta.sDescConta
        
    Next
            
    Contas.ListIndex = -1
    
    'Carrega Centros de Custo analiticos
    lErro = CF("Ccl_Le_Todos_Analiticos",colCcl)
    If lErro <> SUCESSO Then Error 8103
    
    If colCcl.Count = 0 Then Error 9921
    
    Ccl.Clear
    
    For Each objCcl In colCcl
    
        sCcl1 = String(STRING_CCL, 0)
        
        lErro = Mascara_MascararCcl(objCcl.sCcl, sCcl1)
        If lErro <> SUCESSO Then Error 8257
        
        Ccl.AddItem sCcl1 & " " & SEPARADOR & " " & objCcl.sDescCcl
        
    Next
    
    Ccl.ListIndex = -1
    
    'Ativar botões para Ccl
    BotaoMarcarCcl.Enabled = True
    BotaoMarTodosCcl.Enabled = True
    BotaoContasAssociadas.Enabled = True
    BotaoDesmarcarCcl.Enabled = True
    BotaoDesTodosCcl.Enabled = True
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 8102, 8103
            
        Case 8256
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARAR_CONTA", Err, objPlanoConta.sConta)
        
        Case 8257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARAR_CCL", Err, objCcl.sCcl)
            
        Case 9920
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PLANOCONTA_SEM_CONTA_SINTETICA", Err)
        
        Case 9921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_VAZIO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 154957)
    
    End Select
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objContaCcl As ClassContaCcl) As Long

Dim lErro As Long
Dim sConta1 As String
Dim colCcl As New Collection
Dim iIndice As Integer
Dim sCcl As String
Dim sCcl1 As String
Dim vCcl As Variant

On Error GoTo Erro_Trata_Parametros

    'Verifica se sConta está preenchida
    If Not (objContaCcl Is Nothing) Then
        
        'Verifica a existencia da conta
        lErro = CF("PlanoConta_Le_Conta",objContaCcl.sConta)
        If lErro <> SUCESSO And lErro <> 10051 Then Error 10053
        
        'se a conta não estiver cadastrada ==> erro
        If lErro = 10051 Then Error 8104
    
        sConta1 = String(STRING_CONTA, 0)
    
        lErro = Mascara_MascararConta(objContaCcl.sConta, sConta1)
        If lErro <> SUCESSO Then Error 8258
            
        'Marca a conta na ListBox Contas
        For iIndice = 0 To Contas.ListCount - 1
            If Contas.List(iIndice) = sConta1 Then
                Contas.Selected(iIndice) = True
                Exit For
            End If
        Next
    
        Call CF("ContaCcl_Associa_Ccl",colCcl, objContaCcl.sConta)
    
        For Each vCcl In colCcl
    
            sCcl1 = String(STRING_CCL, 0)
            
            lErro = Mascara_MascararCcl(CStr(vCcl), sCcl1)
            If lErro <> SUCESSO Then Error 8259
        
            For iIndice = 0 To Ccl.ListCount - 1
                If Ccl.List(iIndice) = sCcl1 Then
                    Ccl.Selected(iIndice) = True
                    Exit For
                End If
            Next
            
        Next
        
        Ccl.ListIndex = -1
    
    End If

    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 8104
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, objContaCcl.sConta)
            
        Case 8258
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARAR_CONTA", Err, objContaCcl.sConta)
        
        Case 8259
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARAR_CCL", Err, CStr(vCcl))
            
        Case 10053
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154958)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ASSOCIACAO_CONTA_CENTRO_CUSTO_LUCRO_CONTABIL
    Set Form_Load_Ocx = Me
    Caption = "Associação Conta x Centro de Custo/Lucro (Contábil)"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ContaCcl2"
    
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

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

