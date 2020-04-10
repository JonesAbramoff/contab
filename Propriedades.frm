VERSION 5.00
Begin VB.Form Propriedades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propriedades"
   ClientHeight    =   5130
   ClientLeft      =   5790
   ClientTop       =   390
   ClientWidth     =   2670
   Icon            =   "Propriedades.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   2670
   Begin VB.CheckBox Habilitado 
      Caption         =   "Habilitado"
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
      Left            =   1230
      TabIndex        =   10
      Top             =   3210
      Width           =   1425
   End
   Begin VB.CommandButton botaoSalvar 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   75
      TabIndex        =   24
      Top             =   4710
      Width           =   2550
   End
   Begin VB.CommandButton BotaoTelaOriginal 
      Caption         =   "Tela Original"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   75
      TabIndex        =   1
      Top             =   4305
      Width           =   2550
   End
   Begin VB.CommandButton BotaoControleOriginal 
      Caption         =   "Controle Original"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   75
      TabIndex        =   3
      Top             =   3900
      Width           =   2550
   End
   Begin VB.CommandButton BotaoParaTras 
      Caption         =   "Para Trás"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1395
      TabIndex        =   5
      Top             =   3495
      Width           =   1245
   End
   Begin VB.CommandButton BotaoParaFrente 
      Caption         =   "Para Frente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   75
      TabIndex        =   7
      Top             =   3495
      Width           =   1245
   End
   Begin VB.TextBox Container 
      Height          =   300
      Left            =   1230
      TabIndex        =   9
      Top             =   2880
      Width           =   1410
   End
   Begin VB.ComboBox ComboTabStop 
      Height          =   315
      ItemData        =   "Propriedades.frx":014A
      Left            =   1230
      List            =   "Propriedades.frx":0154
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2520
      Width           =   1410
   End
   Begin VB.TextBox Titulo 
      Height          =   300
      Left            =   1230
      TabIndex        =   14
      Top             =   480
      Width           =   1410
   End
   Begin VB.TextBox TabIndex 
      Height          =   300
      Left            =   1230
      TabIndex        =   13
      Top             =   2190
      Width           =   1410
   End
   Begin VB.TextBox Topo 
      Height          =   300
      Left            =   1230
      TabIndex        =   11
      Top             =   1845
      Width           =   1410
   End
   Begin VB.TextBox Esquerda 
      Height          =   300
      Left            =   1230
      TabIndex        =   8
      Top             =   1500
      Width           =   1410
   End
   Begin VB.TextBox Altura 
      Height          =   300
      Left            =   1230
      TabIndex        =   6
      Top             =   1170
      Width           =   1410
   End
   Begin VB.TextBox Largura 
      Height          =   300
      Left            =   1230
      TabIndex        =   4
      Top             =   825
      Width           =   1410
   End
   Begin VB.ComboBox ComboVisivel 
      Height          =   315
      ItemData        =   "Propriedades.frx":0162
      Left            =   -20000
      List            =   "Propriedades.frx":016C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   819
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ComboBox ComboCampos 
      Height          =   315
      Left            =   45
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   2610
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Container:"
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
      Left            =   150
      TabIndex        =   15
      Top             =   2925
      Width           =   1020
   End
   Begin VB.Label LabelTabStop 
      Alignment       =   1  'Right Justify
      Caption         =   "TabStop:"
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
      Left            =   60
      TabIndex        =   16
      Top             =   2595
      Width           =   1095
   End
   Begin VB.Label LabelTitulo 
      Alignment       =   1  'Right Justify
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
      Height          =   210
      Left            =   135
      TabIndex        =   17
      Top             =   525
      Width           =   1020
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "TabIndex:"
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
      Left            =   135
      TabIndex        =   18
      Top             =   2205
      Width           =   1020
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Topo:"
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
      Left            =   135
      TabIndex        =   19
      Top             =   1860
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Esquerda:"
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
      Left            =   135
      TabIndex        =   20
      Top             =   1515
      Width           =   1020
   End
   Begin VB.Label LabelAltura 
      Alignment       =   1  'Right Justify
      Caption         =   "Altura:"
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
      Left            =   120
      TabIndex        =   21
      Top             =   1185
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Largura:"
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
      Left            =   120
      TabIndex        =   22
      Top             =   855
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Visível:"
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
      Left            =   -20000
      TabIndex        =   23
      Top             =   885
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Propriedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objControle1 As Object
Dim gobjIncluido As ClassEdicaoTela_Tela 'Inserido por Wagner

Public Sub ComboCampos_Seleciona(objControle As Object)

Dim iIndice As Integer
Dim Controle As Object
Dim sNome As String
Dim sNome1 As String

On Error GoTo Erro_ComboCampos_Seleciona

    ComboCampos.Clear

    For Each Controle In gobjTelaAtiva.Controls
        If Not (TypeName(Controle) = "Menu") And Not (TypeName(Controle) = "Timer") And Not (TypeName(Controle) = "Line") And Not (TypeName(Controle) = "CommonDialog") And Not (TypeName(Controle) = "Image") Then
            iIndice = -1
            sNome = Controle.Name
            iIndice = Controle.Index
            If iIndice > -1 Then sNome = sNome & "(" & CStr(Controle.Index) & ")"
            ComboCampos.AddItem sNome
            ComboCampos.ItemData(ComboCampos.NewIndex) = iIndice
        End If
    Next
    
    Call CamposInvisiveis.Carrega_Campos_Invisiveis
    
    sNome = objControle.Name
    
    If objControle.Index > -1 Then sNome = sNome & "(" & CStr(objControle.Index) & ")"
    
    For iIndice = 0 To ComboCampos.ListCount - 1
    
        If ComboCampos.List(iIndice) = sNome Then
                            
            Set objControle1 = objControle
            ComboCampos.ListIndex = -1
            ComboCampos.ListIndex = iIndice
            Exit For
            
        End If
    Next
    
'    If ComboCampos.ListIndex <> -1 Then
'        Adiciona_Controle
'    End If
    
    Exit Sub

Erro_ComboCampos_Seleciona:

    Select Case Err
    
        Case 343
            Resume Next
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 165861)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoControleOriginal_Click()

Dim objEdicaoTela_Tela As ClassEdicaoTela_Tela
Dim objEdicaoTela_Controle As ClassEdicaoTela_Controle
Dim objControle2 As Object
Dim iIndice As Integer
Dim iAchou As Integer
Dim bEncontrou As Boolean
Dim sControle As String
Dim iIndice2 As Integer

On Error GoTo Erro_BotaoControleOriginal_Click
    
    bEncontrou = False
    
    If (objControle1 Is Nothing) Then Exit Sub
    
    iIndice2 = -1
    iIndice2 = objControle1.Index

    If iIndice2 > -1 Then
        sControle = objControle1.Name & "(" & objControle1.Index & ")"
    Else
        sControle = objControle1.Name
    End If
    
    'Procura na colecao
    For Each objEdicaoTela_Tela In gcolEdicaoTela
        If objEdicaoTela_Tela.sNomeTela = gobjTelaAtiva.Name Then
            
            For Each objEdicaoTela_Controle In objEdicaoTela_Tela.colEdicaoTela_Controle
                
                If objEdicaoTela_Controle.sNomeControle = sControle Then
                                    
                    bEncontrou = True
                    
                    If objEdicaoTela_Controle.iAlturaPadrao = -1 Then
                        Altura.Text = ""
                    Else
                        objControle1.Height = objEdicaoTela_Controle.iAlturaPadrao
                        Altura.Text = CStr(objEdicaoTela_Controle.iAlturaPadrao)
                    End If
                    
                    
                    objControle1.TabIndex = objEdicaoTela_Controle.iTabIndexPadrao
                    TabIndex.Text = CStr(objEdicaoTela_Controle.iTabIndexPadrao)
                    
                    If (TypeName(objControle1) = "Label") Or (TypeName(objControle1) = "CommandButton") Or (TypeName(objControle1) = "OptionButton") Or (TypeName(objControle1) = "Frame") Or (TypeName(objControle1) = "CheckBox") Then
                        objControle1.Caption = objEdicaoTela_Controle.sTituloPadrao
                        objEdicaoTela_Controle.sTitulo = objEdicaoTela_Controle.sTituloPadrao
                        Titulo.Text = objEdicaoTela_Controle.sTituloPadrao
                    End If
                    
                    objControle1.left = objEdicaoTela_Controle.iEsquerdaPadrao
                    Esquerda.Text = CStr(objEdicaoTela_Controle.iEsquerdaPadrao)
                    objControle1.Width = objEdicaoTela_Controle.iLarguraPadrao
                    Largura.Text = CStr(objEdicaoTela_Controle.iLarguraPadrao)
                    If objEdicaoTela_Controle.iTabStopPadrao <> -1 Then
                        If objEdicaoTela_Controle.iTabStopPadrao = 1 Then
                            objControle1.TabStop = True
                        Else
                            objControle1.TabStop = False
                        End If
                        
                        For iIndice = 0 To ComboTabStop.ListCount - 1
                            If ComboTabStop.ItemData(iIndice) = objEdicaoTela_Controle.iTabStopPadrao Then
                                ComboTabStop.ListIndex = iIndice
                                Exit For
                            End If
                        Next
                    End If
                    
                    objControle1.top = objEdicaoTela_Controle.iTopoPadrao
                    Topo.Text = CStr(objEdicaoTela_Controle.iTopoPadrao)
                    
                    If objEdicaoTela_Controle.iHabilitadoPadrao = MARCADO Then
                        objControle1.Enabled = True
                        Habilitado.Value = vbChecked
                    Else
                        objControle1.Enabled = False
                        Habilitado.Value = vbUnchecked
                    End If
                                
                    iAchou = 0
                                
                    For Each objControle2 In gobjTelaAtiva.Controls
                        If objControle2.Name = objEdicaoTela_Controle.sContainerPadrao Then
                            If objEdicaoTela_Controle.iIndiceContainerPadrao > -1 Then
                                If objControle2.Index = objEdicaoTela_Controle.iIndiceContainerPadrao Then
                                    Set objControle1.Container = objControle2
                                    iAchou = 1
                                    Container.Text = objControle2.Name & "(" & objControle2.Index & ")"
                                    Exit For
                                End If
                            Else
                                Set objControle1.Container = objControle2
                                iAchou = 1
                                Container.Text = objControle2.Name
                                Exit For
                            End If
                        End If
                    Next
    
                    If iAchou <> 1 Then
                        If objEdicaoTela_Controle.sContainerPadrao = gobjTelaAtiva.Name Then
                            Set objControle1.Container = gobjTelaAtiva
                            Container.Text = Me.Name
                            iAchou = 1
                        End If
                    End If
                    
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
    
    If bEncontrou = True Then
        Call Adiciona_Controle
    End If
    
    Exit Sub
    
Erro_BotaoControleOriginal_Click:
    
    Select Case Err
    
        Case 343
            Resume Next
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165862)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoTelaOriginal_Click()

Dim objEdicaoTela_Tela As ClassEdicaoTela_Tela
Dim objEdicaoTela_Controle As ClassEdicaoTela_Controle
Dim objControle As Object
Dim objControle2 As Object
Dim iAchou As Integer
Dim iIndice As Integer
Dim bEncontrouTela As Boolean
Dim iTabIndex As Integer
Dim iPosicaoPrimeiroEspaco As Integer
Dim iPosicaoSegundoEspaco As Integer
Dim sTabIndexBuffer As String
Dim colControles As New Collection
Dim sControle As String
Dim sControle2 As String
Dim sControle1 As String
Dim iIndice2 As Integer
Dim lErro As Long

'Inserido por Wagner
Dim colGrupoUsu As New Collection
Dim objUsu As New ClassUsuarios
Dim objGrupoUsu As ClassGrupoUsuarios
Dim bFlag As Boolean

On Error GoTo Erro_BotaoTelaOriginal_Click
    
    If Not (objControle1 Is Nothing) Then
        
        iIndice2 = -1
        iIndice2 = objControle1.Index
        
        If iIndice2 > -1 Then
            sControle1 = objControle1.Name & "(" & objControle1.Index & ")"
        Else
            sControle1 = objControle1.Name
        End If
    End If
    
    'Procura na colecao
    For Each objEdicaoTela_Tela In gcolEdicaoTela
        
        If objEdicaoTela_Tela.sNomeTela = gobjTelaAtiva.Name Then
            
            bEncontrouTela = True
            
            'Se tem Tab Index Padrão
            If Len(Trim(objEdicaoTela_Tela.sTabIndexPadrao)) > 0 Then
                
                'Atualiza o Buffer
                sTabIndexBuffer = objEdicaoTela_Tela.sTabIndexPadrao
                objEdicaoTela_Tela.sTabIndex = objEdicaoTela_Tela.sTabIndexPadrao
                
                'Limpa os objetos
                gobjEstInicial.List1.Clear
                Set colControles = New Collection
                    
                'Para cada Controle
                For Each objControle In gobjTelaAtiva.Controls
                    
                    iIndice2 = -1
                    iIndice2 = objControle.Index

                    If iIndice2 > -1 Then
                        sControle = objControle.Name & "(" & objControle.Index & ")"
                    Else
                        sControle = objControle.Name
                    End If
                    
'                    If (Not (TypeName(objControle)= "Label")) And (Not (TypeName(objControle)= "PictureBox")) And (Not (TypeName(objControle)= "Frame")) And (Not (TypeName(objControle)= "SSFrame")) And (Not (TypeName(objControle)= "SSPanel")) And (Not (TypeName(objControle)= "Line")) Then
                    If (Not (TypeName(objControle) = "Label")) And (Not (TypeName(objControle) = "PictureBox")) And (Not (TypeName(objControle) = "Frame")) And (Not (TypeName(objControle) = "Line")) And Not (TypeName(objControle) = "CommonDialog") And Not (TypeName(objControle) = "Image") Then
                        If Len(Trim(sTabIndexBuffer)) > 0 Then
                            
                            'Pega o Tab Index
                            iPosicaoPrimeiroEspaco = InStr(1, sTabIndexBuffer, " ")
                            iPosicaoSegundoEspaco = InStr(2, sTabIndexBuffer, " ")
                            If iPosicaoSegundoEspaco = 0 Then iPosicaoSegundoEspaco = Len(sTabIndexBuffer) + 1
                            
                            iTabIndex = CInt(Trim(Mid(sTabIndexBuffer, 2, (iPosicaoSegundoEspaco - iPosicaoPrimeiroEspaco) - 1)))

                            sTabIndexBuffer = right(sTabIndexBuffer, Len(sTabIndexBuffer) - (iPosicaoSegundoEspaco - iPosicaoPrimeiroEspaco))
                        
                            'Guarda o controle
                            colControles.Add objControle
                        
                            'Adiciona o Tab Index na List e a Posicao do controle no ItemData da Lista
                            Call gobjEstInicial.List1.AddItem(iTabIndex)
                            gobjEstInicial.List1.ItemData(gobjEstInicial.List1.NewIndex) = colControles.Count
                        
                            'Se o controle estiver selecionado
                            If Not (objControle1 Is Nothing) Then
                                If sControle1 = sControle Then
                                    'Atualiza a Tela de Propriedades
                                    TabIndex.Text = iTabIndex
                                End If
                            End If
                        End If
                    End If
                Next
                
                'Varre a lista na ordem de Tab Index atualizando os tab  Index dos Controles
                For iIndice = 0 To gobjEstInicial.List1.ListCount - 1
                    colControles.Item(gobjEstInicial.List1.ItemData(iIndice)).TabIndex = gobjEstInicial.List1.List(iIndice)
                Next
                
            End If
            
            For Each objEdicaoTela_Controle In objEdicaoTela_Tela.colEdicaoTela_Controle
                For Each objControle In gobjTelaAtiva.Controls
                    
                    iIndice2 = -1
                    iIndice2 = objControle.Index

                    If iIndice2 > -1 Then
                        sControle = objControle.Name & "(" & objControle.Index & ")"
                    Else
                        sControle = objControle.Name
                    End If
                    
                    If sControle = objEdicaoTela_Controle.sNomeControle Then
                        
                        If objEdicaoTela_Controle.iAlturaPadrao <> -1 Then
                            objControle.Height = objEdicaoTela_Controle.iAlturaPadrao
                        End If
                        
                        objControle.left = objEdicaoTela_Controle.iEsquerdaPadrao
                        objControle.Width = objEdicaoTela_Controle.iLarguraPadrao
                        If objEdicaoTela_Controle.iTabStopPadrao = 1 Then
                            objControle.TabStop = objEdicaoTela_Controle.iTabStopPadrao
                        End If
                        objControle.top = objEdicaoTela_Controle.iTopoPadrao
'                        objControle.Visible = objEdicaoTela_Controle.iVisivelPadrao
                        
                        If objEdicaoTela_Controle.iHabilitadoPadrao = MARCADO Then
                            objControle.Enabled = True
                        Else
                            objControle.Enabled = False
                        End If
                                
                        If (TypeName(objControle) = "Label") Or (TypeName(objControle) = "CommandButton") Or (TypeName(objControle) = "OptionButton") Or (TypeName(objControle) = "Frame") Or (TypeName(objControle) = "CheckBox") Then
                            objControle.Caption = objEdicaoTela_Controle.sTituloPadrao
                            objEdicaoTela_Controle.sTitulo = objEdicaoTela_Controle.sTituloPadrao
                        End If
                        
                        iAchou = 0
                        
                        For Each objControle2 In gobjTelaAtiva.Controls
                            
                            iIndice2 = -1
                            iIndice2 = objControle2.Index

                            sControle2 = objControle2.Name
                            
                            If sControle2 = objEdicaoTela_Controle.sContainerPadrao Then
                                If objEdicaoTela_Controle.iIndiceContainerPadrao > -1 Then
                                    If objControle2.Index = objEdicaoTela_Controle.iIndiceContainerPadrao Then
                                        If Not (objControle1 Is Nothing) Then
                                            If sControle1 = sControle2 Then
                                                Container.Text = objControle2.Name & "(" & objControle2.Index & ")"
                                            End If
                                        End If
                                        Set objControle.Container = objControle2
                                        iAchou = 1
                                        Exit For
                                    End If
                                Else
                                    Set objControle.Container = objControle2
                                    If Not (objControle1 Is Nothing) Then
                                        If sControle1 = sControle Then
                                            Container.Text = objControle2.Name
                                        End If
                                    End If
                                    iAchou = 1
                                    Exit For
                                End If
                            End If
                        Next
    
                        If iAchou <> 1 Then
                            If objEdicaoTela_Controle.sContainerPadrao = gobjTelaAtiva.Name Then
                                Set objControle.Container = gobjTelaAtiva
                                If Not (objControle1 Is Nothing) Then
                                    If sControle1 = sControle Then
                                        Container.Text = Me.Name
                                    End If
                                End If
                                iAchou = 1
                            End If
                        End If
                        
                        If Not (objControle1 Is Nothing) Then
                            'Trata o Controle Ativo para Atualizar os Campos da Tela
                            If sControle1 = sControle Then
                                
                                If objEdicaoTela_Controle.iAlturaPadrao = -1 Then
                                    Altura.Text = ""
                                Else
                                    Altura.Text = CStr(objEdicaoTela_Controle.iAlturaPadrao)
                                End If
                                
                                Esquerda.Text = CStr(objEdicaoTela_Controle.iEsquerdaPadrao)
                                Largura.Text = CStr(objEdicaoTela_Controle.iLarguraPadrao)
                                If objEdicaoTela_Controle.iTabStopPadrao <> -1 Then
                                    For iIndice = 0 To ComboTabStop.ListCount - 1
                                        If ComboTabStop.ItemData(iIndice) = objEdicaoTela_Controle.iTabStopPadrao Then
                                            ComboTabStop.ListIndex = iIndice
                                            Exit For
                                        End If
                                    Next
                                End If
                                
                                If (TypeName(objControle) = "Label") Or (TypeName(objControle) = "CommandButton") Or (TypeName(objControle) = "OptionButton") Or (TypeName(objControle) = "Frame") Or (TypeName(objControle) = "CheckBox") Then
                                    Titulo.Text = objEdicaoTela_Controle.sTituloPadrao
                                End If

                                Topo.Text = CStr(objEdicaoTela_Controle.iTopoPadrao)
                            
                                If objEdicaoTela_Controle.iHabilitadoPadrao = MARCADO Then
                                    Habilitado.Value = vbChecked
                                Else
                                    Habilitado.Value = vbUnchecked
                                End If
                            
                            End If
                        End If
                    Exit For
                    End If
                Next
            Next
            Exit For
        End If
    Next
    
    If bEncontrouTela = True Then
    
        Load GrupoUsuarios
    
        lErro = GrupoUsuarios.Trata_Parametros(colGrupoUsu)
        If lErro <> SUCESSO Then gError 129290
    
        GrupoUsuarios.Show vbModal
    
        lErro = CF("EdicaoTela_Exclui", gcolEdicaoTela.Item(gobjTelaAtiva.Name), colGrupoUsu) 'Inserido por Wagner
                
        objUsu.sCodUsuario = gsUsuario
        lErro = CF("Usuarios_Le", objUsu)
        If lErro <> SUCESSO Then gError 129310
        
        bFlag = False
        
        For Each objGrupoUsu In colGrupoUsu
        
            If objGrupoUsu.sCodGrupo = objUsu.sCodGrupo Then bFlag = True
        
        Next
        
        If bFlag Then Call gcolEdicaoTela.Remove(gobjTelaAtiva.Name)
    End If
    
    Exit Sub
    
Erro_BotaoTelaOriginal_Click:
    
    Select Case Err
    
        Case 343
            Resume Next
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165863)
    
    End Select
    
    Exit Sub

End Sub

Private Sub ComboCampos_Click()

Dim Controle As Object
Dim sNome As String
Dim iPos As Integer

On Error GoTo Erro_ComboCampos_Click

    If ComboCampos.ListIndex <> -1 Then
    
        sNome = ComboCampos.Text
    
        If ComboCampos.ItemData(ComboCampos.ListIndex) > -1 Then
            iPos = InStr(ComboCampos.Text, "(")
            If iPos > 0 Then sNome = left(ComboCampos.Text, iPos - 1)
        End If
        
        For Each Controle In gobjTelaAtiva.Controls
            If Controle.Name = sNome Then
                If ComboCampos.ItemData(ComboCampos.ListIndex) > -1 Then
                    If Controle.Index = ComboCampos.ItemData(ComboCampos.ListIndex) Then Exit For
                Else
                    Exit For
                End If
            End If
        Next
        
        Set objControle1 = Controle
        
        If (TypeName(Controle) = "Label") Or (TypeName(Controle) = "CommandButton") Or (TypeName(Controle) = "OptionButton") Or (TypeName(Controle) = "Frame") Or (TypeName(Controle) = "CheckBox") Then
            Titulo.Enabled = True
            Titulo.Text = Controle.Caption
            LabelTitulo.Enabled = True
        Else
            Titulo.Enabled = False
            Titulo.Text = ""
            LabelTitulo.Enabled = False
        End If
    
        If TypeName(Controle) = "ComboBox" Then
            Altura.Enabled = False
            Altura.Text = ""
            LabelAltura.Enabled = False
        Else
            Altura.Enabled = True
            Altura.Text = Controle.Height
            LabelAltura.Enabled = True
        End If
    
        If Controle.Visible = True Then
            ComboVisivel.ListIndex = 0
        Else
            ComboVisivel.ListIndex = 1
        End If
        
'        If (TypeName(Controle)= "PictureBox") Or (TypeName(Controle)= "Frame") Or (TypeName(Controle)= "SSFrame") Or (TypeName(Controle)= "SSPanel") Or (TypeName(Controle)= "Label") Then
        If (TypeName(Controle) = "PictureBox") Or (TypeName(Controle) = "Frame") Or (TypeName(Controle) = "Label") Then
            TabIndex.Enabled = False
            Label6.Enabled = False
        Else
            TabIndex.Enabled = True
            Label6.Enabled = True
        End If
        Largura.Text = Controle.Width
        Esquerda.Text = Controle.left
        Topo.Text = Controle.top
                
        TabIndex.Text = Controle.TabIndex
        
'        If Not (TypeName(Controle)= "Frame") And Not (TypeName(Controle)= "Label") And Not (TypeName(Controle)= "SSFrame") Then
        If Not (TypeName(Controle) = "Frame") And Not (TypeName(Controle) = "Label") Then
            ComboTabStop.Enabled = True
            LabelTabStop.Enabled = True
            If Controle.TabStop = True Then
                ComboTabStop.ListIndex = 0
            Else
                ComboTabStop.ListIndex = 1
            End If
        Else
            ComboTabStop.ListIndex = -1
            ComboTabStop.Enabled = False
            LabelTabStop.Enabled = False
        End If
        
        sNome = Controle.Container.Name
        If Not (Controle.Container Is gobjTelaAtiva) Then sNome = sNome & "(" & Controle.Container.Index & ")"
        Container.Text = sNome
        
        If Controle.Enabled Then
            Habilitado.Value = vbChecked
        Else
            Habilitado.Value = vbUnchecked
        End If
    
        Call Adiciona_Controle
    
    End If
    
    Exit Sub

Erro_ComboCampos_Click:

    Select Case Err
    
        Case 343
            Resume Next
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 165864)
            
    End Select
    
    Exit Sub

End Sub

Private Sub ComboTabStop_Validate(Cancel As Boolean)

    If Not (objControle1 Is Nothing) Then
        Call Adiciona_Controle
    End If

End Sub

Private Sub Form_Load()

Dim Formato As RECT

On Error GoTo Erro_Form_Load

    Set gobjIncluido = New ClassEdicaoTela_Tela 'Inserido por Wagner
    
    Call GetWindowRect(Me.hWnd, Formato)
    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, Formato.left, Formato.top, Formato.right - Formato.left, Formato.bottom - Formato.top, SWP_SHOWWINDOW)
    Set gobjPropriedades = Me
    
    Exit Sub
    
Erro_Form_Load:

    Select Case Err
            
        Case 129290, 129291
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 165865)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload
    
    If ComboCampos.ListIndex <> -1 Then
        Adiciona_Controle
        Remove_Controle
    End If
    
    gobjmenuEdicao.Checked = False
    Unload gobjCamposInvisiveis
    Set gobjPropriedades = Nothing
       
    Exit Sub
    
Erro_Form_Unload:
    
    Select Case gErr
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165866)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoSalvar_Click()

Dim lErro As Long
Dim objGrupoUsu As New ClassGrupoUsuarios
Dim colGrupoUsu As New Collection
Dim objUsu As New ClassUsuarios
Dim bFlag As Boolean

On Error GoTo Erro_BotaoSalvar_Click
       
    If gobjIncluido Is Nothing Then Exit Sub
    
    bFlag = False
    
    'Inserido por Wagner
    'Chama Tela Para Capturar os grupos a serem alterados
    '####################
    Load GrupoUsuarios

    lErro = GrupoUsuarios.Trata_Parametros(colGrupoUsu)
    If lErro <> SUCESSO Then gError 129290

    GrupoUsuarios.Show vbModal
        
    objUsu.sCodUsuario = gsUsuario
    lErro = CF("Usuarios_Le", objUsu)
    If lErro <> SUCESSO Then gError 129298

    'Inserido por Wagner
    For Each objGrupoUsu In colGrupoUsu
    
        If objGrupoUsu.sCodGrupo = objUsu.sCodGrupo Then bFlag = True
    
        gobjIncluido.sGrupoUsuarios = objGrupoUsu.sCodGrupo

        lErro = CF("EdicaoTela_Grava1", gobjIncluido)
        If lErro <> SUCESSO Then gError 64039
            
    Next
    
    If bFlag Then
        gobjIncluido.sGrupoUsuarios = objUsu.sCodGrupo
    Else
        gcolEdicaoTela.Remove gobjIncluido.sNomeTela
    End If
    '##########################
        
    Exit Sub
    
Erro_BotaoSalvar_Click:
    
    Select Case gErr
            
        Case 64039, 129290, 129298
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165867)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Titulo_Change()
    
    If Not (objControle1 Is Nothing) Then
        
'????
        If (TypeName(objControle1) = "Label") Or (TypeName(objControle1) = "CommandButton") Or (TypeName(objControle1) = "OptionButton") Or (TypeName(objControle1) = "Frame") Or (TypeName(objControle1) = "CheckBox") Then
            objControle1.Caption = Titulo.Text
        End If
    
    End If

End Sub

Private Sub ComboVisivel_Click()

Dim vbMsg As VbMsgBoxResult

    If Not (objControle1 Is Nothing) And ComboVisivel.ListIndex <> -1 Then
        
        If ComboVisivel.ItemData(ComboVisivel.ListIndex) = 1 Then
            objControle1.Visible = True
            If objControle1.Visible = False Then
                ComboVisivel.ListIndex = 1
                vbMsg = Rotina_Aviso(vbOKOnly, "AVISO_NAO_TORNOU_VISIVEL")
            End If
        ElseIf ComboVisivel.ItemData(ComboVisivel.ListIndex) = 0 Then
            objControle1.Visible = False
        End If
                    
        Call CamposInvisiveis.Carrega_Campos_Invisiveis
            
    End If

End Sub

Private Sub Largura_Validate(Cancel As Boolean)
    
Dim lErro As Long

On Error GoTo Erro_Largura_Validate
    
    lErro = Inteiro_Critica3(Largura.Text)
    If lErro <> SUCESSO Then Error 64040
    
    If Not (objControle1 Is Nothing) Then
        objControle1.Width = CInt(Largura.Text)
        Call Adiciona_Controle
    End If
        
    Exit Sub
    
Erro_Largura_Validate:
    
    Cancel = True
    
    Select Case Err
        
        Case 64040
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 165868)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Altura_Validate(Cancel As Boolean)
    
Dim lErro As Long

On Error GoTo Erro_Altura_Validate
    
    lErro = Inteiro_Critica3(Altura.Text)
    If lErro <> SUCESSO Then Error 64041
    
    If Not (objControle1 Is Nothing) Then
        objControle1.Height = CInt(Altura.Text)
        Call Adiciona_Controle
    End If
    
    Exit Sub
    
Erro_Altura_Validate:

    Cancel = True

    Select Case Err
        
        Case 64041
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 165869)

    End Select
    
    Exit Sub
        
End Sub

Private Sub Esquerda_Validate(Cancel As Boolean)
    
Dim lErro As Long

On Error GoTo Erro_Esquerda_Validate

    lErro = Inteiro_Critica2(Esquerda.Text)
    If lErro <> SUCESSO Then Error 64042
    
    If Not (objControle1 Is Nothing) Then
        objControle1.left = CInt(Esquerda.Text)
        Call Adiciona_Controle
    End If
    
    Exit Sub
    
Erro_Esquerda_Validate:

    Cancel = True

    Select Case Err
        
        Case 64042
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 165870)

    End Select
    
    Exit Sub
        
End Sub

Private Sub Titulo_Validate(Cancel As Boolean)

    If Not (objControle1 Is Nothing) Then
        Call Adiciona_Controle
    End If

End Sub

Private Sub Topo_Validate(Cancel As Boolean)
    
Dim lErro As Long

On Error GoTo Erro_Topo_Validate

    lErro = Inteiro_Critica2(Topo.Text)
    If lErro <> SUCESSO Then Error 64043
    
    If Not (objControle1 Is Nothing) Then
        objControle1.top = CInt(Topo.Text)
        Call Adiciona_Controle
    End If

    Exit Sub
    
Erro_Topo_Validate:
    
    Cancel = True

    Select Case Err
        
        Case 64043
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 165871)

    End Select
    
    Exit Sub

End Sub

Private Sub TabIndex_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TabIndex_Validate

'    lErro = Inteiro_Critica3(Topo.Text)
'    If lErro <> SUCESSO Then Error 64044

    lErro = Preenche_TabIndex_Colecao()
    If lErro <> SUCESSO Then Error 64031
    
    Exit Sub
    
Erro_TabIndex_Validate:

    Cancel = True

    Select Case Err
    
        Case 64031, 64044
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 165872)
    
    End Select
    
    Exit Sub

End Sub

Private Sub ComboTabStop_Click()

    If Not (objControle1 Is Nothing) And ComboTabStop.ListIndex <> -1 And ComboTabStop.Enabled = True Then
        
        If ComboTabStop.ItemData(ComboTabStop.ListIndex) = 1 Then
            objControle1.TabStop = True
        ElseIf ComboTabStop.ItemData(ComboTabStop.ListIndex) = 0 Then
            objControle1.TabStop = False
        End If
            
    End If

End Sub

Private Sub Container_Validate(Cancel As Boolean)
    
Dim Controle As Object
Dim iAchou As Integer
Dim sNome As String

On Error GoTo Erro_Container_Validate
    
    If Not (objControle1 Is Nothing) And ComboCampos.ListIndex <> -1 Then
        
        iAchou = 0

        For Each Controle In gobjTelaAtiva.Controls
            sNome = Controle.Name
            If Controle.Index > -1 Then sNome = sNome & "(" & Controle.Index & ")"
            If sNome = Container.Text Then
                Set objControle1.Container = Controle
                iAchou = 1
                Exit For
            End If
        Next
    
        If iAchou <> 1 Then
    
            If Container.Text = gobjTelaAtiva.Name Then
                Set objControle1.Container = gobjTelaAtiva
                iAchou = 1
            End If
            
        End If
        
        If iAchou <> 1 Then
        
            If Container.Text = "" Then
                sNome = objControle1.Container.Name
                If Not (objControle1.Container Is gobjTelaAtiva) Then sNome = sNome & "(" & objControle1.Container.Index & ")"
                Container.Text = sNome
            Else
                Error 60855
            End If
        End If
        
        Call Adiciona_Controle
    
    End If
    
    Exit Sub

Erro_Container_Validate:

    Select Case Err
    
        Case 343
            Resume Next
        
        Case 60855
            Cancel = True
            Call MsgBox(ERRO_CONTAINER_INVALIDO, vbOKOnly, Err)
            'Call Rotina_Erro(vbOKOnly, "ERRO_CONTAINER_INVALIDO", Err)
        
        Case Else
            Cancel = True
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 165873)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoParaFrente_Click()
    
    If Not (objControle1 Is Nothing) Then
        objControle1.ZOrder 0
        Call Adiciona_Controle(0)
    End If

End Sub


Private Sub BotaoParaTras_Click()

    If Not (objControle1 Is Nothing) Then
        objControle1.ZOrder 1
        Call Adiciona_Controle(1)
    End If

End Sub


Public Sub Limpar()

    If ComboCampos.ListIndex <> -1 Then
        Adiciona_Controle
        Remove_Controle
    End If
    
    Set gobjTelaAtiva = Nothing
    ComboCampos.Clear
    Titulo.Text = ""
    Largura.Text = ""
    Altura.Text = ""
    Topo.Text = ""
    Esquerda.Text = ""
    ComboTabStop.ListIndex = -1
    ComboVisivel.ListIndex = -1
    Container.Text = ""
    TabIndex.Text = ""
    Habilitado.Value = vbChecked
    CamposInvisiveis.ListaCamposInvisiveis.Clear

End Sub

Sub Adiciona_Controle(Optional izOrder As Integer = -1)

Dim objEdicaoTela_Tela As ClassEdicaoTela_Tela
Dim objEdicaoTela_Controle As ClassEdicaoTela_Controle
Dim bEncontrouControle As Boolean
Dim bEncontrouTela As Boolean
Dim iPosicaoParentese As Integer
Dim lErro As Long
Dim sControle As String
Dim iIndice As Integer

On Error GoTo Erro_Adiciona_Controle

    bEncontrouControle = False
    bEncontrouTela = False

    iIndice = -1
    iIndice = objControle1.Index

    If iIndice > -1 Then
        sControle = objControle1.Name & "(" & objControle1.Index & ")"
    Else
        sControle = objControle1.Name
    End If

    'Procura na colecao
    For Each objEdicaoTela_Tela In gcolEdicaoTela
        If objEdicaoTela_Tela.sNomeTela = gobjTelaAtiva.Name Then
            bEncontrouTela = True
            For Each objEdicaoTela_Controle In objEdicaoTela_Tela.colEdicaoTela_Controle
                
                If objEdicaoTela_Controle.sNomeControle = sControle Then
                    bEncontrouControle = True
             
                    'Inserido por Wagner
                    If izOrder <> -1 Then objEdicaoTela_Controle.izOrder = izOrder
                                       
                    Exit For
                End If
            Next
            Exit For
        End If
    Next

    'Se não encontrou adiciona a Tela(Controle E Tela Nova)
    If bEncontrouTela = False Then
        Set objEdicaoTela_Controle = New ClassEdicaoTela_Controle
        
        'Alterado por Wagner
        objEdicaoTela_Controle.izOrder = izOrder
        
        Set objEdicaoTela_Tela = New ClassEdicaoTela_Tela
        Set objEdicaoTela_Tela.colEdicaoTela_Controle = New Collection
        objEdicaoTela_Tela.sNomeTela = gobjTelaAtiva.Name
        
'????
        If (TypeName(objControle1) = "Label") Or (TypeName(objControle1) = "Button") Or (TypeName(objControle1) = "OptionButton") Or (TypeName(objControle1) = "Frame") Or (TypeName(objControle1) = "CheckBox") Then
            objEdicaoTela_Controle.sTituloPadrao = objControle1.Caption
            objEdicaoTela_Controle.sTitulo = Titulo.Text
        End If
        
        If Altura.Enabled = False Then
            objEdicaoTela_Controle.iAlturaPadrao = -1
            objEdicaoTela_Controle.iAltura = -1
        Else
            objEdicaoTela_Controle.iAlturaPadrao = objControle1.Height
            objEdicaoTela_Controle.iAltura = CInt(Altura.Text)
        End If
        
'        If (Not (TypeName(objControle1)= "PictureBox")) And (Not (TypeName(objControle1)= "Frame")) And (Not (TypeName(objControle1)= "SSFrame")) And (Not (TypeName(objControle1)= "SSPanel")) And (Not (TypeName(objControle1)= "Label")) Then
        If (Not (TypeName(objControle1) = "PictureBox")) And (Not (TypeName(objControle1) = "Frame")) And (Not (TypeName(objControle1) = "Label")) Then
            objEdicaoTela_Controle.iTabIndexPadrao = objControle1.TabIndex
            objEdicaoTela_Controle.iTabIndex = CInt(TabIndex.Text)
        End If
        
        objEdicaoTela_Controle.iEsquerdaPadrao = objControle1.left
        objEdicaoTela_Controle.iEsquerda = CInt(Esquerda.Text)
        objEdicaoTela_Controle.iLarguraPadrao = objControle1.Width
        objEdicaoTela_Controle.iLargura = CInt(Largura.Text)
        
        If ComboTabStop.Enabled = False Then
            objEdicaoTela_Controle.iTabStop = -1
            objEdicaoTela_Controle.iTabStopPadrao = -1
        Else
            objEdicaoTela_Controle.iTabStop = ComboTabStop.ItemData(ComboTabStop.ListIndex)
            If objControle1.TabStop = True Then
                objEdicaoTela_Controle.iTabStopPadrao = 1
            Else
                objEdicaoTela_Controle.iTabStopPadrao = 0
            End If
        End If
        
        objEdicaoTela_Controle.iTopoPadrao = objControle1.top
        objEdicaoTela_Controle.iTopo = CInt(Topo.Text)
        
        If objControle1.Enabled Then
            objEdicaoTela_Controle.iHabilitadoPadrao = MARCADO
        Else
            objEdicaoTela_Controle.iHabilitadoPadrao = DESMARCADO
        End If
        If Habilitado.Value = vbChecked Then
            objEdicaoTela_Controle.iHabilitado = MARCADO
        Else
            objEdicaoTela_Controle.iHabilitado = DESMARCADO
        End If
        
        objEdicaoTela_Controle.iVisivelPadrao = objControle1.Visible
        objEdicaoTela_Controle.iVisivel = ComboVisivel.ItemData(ComboVisivel.ListIndex)
        If gobjTelaAtiva.Name = objControle1.Container.Name Then
            objEdicaoTela_Controle.iIndiceContainerPadrao = -1
            objEdicaoTela_Controle.sContainerPadrao = objControle1.Container.Name
        Else
            objEdicaoTela_Controle.iIndiceContainerPadrao = -1
            objEdicaoTela_Controle.iIndiceContainerPadrao = objControle1.Container.Index
            objEdicaoTela_Controle.sContainerPadrao = objControle1.Container.Name
        End If
        
        iPosicaoParentese = InStr(Container.Text, "(")
        If iPosicaoParentese > 1 Then
             objEdicaoTela_Controle.sContainer = Mid(Container.Text, 1, iPosicaoParentese - 1)
             objEdicaoTela_Controle.iIndiceContainer = CInt(Mid(Container.Text, iPosicaoParentese + 1, Len(Trim(Container.Text)) - (iPosicaoParentese + 1)))
        Else
             objEdicaoTela_Controle.sContainer = Container.Text
             objEdicaoTela_Controle.iIndiceContainer = -1
        End If
        
        objEdicaoTela_Controle.sNomeControle = sControle
        objEdicaoTela_Controle.sNomeTela = gobjTelaAtiva.Name

        objEdicaoTela_Tela.colEdicaoTela_Controle.Add objEdicaoTela_Controle, sControle
        Set gobjIncluido = objEdicaoTela_Tela 'Inserido por Wagner
        gcolEdicaoTela.Add objEdicaoTela_Tela, objEdicaoTela_Tela.sNomeTela
    Else
    'Se encontrou a Tela e não encontrou o Controle
        If bEncontrouControle = False Then
            If (objEdicaoTela_Tela.colEdicaoTela_Controle Is Nothing) Then
                Set objEdicaoTela_Tela.colEdicaoTela_Controle = New Collection
            End If
            
            Set objEdicaoTela_Tela = gcolEdicaoTela.Item(gobjTelaAtiva.Name)
            Set gobjIncluido = objEdicaoTela_Tela
            Set objEdicaoTela_Controle = New ClassEdicaoTela_Controle
            objEdicaoTela_Tela.sNomeTela = gobjTelaAtiva.Name
    
            If (TypeName(objControle1) = "Label") Or (TypeName(objControle1) = "CommandButton") Or (TypeName(objControle1) = "OptionButton") Or (TypeName(objControle1) = "Frame") Or (TypeName(objControle1) = "CheckBox") Then
                objEdicaoTela_Controle.sTituloPadrao = objControle1.Caption
                objEdicaoTela_Controle.sTitulo = Titulo.Text
            End If
        
            'Alterado por Wagner
            objEdicaoTela_Controle.izOrder = izOrder
          
            If Altura.Enabled = False Then
                objEdicaoTela_Controle.iAlturaPadrao = -1
                objEdicaoTela_Controle.iAltura = -1
            Else
                objEdicaoTela_Controle.iAlturaPadrao = objControle1.Height
                If Len(Trim(Altura.Text)) > 0 Then objEdicaoTela_Controle.iAltura = CInt(Altura.Text)
            End If
            
'            If (Not (TypeName(objControle1)= "PictureBox")) And (Not (TypeName(objControle1)= "Frame")) And (Not (TypeName(objControle1)=  "SSFrame")) And (Not (TypeName(objControle1)= "SSPanel")) And (Not (TypeName(objControle1)= "Label")) Then
            If (Not (TypeName(objControle1) = "PictureBox")) And (Not (TypeName(objControle1) = "Frame")) And (Not (TypeName(objControle1) = "Label")) Then
                objEdicaoTela_Controle.iTabIndexPadrao = objControle1.TabIndex
                objEdicaoTela_Controle.iTabIndex = CInt(TabIndex.Text)
            End If
            
            objEdicaoTela_Controle.iEsquerdaPadrao = objControle1.left
            If Len(Trim(Esquerda.Text)) > 0 Then objEdicaoTela_Controle.iEsquerda = CInt(Esquerda.Text)
            objEdicaoTela_Controle.iLarguraPadrao = objControle1.Width
            If Len(Trim(Largura.Text)) > 0 Then objEdicaoTela_Controle.iLargura = CInt(Largura.Text)
            
            If ComboTabStop.Enabled = False Then
                objEdicaoTela_Controle.iTabStop = -1
                objEdicaoTela_Controle.iTabStopPadrao = -1
            Else
                objEdicaoTela_Controle.iTabStop = ComboTabStop.ItemData(ComboTabStop.ListIndex)
                If objControle1.TabStop = True Then
                    objEdicaoTela_Controle.iTabStopPadrao = 1
                Else
                    objEdicaoTela_Controle.iTabStopPadrao = 0
                End If
            End If
            
            objEdicaoTela_Controle.iTopoPadrao = objControle1.top
            If Len(Trim(Topo.Text)) > 0 Then objEdicaoTela_Controle.iTopo = CInt(Topo.Text)
            
            If objControle1.Enabled Then
                objEdicaoTela_Controle.iHabilitadoPadrao = MARCADO
            Else
                objEdicaoTela_Controle.iHabilitadoPadrao = DESMARCADO
            End If
            If Habilitado.Value = vbChecked Then
                objEdicaoTela_Controle.iHabilitado = MARCADO
            Else
                objEdicaoTela_Controle.iHabilitado = DESMARCADO
            End If
            
            objEdicaoTela_Controle.iVisivelPadrao = objControle1.Visible
            If Len(Trim(ComboVisivel.Text)) > 0 Then objEdicaoTela_Controle.iVisivel = ComboVisivel.ItemData(ComboVisivel.ListIndex)
              
            If Trim(gobjTelaAtiva.Name) = Trim(objControle1.Container.Name) Then
                objEdicaoTela_Controle.iIndiceContainerPadrao = -1
                objEdicaoTela_Controle.sContainerPadrao = objControle1.Container.Name
            Else
                objEdicaoTela_Controle.iIndiceContainerPadrao = -1
                objEdicaoTela_Controle.iIndiceContainerPadrao = objControle1.Container.Index
                objEdicaoTela_Controle.sContainerPadrao = objControle1.Container.Name
            End If
            
            If Len(Trim(Container.Text)) > 0 Then
                iPosicaoParentese = InStr(Container.Text, "(")
                If iPosicaoParentese > 1 Then
                     objEdicaoTela_Controle.sContainer = Mid(Container.Text, 1, iPosicaoParentese - 1)
                     objEdicaoTela_Controle.iIndiceContainer = CInt(Mid(Container.Text, iPosicaoParentese + 1, Len(Trim(Container.Text)) - (iPosicaoParentese + 1)))
                Else
                     objEdicaoTela_Controle.sContainer = Container.Text
                     objEdicaoTela_Controle.iIndiceContainer = -1
                End If
            End If
            
            objEdicaoTela_Controle.sNomeControle = sControle
            objEdicaoTela_Controle.sNomeTela = gobjTelaAtiva.Name
            objEdicaoTela_Tela.colEdicaoTela_Controle.Add objEdicaoTela_Controle, sControle
        
        ElseIf bEncontrouControle = True Then
            
            Set objEdicaoTela_Tela = gcolEdicaoTela.Item(gobjTelaAtiva.Name)
            Set gobjIncluido = objEdicaoTela_Tela
            Set objEdicaoTela_Controle = objEdicaoTela_Tela.colEdicaoTela_Controle(sControle)
            
            If Altura.Enabled = True Then
                If Len(Trim(Altura.Text)) > 0 Then objEdicaoTela_Controle.iAltura = CInt(Altura.Text)
            Else
                objEdicaoTela_Controle.iAltura = -1
            End If
            
'????
            If (TypeName(objControle1) = "Label") Or (TypeName(objControle1) = "CommandButton") Or (TypeName(objControle1) = "OptionButton") Or (TypeName(objControle1) = "Frame") Or (TypeName(objControle1) = "CheckBox") Then
                objEdicaoTela_Controle.sTitulo = Titulo.Text
            End If
            
'            If (Not (TypeName(objControle1)= "PictureBox")) And (Not (TypeName(objControle1)= "Frame")) And (Not (TypeName(objControle1)= "SSFrame")) And (Not (TypeName(objControle1)= "SSPanel")) And (Not (TypeName(objControle1)= "Label")) Then
            If (Not (TypeName(objControle1) = "PictureBox")) And (Not (TypeName(objControle1) = "Frame")) And (Not (TypeName(objControle1) = "Label")) Then
                If Len(Trim(TabIndex.Text)) > 0 Then objEdicaoTela_Controle.iTabIndex = CInt(TabIndex.Text)
            End If

            If Len(Trim(Esquerda.Text)) > 0 Then objEdicaoTela_Controle.iEsquerda = CInt(Esquerda.Text)
            If Len(Trim(Largura.Text)) > 0 Then objEdicaoTela_Controle.iLargura = CInt(Largura.Text)
            
            If ComboTabStop.Enabled = False Then
                objEdicaoTela_Controle.iTabStop = -1
            Else
                objEdicaoTela_Controle.iTabStop = ComboTabStop.ItemData(ComboTabStop.ListIndex)
            End If
            
            If Len(Trim(Topo.Text)) > 0 Then objEdicaoTela_Controle.iTopo = CInt(Topo.Text)
            
            If Habilitado.Value = vbChecked Then
                objEdicaoTela_Controle.iHabilitado = MARCADO
            Else
                objEdicaoTela_Controle.iHabilitado = DESMARCADO
            End If
            
            If Len(Trim(ComboVisivel)) > 0 Then objEdicaoTela_Controle.iVisivel = ComboVisivel.ItemData(ComboVisivel.ListIndex)
                    
            If Len(Trim(Container.Text)) > 0 Then
                iPosicaoParentese = InStr(Container.Text, "(")
                If iPosicaoParentese > 1 Then
                     objEdicaoTela_Controle.sContainer = Mid(Container.Text, 1, iPosicaoParentese - 1)
                     objEdicaoTela_Controle.iIndiceContainer = CInt(Mid(Container.Text, iPosicaoParentese + 1, Len(Trim(Container.Text)) - (iPosicaoParentese + 1)))
                Else
                     objEdicaoTela_Controle.sContainer = Container.Text
                     objEdicaoTela_Controle.iIndiceContainer = -1
                End If
            End If
            
            objEdicaoTela_Controle.sNomeControle = sControle
            objEdicaoTela_Controle.sNomeTela = gobjTelaAtiva.Name
            
        End If

    End If
    
    Exit Sub
    
Erro_Adiciona_Controle:
    
    Select Case Err
    
        Case 343
            Resume Next
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165874)
    
    End Select
    
    Exit Sub
        
End Sub

Sub Remove_Controle()

Dim objEdicaoTela_Tela As ClassEdicaoTela_Tela
Dim objEdicaoTela_Controle As ClassEdicaoTela_Controle
Dim iIndice As Integer
Dim bEncontrou As Boolean
Dim iIndice2 As Integer
Dim sControle As String
Dim iIndice3 As Integer
Dim lErro As Long

'Inserido por Wagner
Dim colGrupoUsu As New Collection
Dim objUsu As New ClassUsuarios
Dim objGrupoUsu As ClassGrupoUsuarios
Dim bFlag As Boolean

On Error GoTo Erro_Remove_Controle

    bEncontrou = False
    
    'Percorre a colecao de Telas
    For iIndice = gcolEdicaoTela.Count To 1 Step -1
        
        Set objEdicaoTela_Tela = gcolEdicaoTela.Item(iIndice)
        
        If objEdicaoTela_Tela.sNomeTela = gobjTelaAtiva.Name Then
                
            'Percorre a colecao de Controles
            For iIndice2 = objEdicaoTela_Tela.colEdicaoTela_Controle.Count To 1 Step -1
                
                Set objEdicaoTela_Controle = objEdicaoTela_Tela.colEdicaoTela_Controle.Item(iIndice2)
                
                iIndice3 = -1
                iIndice3 = objControle1.Index

                If iIndice3 > -1 Then
                    sControle = objControle1.Name & "(" & objControle1.Index & ")"
                Else
                    sControle = objControle1.Name
                End If

'                If objEdicaoTela_Controle.sNomeControle = sControle Then
                    
                    'Se tiver algunm controle que tenha suas propriedades igual ao padrao
                    If objEdicaoTela_Controle.iAlturaPadrao = objEdicaoTela_Controle.iAltura And objEdicaoTela_Controle.iEsquerdaPadrao = objEdicaoTela_Controle.iEsquerda And objEdicaoTela_Controle.iLarguraPadrao = objEdicaoTela_Controle.iLargura And objEdicaoTela_Controle.iTabStopPadrao = objEdicaoTela_Controle.iTabStop And objEdicaoTela_Controle.iTopoPadrao = objEdicaoTela_Controle.iTopo And objEdicaoTela_Controle.iIndiceContainerPadrao = objEdicaoTela_Controle.iIndiceContainer And objEdicaoTela_Controle.sContainer = objEdicaoTela_Controle.sContainerPadrao And objEdicaoTela_Controle.iTabIndex = objEdicaoTela_Controle.iTabIndexPadrao And objEdicaoTela_Controle.sTitulo = objEdicaoTela_Controle.sTituloPadrao And objEdicaoTela_Controle.izOrder = -1 And objEdicaoTela_Controle.iHabilitado = objEdicaoTela_Controle.iHabilitadoPadrao Then
                        
                        'Remove da Colecao
                        objEdicaoTela_Tela.colEdicaoTela_Controle.Remove (iIndice2)
                          
                    End If
'                End If
            Next
        
            'Se a Tela está sem controel e não houve alteracao no Tab Index
            If objEdicaoTela_Tela.colEdicaoTela_Controle.Count = 0 And objEdicaoTela_Tela.sTabIndex = objEdicaoTela_Tela.sTabIndexPadrao Then
                'Remove da coleção
                
                
                Load GrupoUsuarios
            
                lErro = GrupoUsuarios.Trata_Parametros(colGrupoUsu)
                If lErro <> SUCESSO Then gError 129290
            
                GrupoUsuarios.Show vbModal
        
                lErro = CF("EdicaoTela_Exclui", gcolEdicaoTela.Item(iIndice), colGrupoUsu) 'Inserido por Wagner
                
                objUsu.sCodUsuario = gsUsuario
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO Then gError 129310
                
                bFlag = False
                
                For Each objGrupoUsu In colGrupoUsu
                
                    If objGrupoUsu.sCodGrupo = objUsu.sCodGrupo Then bFlag = True
                
                Next
                
                If bFlag Then Call gcolEdicaoTela.Remove(iIndice)
            
            End If
            
            Exit For
        
        End If
    Next
    
    Exit Sub
    
Erro_Remove_Controle:
    
    Select Case Err
    
        Case 343
            Resume Next
            
        Case 129310
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165875)
    
    End Select
    
    Exit Sub
   
End Sub

Function Preenche_TabIndex_Colecao() As Long

Dim objEdicaoTela_Tela As ClassEdicaoTela_Tela
Dim objEdicaoTela_Controle As ClassEdicaoTela_Controle
Dim objControle As Object
Dim sTabIndexPadrao As String
Dim sTabIndexAtual As String
Dim bEncontrou As Boolean

On Error GoTo Erro_Preenche_TabIndex_Colecao

    sTabIndexPadrao = " "
    sTabIndexAtual = " "
    bEncontrou = False

    For Each objEdicaoTela_Tela In gcolEdicaoTela
    
        If objEdicaoTela_Tela.sNomeTela = gobjTelaAtiva.Name Then
            
            bEncontrou = True
            
            'Se está Inserindo o Tab Index então
            If Len(Trim(objEdicaoTela_Tela.sTabIndexPadrao)) = 0 Then
                'Percorre Todos os controles para guardar os TabIndex Padrões
                For Each objControle In gobjTelaAtiva.Controls
'                    If (Not (TypeName(objControle)= "Label")) And (Not (TypeName(objControle)= "PictureBox")) And (Not (TypeName(objControle)= "Frame")) And (Not (TypeName(objControle)= "SSFrame")) And (Not (TypeName(objControle)= "SSPanel")) And (Not (TypeName(objControle)= "Line")) Then
                    If (Not (TypeName(objControle) = "Label")) And (Not (TypeName(objControle) = "PictureBox")) And (Not (TypeName(objControle) = "Frame")) And (Not (TypeName(objControle) = "Line")) And Not (TypeName(objControle) = "CommonDialog") And Not (TypeName(objControle) = "Image") Then
                        objEdicaoTela_Tela.sTabIndexPadrao = objEdicaoTela_Tela.sTabIndexPadrao & " " & objControle.TabIndex
                    End If
                Next
                objEdicaoTela_Tela.sTabIndexPadrao = objEdicaoTela_Tela.sTabIndexPadrao & " "
            End If
            
            objControle1.TabIndex = CInt(TabIndex.Text)

            objEdicaoTela_Tela.sTabIndex = ""
            'Percorre Todos os controles para guardar os TabIndex Atuais
            For Each objControle In gobjTelaAtiva.Controls
'                If (Not (TypeName(objControle)= "Label")) And (Not (TypeName(objControle)= "PictureBox")) And (Not (TypeName(objControle)= "Frame")) And (Not (TypeName(objControle)= "SSFrame")) And (Not (TypeName(objControle)= "SSPanel")) And (Not (TypeName(objControle)= "Line")) Then
                If (Not (TypeName(objControle) = "Label")) And (Not (TypeName(objControle) = "PictureBox")) And (Not (TypeName(objControle) = "Frame")) And (Not (TypeName(objControle) = "Line")) And Not (TypeName(objControle) = "CommonDialog") And Not (TypeName(objControle) = "Image") Then
                    objEdicaoTela_Tela.sTabIndex = objEdicaoTela_Tela.sTabIndex & " " & objControle.TabIndex
                End If
            Next
            objEdicaoTela_Tela.sTabIndex = objEdicaoTela_Tela.sTabIndex & " "
            
        End If
    Next
 
    If bEncontrou = False Then
        
        Set objEdicaoTela_Tela = New ClassEdicaoTela_Tela
        Set objEdicaoTela_Tela.colEdicaoTela_Controle = New Collection
        
        objEdicaoTela_Tela.sNomeTela = gobjTelaAtiva.Name
        
        objEdicaoTela_Tela.sTabIndexPadrao = ""
        'Percorre Todos os controles para guardar os TabIndex Padrões
        For Each objControle In gobjTelaAtiva.Controls
'            If (Not (TypeName(objControle)= "Label")) And (Not (TypeName(objControle)= "PictureBox")) And (Not (TypeName(objControle)= "Frame")) And (Not (TypeName(objControle)= "SSFrame")) And (Not (TypeName(objControle)= "SSPanel")) And (Not (TypeName(objControle)= "Line")) Then
             If (Not (TypeName(objControle) = "Label")) And (Not (TypeName(objControle) = "PictureBox")) And (Not (TypeName(objControle) = "Frame")) And (Not (TypeName(objControle) = "Line")) And Not (TypeName(objControle) = "CommonDialog") And Not (TypeName(objControle) = "Image") Then
                objEdicaoTela_Tela.sTabIndexPadrao = objEdicaoTela_Tela.sTabIndexPadrao & " " & objControle.TabIndex
            End If
        Next
        objEdicaoTela_Tela.sTabIndexPadrao = objEdicaoTela_Tela.sTabIndexPadrao & " "
        
        If Not (objControle1 Is Nothing) Then
            objControle1.TabIndex = CInt(TabIndex.Text)
        End If

        objEdicaoTela_Tela.sTabIndex = ""
        'Percorre Todos os controles para guardar os TabIndex Atuais
        For Each objControle In gobjTelaAtiva.Controls
'            If (Not (TypeName(objControle)= "Label")) And (Not (TypeName(objControle)= "PictureBox")) And (Not (TypeName(objControle)= "Frame")) And (Not (TypeName(objControle)= "SSFrame")) And (Not (TypeName(objControle)= "SSPanel")) And (Not (TypeName(objControle)= "Line")) Then
            If (Not (TypeName(objControle) = "Label")) And (Not (TypeName(objControle) = "PictureBox")) And (Not (TypeName(objControle) = "Frame")) And (Not (TypeName(objControle) = "Line")) And Not (TypeName(objControle) = "CommonDialog") And Not (TypeName(objControle) = "Image") Then
                objEdicaoTela_Tela.sTabIndex = objEdicaoTela_Tela.sTabIndex & " " & objControle.TabIndex
            End If
        Next
        objEdicaoTela_Tela.sTabIndex = objEdicaoTela_Tela.sTabIndex & " "
        Set gobjIncluido = objEdicaoTela_Tela 'Inserido por Wagner
        gcolEdicaoTela.Add objEdicaoTela_Tela, objEdicaoTela_Tela.sNomeTela
    
    End If
    
    Preenche_TabIndex_Colecao = SUCESSO
    
    Exit Function
    
Erro_Preenche_TabIndex_Colecao:
    
    Preenche_TabIndex_Colecao = Err
    
    Select Case Err
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165876)
    
    End Select
    
    Exit Function
    
End Function

Function Inteiro_Critica2(sNumero As String) As Long
'Critica se é Tipo inteiro (entre -32000 e 32000)

Dim lErro As Long
Dim dNumero As Double

On Error GoTo Error_Inteiro_Critica2

    If IsNumeric(sNumero) Then
        dNumero = 0#
        dNumero = CDbl(sNumero)
    Else
        Error 64036
    End If

    If dNumero > 32000 Then Error 64037
    If dNumero < -32000 Then Error 64050
    
    Inteiro_Critica2 = SUCESSO

    Exit Function

Error_Inteiro_Critica2:

    Inteiro_Critica2 = Err

    Select Case Err

        Case 64036
            Call MsgBox("O valor " & sNumero & " tem que ser numérico.", vbOKOnly, Err)

        Case 64037, 64050
            Call MsgBox("O número " & sNumero & " não está entre -32000 e 32000.", vbOKOnly, Err)
            
        Case Else
            Call MsgBox(ERRO_FORNECIDO_PELO_VB, vbOKOnly, Err)

    End Select

    Exit Function

End Function

Function Inteiro_Critica3(sNumero As String) As Long
'Critica se é Tipo inteiro (entre 0 e 32000)

Dim lErro As Long
Dim dNumero As Double

On Error GoTo Error_Inteiro_Critica3

    If IsNumeric(sNumero) Then
        dNumero = 0#
        dNumero = CDbl(sNumero)
    Else
        Error 64052
    End If

    If dNumero > 32000 Then Error 64053
    If dNumero < 0 Then Error 64054
    
    Inteiro_Critica3 = SUCESSO

    Exit Function

Error_Inteiro_Critica3:

    Inteiro_Critica3 = Err

    Select Case Err

        Case 64052
            Call MsgBox("O valor " & sNumero & " tem que ser numérico.", vbOKOnly, Err)

        Case 64053, 64054
            Call MsgBox("O número " & sNumero & " não está entre 0 e 32000.", vbOKOnly, Err)
            
        Case Else
            Call MsgBox(ERRO_FORNECIDO_PELO_VB, vbOKOnly, Err)

    End Select

    Exit Function

End Function

Private Sub Habilitado_Click()
    
Dim lErro As Long

On Error GoTo Erro_Habilitado_Click
    
    If Not (objControle1 Is Nothing) Then
        If Habilitado.Value = vbChecked Then
            objControle1.Enabled = True
        Else
            objControle1.Enabled = False
        End If
        Call Adiciona_Controle
    End If

    Exit Sub
    
Erro_Habilitado_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165871)

    End Select
    
    Exit Sub

End Sub

