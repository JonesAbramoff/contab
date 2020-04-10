VERSION 5.00
Begin VB.Form Imprimir 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Leg 
      Appearance      =   0  'Flat
      Columns         =   5
      Height          =   900
      IntegralHeight  =   0   'False
      ItemData        =   "Imprimir.frx":0000
      Left            =   165
      List            =   "Imprimir.frx":0002
      TabIndex        =   5
      Top             =   5805
      Width           =   7665
   End
   Begin VB.Frame Figura 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5085
      Left            =   150
      TabIndex        =   0
      Top             =   1545
      Width           =   7665
      Begin VB.PictureBox Linha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   30
         Index           =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   165
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.PictureBox Seta 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   135
         Index           =   0
         Left            =   2220
         Picture         =   "Imprimir.frx":0004
         ScaleHeight     =   67.5
         ScaleMode       =   0  'User
         ScaleWidth      =   28.846
         TabIndex        =   7
         Top             =   975
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   660
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   180
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin VB.Label LinhaColuna 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7050
      TabIndex        =   6
      Top             =   60
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Texto2 
      BackColor       =   &H80000005&
      Caption         =   "bbbbbbbbbbbb"
      Height          =   675
      Left            =   5790
      TabIndex        =   4
      Top             =   795
      Width           =   2055
   End
   Begin VB.Label Texto 
      BackColor       =   &H80000005&
      Caption         =   "bbbbbbbbbbbb"
      Height          =   675
      Left            =   165
      TabIndex        =   2
      Top             =   795
      Width           =   5565
   End
   Begin VB.Label Nome 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "aaaaaaaaaa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   210
      TabIndex        =   1
      Top             =   105
      Width           =   7665
   End
End
Attribute VB_Name = "Imprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAX_WIDTH_TELA = 15000
Const MAX_HEIGHT_TELA = 9000

Public Function Imprimir_Layout(ByVal objTelaGraficoImp As ClassTelaGraficoImpressao) As Long

Dim lErro As Long
Dim objTelaGraficoImpItens As ClassTelaGraficoImpItens
Dim iIndice As Integer
Dim lMaisBaixo As Long
Dim lMaisDireita As Long
Dim lMaisAlto As Long
Dim lMaisEsquerda As Long
Dim lAumentoAltura As Long
Dim lAumentoLargura As Long
Dim dFatorDeImpressao As Double
Dim dFatorDeImpressaoAlt As Double
Dim lNumLeg As Long
Dim iNumCol As Integer
Dim objControle As Object
Dim bForaAlt As Boolean
Dim bForaComp As Boolean
Dim iLinha As Integer
Dim iColuna As Integer
Dim iNumFiguras As Integer
Dim bLinhas As Boolean
Dim lLarguraFixa As Long
Dim lAlturaFixa As Long
Dim bFora As Boolean

On Error GoTo Erro_Imprimir_Layout

    lMaisEsquerda = 900000
    lMaisBaixo = 900000
    
    Texto.Caption = objTelaGraficoImp.sTexto
    Texto2.Caption = objTelaGraficoImp.sTexto2
    Nome.Caption = objTelaGraficoImp.sNome

    For Each objTelaGraficoImpItens In objTelaGraficoImp.colItens
    
        If objTelaGraficoImpItens.lFontSize = 0 Then
            objTelaGraficoImpItens.lFontSize = objTelaGraficoImp.lFontSize
        End If
        
        If objTelaGraficoImpItens.lBackColor = 0 Then
            objTelaGraficoImpItens.lBackColor = objTelaGraficoImp.lBackColor
        End If
    
        If objTelaGraficoImpItens.lForeColor = 0 Then
            objTelaGraficoImpItens.lForeColor = objTelaGraficoImp.lForeColor
        End If
    
        If Len(Trim(objTelaGraficoImpItens.sFontName)) = 0 Then
            objTelaGraficoImpItens.sFontName = objTelaGraficoImp.sFontName
        End If
        
        If objTelaGraficoImpItens.lLeft < lMaisEsquerda Then
            lMaisEsquerda = objTelaGraficoImpItens.lLeft
        End If
        
        If objTelaGraficoImpItens.lTop < lMaisBaixo Then
            lMaisBaixo = objTelaGraficoImpItens.lTop
        End If
        
        If objTelaGraficoImpItens.lLeft + objTelaGraficoImpItens.lWidth > lMaisDireita Then
            lMaisDireita = objTelaGraficoImpItens.lLeft + objTelaGraficoImpItens.lWidth
        End If
        
        If objTelaGraficoImpItens.lTop + objTelaGraficoImpItens.lHeight > lMaisAlto Then
            lMaisAlto = objTelaGraficoImpItens.lTop + objTelaGraficoImpItens.lHeight
        End If
        
        lAumentoAltura = lMaisAlto - lMaisBaixo - Figura.Height
        If lAumentoAltura < 0 Then lAumentoAltura = 0
        
        lAumentoLargura = lMaisDireita - lMaisEsquerda - Figura.Width
        If lAumentoLargura < 0 Then lAumentoLargura = 0
        
    Next
    
    If Me.Width + lAumentoLargura < MAX_WIDTH_TELA Then
        dFatorDeImpressao = 1
        Me.Width = Me.Width + lAumentoLargura
        Figura.Width = Figura.Width + lAumentoLargura
    Else
        dFatorDeImpressao = (Me.Width + lAumentoLargura) / MAX_WIDTH_TELA
        Figura.Width = -(Me.Width - Figura.Width) + MAX_WIDTH_TELA
        Me.Width = MAX_WIDTH_TELA
    End If
    
    If Me.Height + lAumentoAltura < MAX_HEIGHT_TELA Then
        dFatorDeImpressaoAlt = 1
        Figura.Height = Figura.Height + lAumentoAltura
        Me.Height = Me.Height + lAumentoAltura
    Else
        dFatorDeImpressaoAlt = (Me.Height + lAumentoAltura) / MAX_HEIGHT_TELA
        Figura.Height = -(Me.Height - Figura.Height) + MAX_HEIGHT_TELA + 500
        Me.Height = MAX_HEIGHT_TELA + 500
    End If
    
    'TESTE
    dFatorDeImpressao = 1
    dFatorDeImpressaoAlt = 1
        
    Texto.Width = Figura.Width
    Nome.Width = Figura.Width
          
    For Each objTelaGraficoImpItens In objTelaGraficoImp.colItens
    
        If objTelaGraficoImpItens.iLegenda = MARCADO Then
            lNumLeg = lNumLeg + 1
        End If
     
        If objTelaGraficoImpItens.iTipo = TELAGRAFICOIMPITENS_TIPO_TEXT Or objTelaGraficoImpItens.iTipo = TELAGRAFICOIMPITENS_TIPO_TEXT_FIXO_LINHA Or objTelaGraficoImpItens.iTipo = TELAGRAFICOIMPITENS_TIPO_TEXT_FIXO_COLUNA Then
        
            iIndice = Text.UBound + 1
            'Inclui na tela um novo Controle para essa coluna
            Load Text(iIndice)
            'Traz o controle recem desenhado para a frente
            Text(iIndice).ZOrder
            'Torna o controle visível
            Text(iIndice).Visible = True
            
            Text(iIndice).Text = objTelaGraficoImpItens.sText
            Text(iIndice).FontName = objTelaGraficoImpItens.sFontName
            
            Text(iIndice).BackColor = objTelaGraficoImpItens.lBackColor
            Text(iIndice).ForeColor = objTelaGraficoImpItens.lForeColor
            Text(iIndice).Width = objTelaGraficoImpItens.lWidth / dFatorDeImpressao
            Text(iIndice).Top = (objTelaGraficoImpItens.lTop - lMaisBaixo) / dFatorDeImpressaoAlt
            Text(iIndice).Left = (objTelaGraficoImpItens.lLeft - lMaisEsquerda) / dFatorDeImpressao
            Text(iIndice).Height = objTelaGraficoImpItens.lHeight / dFatorDeImpressaoAlt
            Text(iIndice).FontSize = objTelaGraficoImpItens.lFontSize
            Text(iIndice).BorderStyle = objTelaGraficoImpItens.iBorderStyle
            
            Select Case objTelaGraficoImpItens.iTipo
                Case TELAGRAFICOIMPITENS_TIPO_TEXT_FIXO_COLUNA
                    Text(iIndice).Tag = "COLUNA"
                    lLarguraFixa = Text(iIndice).Width
                Case TELAGRAFICOIMPITENS_TIPO_TEXT_FIXO_LINHA
                    Text(iIndice).Tag = "LINHA"
                    lAlturaFixa = Text(iIndice).Height
                Case Else
                    Text(iIndice).Tag = ""
            End Select
            
        ElseIf objTelaGraficoImpItens.iTipo = TELAGRAFICOIMPITENS_TIPO_LINE Then
        
            iIndice = Linha.UBound + 1
            'Inclui na tela um novo Controle para essa coluna
            Load Linha(iIndice)
            'Traz o controle recem desenhado para a frente
            Linha(iIndice).ZOrder
            'Torna o controle visível
            Linha(iIndice).Visible = True
'            Linha(iIndice).BorderWidth = 2
            'Linha(iIndice).BorderColor = objTelaGraficoImpItens.lBackColor
            Linha(iIndice).Width = objTelaGraficoImpItens.lWidth / dFatorDeImpressao
            Linha(iIndice).Top = (objTelaGraficoImpItens.lTop - lMaisBaixo) / dFatorDeImpressaoAlt
            Linha(iIndice).Left = (objTelaGraficoImpItens.lLeft - lMaisEsquerda) / dFatorDeImpressao
            Linha(iIndice).Height = objTelaGraficoImpItens.lHeight / dFatorDeImpressaoAlt
'            Linha(iIndice).X1 = (objTelaGraficoImpItens.lLeft - lMaisEsquerda) / dFatorDeImpressao
'            Linha(iIndice).Y1 = (objTelaGraficoImpItens.lTop - lMaisBaixo) / dFatorDeImpressaoAlt
'            Linha(iIndice).X2 = Linha(iIndice).X1 + objTelaGraficoImpItens.lWidth / dFatorDeImpressao
'            Linha(iIndice).Y2 = Linha(iIndice).Y1 + (objTelaGraficoImpItens.lHeight / dFatorDeImpressaoAlt)
            Linha(iIndice).BorderStyle = objTelaGraficoImpItens.iBorderStyle
        
        ElseIf objTelaGraficoImpItens.iTipo = TELAGRAFICOIMPITENS_TIPO_SETA Then
        
            iIndice = Linha.UBound + 1
            'Inclui na tela um novo Controle para essa coluna
            Load Seta(iIndice)
            'Traz o controle recem desenhado para a frente
            Seta(iIndice).ZOrder
            'Torna o controle visível
            Seta(iIndice).Visible = True
            
            Seta(iIndice).Width = objTelaGraficoImpItens.lWidth
            Seta(iIndice).Top = (objTelaGraficoImpItens.lTop - lMaisBaixo)
            Seta(iIndice).Left = (objTelaGraficoImpItens.lLeft - lMaisEsquerda)
            Seta(iIndice).Height = objTelaGraficoImpItens.lHeight
        
        End If
        
    Next
    
    If lNumLeg <> 0 Then
    
        Me.Height = Me.Height + Leg.Height + 50
        Leg.Visible = True
        Leg.Width = Figura.Width - 200
        Leg.Left = Leg.Left + 100
        Leg.Top = Figura.Top + Figura.Height + 50
                        
        iNumCol = Round(lNumLeg / 4)
        
        If iNumCol * 4 >= lNumLeg Then
            Leg.Columns = iNumCol
        Else
            Leg.Columns = iNumCol + 1
        End If
          
        For Each objTelaGraficoImpItens In objTelaGraficoImp.colItens
            If objTelaGraficoImpItens.iLegenda = MARCADO Then
                Leg.AddItem objTelaGraficoImpItens.sText & SEPARADOR & objTelaGraficoImpItens.sDescricao
            End If
        Next
    
    Else
        Leg.Visible = False
    End If
    
    Me.Refresh
              
    If Len(Trim(objTelaGraficoImp.sNomeArqFigura)) = 0 Then
        Figura.Width = -(Me.Width - Figura.Width) + MAX_WIDTH_TELA
        Me.Width = MAX_WIDTH_TELA
        Leg.Width = Figura.Width - 200
        Texto.Width = Figura.Width
        Nome.Width = Figura.Width
    Else
        Me.Show
        Me.Top = 0
        Me.Left = 0
    End If
    
    iLinha = 1
    iColuna = 1
    iNumFiguras = 0
    bLinhas = False
    bFora = False
    bForaAlt = True
    Do While bForaAlt Or bForaComp
    
        If Printer.Orientation <> 2 Then
            Printer.Orientation = 2
        End If
    
        iNumFiguras = iNumFiguras + 1
        bForaAlt = False
        bForaComp = False
        For Each objControle In Controls
            
            If objControle.Container.Name = Figura.Name Then
        
                'Se deixaria de imprimir alguma coisa porque ficou abaixo do Frame
                If objControle.Height + objControle.Top > Figura.Height + Figura.Top Then
                    bForaAlt = True
                End If
                'Se deixaria de imprimir alguma coisa porque ficou a direita do Frame
                If objControle.Width + objControle.Left > Figura.Width + Figura.Left Then
                    bForaComp = True
                End If
            
            End If
            
        Next
        
        If bForaAlt Or bForaComp Then bFora = True
        If bForaAlt Then bLinhas = True
        
        If bFora Then
            LinhaColuna.Left = Me.Width - LinhaColuna.Width - 250
            LinhaColuna.Visible = True
            LinhaColuna.Caption = CStr(iLinha) & SEPARADOR & CStr(iColuna)
        End If
        
'        If bLinhas Then
'            If iLinha = 1 Then
'                Nome.Visible = True
'                Texto.Visible = True
'                Texto2.Visible = True
'                Leg.Visible = False
'            Else
'                Nome.Visible = False
'                Texto.Visible = False
'                Texto2.Visible = False
'                Leg.Visible = True
'            End If
'        End If
            
        Me.AutoRedraw = True
        Me.Refresh
        DoEvents
        If Len(Trim(objTelaGraficoImp.sNomeArqFigura)) = 0 Then
            Me.PrintForm
            Printer.EndDoc
        Else
            If iNumFiguras = 1 Then
                SavePicture CaptureClient(Me), objTelaGraficoImp.sNomeArqFigura
            Else
                SavePicture CaptureClient(Me), Left(objTelaGraficoImp.sNomeArqFigura, Len(objTelaGraficoImp.sNomeArqFigura) - 4) & SEPARADOR & CStr(iNumFiguras) & ".bmp"
            End If
        End If
        
        If bForaComp Then
            iColuna = iColuna + 1
            For Each objControle In Controls
                If objControle.Container.Name = "Figura" Then
                    If objControle.Tag <> "COLUNA" Then
                        objControle.Left = objControle.Left - Figura.Width + lLarguraFixa + (Figura.Width Mod lLarguraFixa)
                    End If
                End If
            Next
        ElseIf bForaAlt Then
            For Each objControle In Controls
                If objControle.Container.Name = "Figura" Then
                    If objControle.Tag <> "COLUNA" Then
                        objControle.Left = objControle.Left + ((Figura.Width - lLarguraFixa - (Figura.Width Mod lLarguraFixa)) * (iColuna - 1))
                    End If
                End If
            Next
            iColuna = 1
            iLinha = iLinha + 1
            For Each objControle In Controls
                If objControle.Container.Name = "Figura" Then
                    If objControle.Tag <> "LINHA" Then
                        objControle.Top = objControle.Top - Figura.Height + lAlturaFixa + (Figura.Height Mod lAlturaFixa)
                    End If
                End If
            Next
        End If
        
    Loop
    
    objTelaGraficoImp.iNumFiguras = iNumFiguras

    If Len(Trim(objTelaGraficoImp.sNomeArqFigura)) <> 0 Then
        Me.Visible = False
        Unload Me
    End If

    Imprimir_Layout = SUCESSO

    Exit Function

Erro_Imprimir_Layout:

    Printer.EndDoc
    Me.Visible = False
    Unload Me

    Imprimir_Layout = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185741)

    End Select

    Exit Function

End Function
