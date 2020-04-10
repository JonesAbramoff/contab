VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TRPLiberaOcrAux 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   3210
      Picture         =   "TRPLiberaOcrAux.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5355
      Width           =   1005
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   4980
      Picture         =   "TRPLiberaOcrAux.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5355
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de Inativações de Vouchers Pagos com Cartão de Crédito"
      Height          =   5190
      Left            =   75
      TabIndex        =   11
      Top             =   75
      Width           =   9405
      Begin VB.TextBox ISS 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   3660
         TabIndex        =   50
         Text            =   "ISS"
         Top             =   1380
         Width           =   615
      End
      Begin VB.TextBox COFINS 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   6210
         TabIndex        =   49
         Text            =   "COFINS"
         Top             =   1095
         Width           =   585
      End
      Begin VB.TextBox PIS 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   7185
         TabIndex        =   48
         Text            =   "PIS"
         Top             =   1125
         Width           =   615
      End
      Begin VB.TextBox TAR 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5055
         TabIndex        =   47
         Text            =   "TAR"
         Top             =   1320
         Width           =   615
      End
      Begin VB.Frame Detalhe 
         Caption         =   "Detalhe"
         Enabled         =   0   'False
         Height          =   2880
         Left            =   75
         TabIndex        =   20
         Top             =   2205
         Width           =   9240
         Begin VB.CheckBox optImpostos 
            Caption         =   "Abater Impostos"
            Enabled         =   0   'False
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
            Left            =   1155
            TabIndex        =   40
            Top             =   1620
            Width           =   1875
         End
         Begin VB.CheckBox optTarifa 
            Caption         =   "Abater Tarifa"
            Enabled         =   0   'False
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
            Left            =   1155
            TabIndex        =   37
            Top             =   1155
            Width           =   1545
         End
         Begin MSMask.MaskEdBox PercTarifa 
            Height          =   330
            Left            =   5235
            TabIndex        =   1
            Top             =   1110
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercPis 
            Height          =   330
            Left            =   5235
            TabIndex        =   3
            Top             =   1545
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercCofins 
            Height          =   330
            Left            =   5235
            TabIndex        =   5
            Top             =   1980
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercIss 
            Height          =   330
            Left            =   5235
            TabIndex        =   7
            Top             =   2415
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorTarifa 
            Height          =   345
            Left            =   7725
            TabIndex        =   2
            Top             =   1110
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorPis 
            Height          =   345
            Left            =   7725
            TabIndex        =   4
            Top             =   1545
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorCofins 
            Height          =   345
            Left            =   7725
            TabIndex        =   6
            Top             =   1980
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorISS 
            Height          =   345
            Left            =   7725
            TabIndex        =   8
            Top             =   2415
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Caption         =   "Valor ISS:"
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
            Index           =   15
            Left            =   6765
            TabIndex        =   46
            Top             =   2490
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "% ISS:"
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
            Index           =   14
            Left            =   4575
            TabIndex        =   45
            Top             =   2490
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Valor Cofins:"
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
            Index           =   13
            Left            =   6540
            TabIndex        =   44
            Top             =   2040
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "% Cofins:"
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
            Index           =   12
            Left            =   4350
            TabIndex        =   43
            Top             =   2085
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "Valor PIS:"
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
            Index           =   11
            Left            =   6765
            TabIndex        =   42
            Top             =   1620
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "% PIS:"
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
            Index           =   10
            Left            =   4545
            TabIndex        =   41
            Top             =   1620
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Valor Tarifa:"
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
            Index           =   9
            Left            =   6570
            TabIndex        =   39
            Top             =   1155
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "% Tarifa:"
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
            Index           =   8
            Left            =   4365
            TabIndex        =   38
            Top             =   1155
            Width           =   780
         End
         Begin VB.Label ValorOcrNovo 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1140
            TabIndex        =   36
            Top             =   2415
            Width           =   1665
         End
         Begin VB.Label Label1 
            Caption         =   "Valor Novo:"
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
            Index           =   7
            Left            =   60
            TabIndex        =   35
            Top             =   2475
            Width           =   1020
         End
         Begin VB.Label Label1 
            Caption         =   "Núm. Ocr:"
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
            Index           =   6
            Left            =   195
            TabIndex        =   34
            Top             =   285
            Width           =   885
         End
         Begin VB.Label NumOcr 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1155
            TabIndex        =   33
            Top             =   225
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Vou:"
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
            Index           =   0
            Left            =   2385
            TabIndex        =   32
            Top             =   285
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Série Vou:"
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
            Index           =   2
            Left            =   4275
            TabIndex        =   31
            Top             =   285
            Width           =   930
         End
         Begin VB.Label Label1 
            Caption         =   "Núm. Vou:"
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
            Index           =   3
            Left            =   6720
            TabIndex        =   30
            Top             =   285
            Width           =   930
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente:"
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
            Index           =   1
            Left            =   405
            TabIndex        =   29
            Top             =   705
            Width           =   630
         End
         Begin VB.Label Label1 
            Caption         =   "Data de Emissão:"
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
            Index           =   5
            Left            =   6150
            TabIndex        =   28
            Top             =   735
            Width           =   1545
         End
         Begin VB.Label TipoVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   3255
            TabIndex        =   27
            Top             =   225
            Width           =   795
         End
         Begin VB.Label SerieVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5235
            TabIndex        =   26
            Top             =   225
            Width           =   795
         End
         Begin VB.Label NumVou 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   7725
            TabIndex        =   25
            Top             =   225
            Width           =   1275
         End
         Begin VB.Label ClienteOcr 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1155
            TabIndex        =   24
            Top             =   660
            Width           =   4875
         End
         Begin VB.Label DataEmissaoOcr 
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   7725
            TabIndex        =   23
            Top             =   660
            Width           =   1290
         End
         Begin VB.Label ValorOcrAtual 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1155
            TabIndex        =   22
            Top             =   1980
            Width           =   1665
         End
         Begin VB.Label Label1 
            Caption         =   "Valor Atual:"
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
            Index           =   4
            Left            =   75
            TabIndex        =   21
            Top             =   2040
            Width           =   1020
         End
      End
      Begin VB.TextBox ValorNovo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   7860
         TabIndex        =   19
         Text            =   "Valor"
         Top             =   615
         Width           =   945
      End
      Begin VB.TextBox Valor 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5415
         TabIndex        =   18
         Text            =   "Valor"
         Top             =   600
         Width           =   1005
      End
      Begin VB.TextBox Tipo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4755
         TabIndex        =   17
         Text            =   "Tipo"
         Top             =   600
         Width           =   510
      End
      Begin VB.TextBox Ocr 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Text            =   "Ocr"
         Top             =   600
         Width           =   780
      End
      Begin VB.TextBox Serie 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4170
         TabIndex        =   15
         Text            =   "Serie"
         Top             =   585
         Width           =   450
      End
      Begin VB.TextBox Vou 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   3195
         TabIndex        =   14
         Text            =   "Vou"
         Top             =   600
         Width           =   765
      End
      Begin VB.CheckBox Tarifa 
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
         Left            =   1365
         TabIndex        =   13
         Top             =   630
         Width           =   585
      End
      Begin VB.CheckBox Impostos 
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
         Left            =   1395
         TabIndex        =   12
         Top             =   960
         Width           =   450
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   1905
         Left            =   45
         TabIndex        =   0
         Top             =   255
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   3360
         _Version        =   393216
         Rows            =   15
         Cols            =   8
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
      End
   End
End
Attribute VB_Name = "TRPLiberaOcrAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iLinhaAnt As Integer

'Grid Itens:
Dim objGridItens As AdmGrid
Dim iGrid_Impostos_Col As Integer
Dim iGrid_Tarifa_Col As Integer
Dim iGrid_Ocr_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_Serie_Col  As Integer
Dim iGrid_Vou_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_TAR_Col As Integer
Dim iGrid_PIS_Col As Integer
Dim iGrid_COFINS_Col As Integer
Dim iGrid_ISS_Col As Integer
Dim iGrid_ValorN_Col As Integer

Dim gcolOcorrencias As Collection
Dim gcolOcorrenciasTela As Collection

Private Function Traz_Ocorrencias_Tela() As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOcr As ClassTRPOcorrencias
Dim objVou As ClassTRPVouchers
Dim iFilialEmpCorporator As Integer
Dim iFilialEmpCoinfo As Integer
Dim objTitRecTRP As ClassTitulosRecTRP
Dim dPercTarifa As Double

On Error GoTo Erro_Traz_Ocorrencias_Tela

    'Limpa o GridItens
    Call Grid_Limpa(objGridItens)
    
    Set gcolOcorrenciasTela = New Collection
    
    For Each objOcr In gcolOcorrencias
    
        'Se é uma inativação
        If objOcr.iOrigem = INATIVACAO_AUTOMATICA_CODIGO Then
        
            Set objVou = New ClassTRPVouchers

            objVou.sTipVou = objOcr.sTipoDoc
            objVou.sSerie = objOcr.sSerie
            objVou.lNumVou = objOcr.lNumVou
    
            lErro = CF("TRPVouchers_Le", objVou)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192786
            
            'Se é um voucher de cartão
            If objVou.iCartao = MARCADO Then
            
                gcolOcorrenciasTela.Add objOcr
    
                iLinha = iLinha + 1
        
                'Passa para a tela os dados do Itens em questão
                GridItens.TextMatrix(iLinha, iGrid_Impostos_Col) = CStr(DESMARCADO)
                GridItens.TextMatrix(iLinha, iGrid_Tarifa_Col) = CStr(DESMARCADO)
                GridItens.TextMatrix(iLinha, iGrid_Tipo_Col) = objVou.sTipVou
                GridItens.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objOcr.dValorTotal, "STANDARD")
                GridItens.TextMatrix(iLinha, iGrid_ValorN_Col) = Format(objOcr.dValorTotal, "STANDARD")
                GridItens.TextMatrix(iLinha, iGrid_Vou_Col) = CStr(objVou.lNumVou)
                GridItens.TextMatrix(iLinha, iGrid_Serie_Col) = objVou.sSerie
                GridItens.TextMatrix(iLinha, iGrid_Ocr_Col) = CStr(objOcr.lCodigo)
                
                dPercTarifa = 0
                
                'Se o voucher está associado a um título
                If objVou.iTipoDocDestino = TRP_TIPO_DOC_DESTINO_TITREC And objVou.lNumIntDocDestino <> 0 Then
                
                    Set objTitRecTRP = New ClassTitulosRecTRP
                    
                    objTitRecTRP.lNumIntDocTitRec = objVou.lNumIntDocDestino
                                        
                    lErro = CF("TitulosRecTRP_Le", objTitRecTRP)
                    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192788
                    
                    If lErro <> ERRO_LEITURA_SEM_DADOS Then
                        If objTitRecTRP.dValorBruto <> 0 Then
                            dPercTarifa = objTitRecTRP.dValorTarifa / objTitRecTRP.dValorBruto
                        End If
                    End If

                End If
                
                GridItens.TextMatrix(iLinha, iGrid_TAR_Col) = Format(Abs(objOcr.dValorTotal) * 0.05, "STANDARD")
                GridItens.TextMatrix(iLinha, iGrid_PIS_Col) = Format(Abs(objOcr.dValorTotal) * 0.0065, "STANDARD")
                GridItens.TextMatrix(iLinha, iGrid_COFINS_Col) = Format(Abs(objOcr.dValorTotal) * 0.03, "STANDARD")
                
            End If
            
        End If
        
    Next
    
    objGridItens.iLinhasExistentes = iLinha
    
    Call Grid_Refresh_Checkbox(objGridItens)
            
    Traz_Ocorrencias_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Ocorrencias_Tela:

    Traz_Ocorrencias_Tela = gErr
    
    Select Case gErr
    
        Case 192786 To 192788

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192789)

    End Select

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridItens = Nothing
    Set gcolOcorrencias = Nothing
    Set gcolOcorrenciasTela = Nothing
    
End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer
Dim colColecoes As New Collection

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If
    
    colColecoes.Add gcolOcorrencias
    colColecoes.Add gcolOcorrenciasTela
    
    Call Ordenacao_ClickGrid(objGridItens, , colColecoes)

End Sub

Private Sub GridItens_GotFocus()
    Call Grid_Recebe_Foco(objGridItens)
End Sub

Private Sub GridItens_EnterCell()
    Call Grid_Entrada_Celula(objGridItens, iAlterado)
End Sub

Private Sub GridItens_LeaveCell()
    Call Saida_Celula(objGridItens)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridItens)
End Sub

Private Sub GridItens_RowColChange()
    Call Grid_RowColChange(objGridItens)
    
    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
        Detalhe.Enabled = True
    Else
        Detalhe.Enabled = False
    End If
    
    Call Recolhe_Dados(iLinhaAnt)
    Call Mostra_Dados(GridItens.Row)
    
    iLinhaAnt = GridItens.Row

End Sub

Private Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Public Sub Form_Load()
    
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load
    
    giRetornoTela = vbCancel
    
    Set objGridItens = New AdmGrid
    
    'Executa a Inicialização do grid Itens
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError 192790
        
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 192790

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192791)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional colOcorrencias As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    Set gcolOcorrencias = colOcorrencias
    
    lErro = Traz_Ocorrencias_Tela
    If lErro <> SUCESSO Then gError 192792
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 192792

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192793)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    Call Ordenacao_Limpa(objGridItens)

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceélula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridItens.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                    
            End Select
            
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 192794

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 192794
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192795)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Itens
    
    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Tarifa")
    objGridInt.colColuna.Add ("Imp.")
    objGridInt.colColuna.Add ("Núm.Ocr")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Série")
    objGridInt.colColuna.Add ("Núm.Vou")
    objGridInt.colColuna.Add ("Valor Atual")
    objGridInt.colColuna.Add ("Tar.")
    objGridInt.colColuna.Add ("Pis")
    objGridInt.colColuna.Add ("Cofins")
    objGridInt.colColuna.Add ("Iss")
    objGridInt.colColuna.Add ("Novo Valor")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Tarifa.Name)
    objGridInt.colCampo.Add (Impostos.Name)
    objGridInt.colCampo.Add (Ocr.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (Serie.Name)
    objGridInt.colCampo.Add (Vou.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (TAR.Name)
    objGridInt.colCampo.Add (PIS.Name)
    objGridInt.colCampo.Add (COFINS.Name)
    objGridInt.colCampo.Add (ISS.Name)
    objGridInt.colCampo.Add (ValorNovo.Name)

    iGrid_Tarifa_Col = 1
    iGrid_Impostos_Col = 2
    iGrid_Ocr_Col = 3
    iGrid_Tipo_Col = 4
    iGrid_Serie_Col = 5
    iGrid_Vou_Col = 6
    iGrid_Valor_Col = 7
    iGrid_TAR_Col = 8
    iGrid_PIS_Col = 9
    iGrid_COFINS_Col = 10
    iGrid_ISS_Col = 11
    iGrid_ValorN_Col = 12
    
    objGridInt.objGrid = GridItens

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não permite incluir novas linhas
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Passa os itens do Grid para a colecao
    lErro = Move_Tela_Memoria(gcolOcorrencias)
    If lErro <> SUCESSO Then gError 192796
  
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 192796

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192797)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal colOcorrencias As Collection) As Long
'Altera as informações da Ocorrência

Dim lErro As Long
Dim iLinha As Integer
Dim objOcorrencia As ClassTRPOcorrencias
Dim objOcorrenciaAux As ClassTRPOcorrencias
Dim objOcorrenciaDet As ClassTRPOcorrenciaDet
Dim colOcorrenciasAux As New Collection
Dim bAchou As Boolean
Dim dValor As Double

On Error GoTo Erro_Move_Tela_Memoria

    For iLinha = 1 To objGridItens.iLinhasExistentes
    
        Set objOcorrencia = gcolOcorrenciasTela.Item(iLinha)
        Set objOcorrenciaAux = New ClassTRPOcorrencias
        
        objOcorrenciaAux.lCodigo = objOcorrencia.lCodigo
        
        'Lê o a Ocorrência
        lErro = CF("TRPOcorrencias_Le", objOcorrenciaAux)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192798
    
        'TARIFA
        If StrParaInt(GridItens.TextMatrix(iLinha, iGrid_Tarifa_Col)) = MARCADO Then
        
            bAchou = False
            For Each objOcorrenciaDet In objOcorrenciaAux.colDetalhamento
                If objOcorrenciaDet.iTipo = INATIVACAO_AUTOMATICA_TIPO_TAR_CODIGO Then
                    bAchou = True
                    Exit For
                End If
            Next
        
            If Not bAchou Then
            
                Set objOcorrenciaDet = New ClassTRPOcorrenciaDet
                
                objOcorrenciaDet.iSeq = objOcorrenciaAux.colDetalhamento.Count + 1
                objOcorrenciaDet.iTipo = INATIVACAO_AUTOMATICA_TIPO_TAR_CODIGO
                
                objOcorrenciaAux.colDetalhamento.Add objOcorrenciaDet
                
            End If
            
            objOcorrenciaDet.lNumIntDocOCR = objOcorrencia.lNumIntDoc
            objOcorrenciaDet.dValor = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_TAR_Col))
                    
        End If
        
        'IMPOSTOS
        If StrParaInt(GridItens.TextMatrix(iLinha, iGrid_Impostos_Col)) = MARCADO Then
        
            'ISS
            bAchou = False
            For Each objOcorrenciaDet In objOcorrenciaAux.colDetalhamento
                If objOcorrenciaDet.iTipo = INATIVACAO_AUTOMATICA_TIPO_ISS_CODIGO Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If Not bAchou Then
        
                Set objOcorrenciaDet = New ClassTRPOcorrenciaDet
                
                objOcorrenciaDet.iSeq = objOcorrenciaAux.colDetalhamento.Count + 1
                objOcorrenciaDet.iTipo = INATIVACAO_AUTOMATICA_TIPO_ISS_CODIGO
                
                objOcorrenciaAux.colDetalhamento.Add objOcorrenciaDet
                
            End If
                
            objOcorrenciaDet.lNumIntDocOCR = objOcorrencia.lNumIntDoc
            objOcorrenciaDet.dValor = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ISS_Col))
            
            'PIS
            bAchou = False
            For Each objOcorrenciaDet In objOcorrenciaAux.colDetalhamento
                If objOcorrenciaDet.iTipo = INATIVACAO_AUTOMATICA_TIPO_PIS_CODIGO Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If Not bAchou Then
            
                Set objOcorrenciaDet = New ClassTRPOcorrenciaDet
                
                objOcorrenciaDet.iSeq = objOcorrenciaAux.colDetalhamento.Count + 1
                objOcorrenciaDet.iTipo = INATIVACAO_AUTOMATICA_TIPO_PIS_CODIGO
                
                objOcorrenciaAux.colDetalhamento.Add objOcorrenciaDet
                
            End If
            
            objOcorrenciaDet.lNumIntDocOCR = objOcorrencia.lNumIntDoc
            objOcorrenciaDet.dValor = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_PIS_Col))
            
            'COFINS
            bAchou = False
            For Each objOcorrenciaDet In objOcorrenciaAux.colDetalhamento
                If objOcorrenciaDet.iTipo = INATIVACAO_AUTOMATICA_TIPO_COFINS_CODIGO Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If Not bAchou Then
            
                Set objOcorrenciaDet = New ClassTRPOcorrenciaDet
            
                objOcorrenciaDet.iSeq = objOcorrenciaAux.colDetalhamento.Count + 1
                objOcorrenciaDet.iTipo = INATIVACAO_AUTOMATICA_TIPO_COFINS_CODIGO
            
                objOcorrenciaAux.colDetalhamento.Add objOcorrenciaDet
            
            End If
            
            objOcorrenciaDet.lNumIntDocOCR = objOcorrencia.lNumIntDoc
            objOcorrenciaDet.dValor = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_COFINS_Col))
        
        End If
        
        dValor = 0
        For Each objOcorrenciaDet In objOcorrenciaAux.colDetalhamento
            dValor = dValor + objOcorrenciaDet.dValor
        Next
        
        objOcorrenciaAux.dValorTotal = dValor
        
        colOcorrenciasAux.Add objOcorrenciaAux
        
    Next
    
    For Each objOcorrenciaAux In colOcorrenciasAux
        iLinha = -1
        For Each objOcorrencia In colOcorrencias
            iLinha = iLinha + 1
            If objOcorrencia.lCodigo = objOcorrenciaAux.lCodigo Then
                Set objOcorrencia.colDetalhamento = objOcorrenciaAux.colDetalhamento
                objOcorrencia.dValorTotal = objOcorrenciaAux.dValorTotal
                Exit For
            End If
        Next
    Next
    
'    For Each objOcorrenciaAux In colOcorrenciasAux
'        colOcorrencias.Add objOcorrenciaAux
'    Next
   
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 192798
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192799)
            
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LIBERACAO_BLOQUEIO_BLOQUEIOS
    Set Form_Load_Ocx = Me
    Caption = "Liberação de ocorrências"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TRPLiberaOcrAux"
    
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

Private Sub Impostos_Click()
    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
        Detalhe.Enabled = True
    Else
        Detalhe.Enabled = False
    End If
    
    Call Recolhe_Dados(iLinhaAnt)
    Call Mostra_Dados(GridItens.Row)
    
    iLinhaAnt = GridItens.Row
End Sub

Private Sub Tarifa_Click()
    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
        Detalhe.Enabled = True
    Else
        Detalhe.Enabled = False
    End If
    
    Call Recolhe_Dados(iLinhaAnt)
    Call Mostra_Dados(GridItens.Row)
    
    iLinhaAnt = GridItens.Row
End Sub

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

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Function Mostra_Dados(ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objOrc As ClassTRPOcorrencias
Dim objCliente As New ClassCliente

On Error GoTo Erro_Mostra_Dados

    If iLinha <> 0 And Not (gcolOcorrenciasTela Is Nothing) And iLinha <= objGridItens.iLinhasExistentes Then
    
        Set objOrc = gcolOcorrenciasTela.Item(iLinha)
        
        objCliente.lCodigo = objOrc.lCliente
    
        'le o cliente
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 192800
            
        NumOcr.Caption = CStr(objOrc.lCodigo)
        NumVou.Caption = CStr(objOrc.lNumVou)
        TipoVou.Caption = objOrc.sTipoDoc
        SerieVou.Caption = objOrc.sSerie
        ValorOcrAtual.Caption = Format(objOrc.dValorTotal, "STANDARD")
        DataEmissaoOcr.Caption = Format(objOrc.dtDataEmissao, "dd/mm/yyyy")
        ClienteOcr.Caption = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
        
        If StrParaInt(GridItens.TextMatrix(iLinha, iGrid_Tarifa_Col)) = MARCADO Then
            optTarifa.Value = vbChecked
        Else
            optTarifa.Value = vbUnchecked
        End If
        
        If StrParaInt(GridItens.TextMatrix(iLinha, iGrid_Impostos_Col)) = MARCADO Then
            optImpostos.Value = vbChecked
        Else
            optImpostos.Value = vbUnchecked
        End If
    
        ValorTarifa.Text = Format(GridItens.TextMatrix(iLinha, iGrid_TAR_Col), "STANDARD")
        Call ValorTarifa_Validate(bSGECancelDummy)
        
        ValorPis.Text = Format(GridItens.TextMatrix(iLinha, iGrid_PIS_Col), "STANDARD")
        Call ValorPIS_Validate(bSGECancelDummy)
        
        ValorCofins.Text = Format(GridItens.TextMatrix(iLinha, iGrid_COFINS_Col), "STANDARD")
        Call ValorCofins_Validate(bSGECancelDummy)
        
        ValorISS.Text = Format(GridItens.TextMatrix(iLinha, iGrid_ISS_Col), "STANDARD")
        Call ValorISS_Validate(bSGECancelDummy)
               
    End If
    
    Mostra_Dados = SUCESSO
    
    Exit Function
    
Erro_Mostra_Dados:

    Mostra_Dados = gErr

    Select Case gErr
    
        Case 192800

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192801)

    End Select

    Exit Function

End Function

Private Function Recolhe_Dados(ByVal iLinha As Integer) As Long

On Error GoTo Erro_Recolhe_Dados

    If iLinha > 0 And Not (gcolOcorrenciasTela Is Nothing) And iLinha <= objGridItens.iLinhasExistentes Then
    
        GridItens.TextMatrix(iLinha, iGrid_TAR_Col) = Format(ValorTarifa.Text, "STANDARD")
        GridItens.TextMatrix(iLinha, iGrid_ISS_Col) = Format(ValorISS.Text, "STANDARD")
        GridItens.TextMatrix(iLinha, iGrid_PIS_Col) = Format(ValorPis.Text, "STANDARD")
        GridItens.TextMatrix(iLinha, iGrid_COFINS_Col) = Format(ValorCofins.Text, "STANDARD")
        GridItens.TextMatrix(iLinha, iGrid_ValorN_Col) = Format(ValorOcrNovo.Caption, "STANDARD")
        
    End If
    
    Recolhe_Dados = SUCESSO
    
    Exit Function
    
Erro_Recolhe_Dados:

    Recolhe_Dados = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192802)

    End Select

    Exit Function

End Function

Private Sub ValorTarifa_GotFocus()
    Call MaskEdBox_TrataGotFocus(ValorTarifa, iAlterado)
End Sub

Private Sub ValorTarifa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPerc As Double

On Error GoTo Erro_ValorTarifa_Validate

    'Veifica se ValorTarifa está preenchida
    If Len(Trim(ValorTarifa.Text)) <> 0 Then

        'Critica a ValorTarifa
        lErro = Valor_Double_Critica(ValorTarifa.Text)
        If lErro <> SUCESSO Then gError 192803
        
        If StrParaDbl(ValorOcrAtual.Caption) > DELTA_VALORMONETARIO Then
            dPerc = StrParaDbl(ValorTarifa.Text) / Abs(StrParaDbl(ValorOcrAtual.Caption))
        Else
            dPerc = 0
        End If
        
        PercTarifa.Text = Formata_Estoque(dPerc * 100)
        
    Else
    
        PercTarifa.Text = ""
       
    End If
    
    Call Calcula_Valor_Novo

    Exit Sub

Erro_ValorTarifa_Validate:

    Cancel = True

    Select Case gErr

        Case 192803

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192804)

    End Select

    Exit Sub

End Sub

Private Sub ValorISS_GotFocus()
    Call MaskEdBox_TrataGotFocus(ValorISS, iAlterado)
End Sub

Private Sub ValorISS_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPerc As Double

On Error GoTo Erro_ValorISS_Validate

    'Veifica se ValorISS está preenchida
    If Len(Trim(ValorISS.Text)) <> 0 Then

        'Critica a ValorISS
        lErro = Valor_Double_Critica(ValorISS.Text)
        If lErro <> SUCESSO Then gError 192805
        
        If StrParaDbl(ValorOcrAtual.Caption) > DELTA_VALORMONETARIO Then
            dPerc = StrParaDbl(ValorISS.Text) / Abs(StrParaDbl(ValorOcrAtual.Caption))
        Else
            dPerc = 0
        End If
        
        PercIss.Text = Formata_Estoque(dPerc * 100)
        
    Else
    
        PercIss.Text = ""
       
    End If
    
    Call Calcula_Valor_Novo

    Exit Sub

Erro_ValorISS_Validate:

    Cancel = True

    Select Case gErr

        Case 192805

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192806)

    End Select

    Exit Sub

End Sub

Private Sub ValorPIS_GotFocus()
    Call MaskEdBox_TrataGotFocus(ValorPis, iAlterado)
End Sub

Private Sub ValorPIS_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPerc As Double

On Error GoTo Erro_ValorPIS_Validate

    'Veifica se ValorPIS está preenchida
    If Len(Trim(ValorPis.Text)) <> 0 Then

        'Critica a ValorPIS
        lErro = Valor_Double_Critica(ValorPis.Text)
        If lErro <> SUCESSO Then gError 192807
        
        If StrParaDbl(ValorOcrAtual.Caption) > DELTA_VALORMONETARIO Then
            dPerc = StrParaDbl(ValorPis.Text) / Abs(StrParaDbl(ValorOcrAtual.Caption))
        Else
            dPerc = 0
        End If
               
        PercPis.Text = Formata_Estoque(dPerc * 100)
        
    Else
    
        PercPis.Text = ""
       
    End If
    
    Call Calcula_Valor_Novo

    Exit Sub

Erro_ValorPIS_Validate:

    Cancel = True

    Select Case gErr

        Case 192807

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192808)

    End Select

    Exit Sub

End Sub

Private Sub ValorCofins_GotFocus()
    Call MaskEdBox_TrataGotFocus(ValorCofins, iAlterado)
End Sub

Private Sub ValorCofins_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPerc As Double

On Error GoTo Erro_ValorCofins_Validate

    'Veifica se ValorCofins está preenchida
    If Len(Trim(ValorCofins.Text)) <> 0 Then

        'Critica a ValorCofins
        lErro = Valor_Double_Critica(ValorCofins.Text)
        If lErro <> SUCESSO Then gError 192809
        
        If StrParaDbl(ValorOcrAtual.Caption) > DELTA_VALORMONETARIO Then
            dPerc = StrParaDbl(ValorCofins.Text) / Abs(StrParaDbl(ValorOcrAtual.Caption))
        Else
            dPerc = 0
        End If
        PercCofins.Text = Formata_Estoque(dPerc * 100)
        
    Else
    
        PercCofins.Text = ""
       
    End If
    
    Call Calcula_Valor_Novo

    Exit Sub

Erro_ValorCofins_Validate:

    Cancel = True

    Select Case gErr

        Case 192809

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192810)

    End Select

    Exit Sub

End Sub

Private Sub PercTarifa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPerc As Double

On Error GoTo Erro_PercTarifa_Validate

    'Verifica se PercTarifa está preenchida
    If Len(Trim(PercTarifa.Text)) <> 0 Then

       'Critica a PercTarifa
       lErro = Porcentagem_Critica(PercTarifa.Text)
       If lErro <> SUCESSO Then gError 192811
       
       dPerc = StrParaDbl(PercTarifa.Text) / 100
       
       ValorTarifa.Text = Format(dPerc * Abs(StrParaDbl(ValorOcrAtual.Caption)), "STANDARD")
       
    Else
    
        ValorTarifa.Text = ""

    End If
    
    Call Calcula_Valor_Novo

    Exit Sub

Erro_PercTarifa_Validate:

    Cancel = True

    Select Case gErr

        Case 192811

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192812)

    End Select

    Exit Sub

End Sub

Private Sub PercTarifa_GotFocus()
    Call MaskEdBox_TrataGotFocus(PercTarifa, iAlterado)
End Sub

Private Sub PercPIS_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPerc As Double

On Error GoTo Erro_PercPIS_Validate

    'Verifica se PercPIS está preenchida
    If Len(Trim(PercPis.Text)) <> 0 Then

       'Critica a PercPIS
       lErro = Porcentagem_Critica(PercPis.Text)
       If lErro <> SUCESSO Then gError 192813
       
       dPerc = StrParaDbl(PercPis.Text) / 100
       
       ValorPis.Text = Format(dPerc * Abs(StrParaDbl(ValorOcrAtual.Caption)), "STANDARD")
       
    Else
    
        ValorPis.Text = ""

    End If
    
    Call Calcula_Valor_Novo

    Exit Sub

Erro_PercPIS_Validate:

    Cancel = True

    Select Case gErr

        Case 192813

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192814)

    End Select

    Exit Sub

End Sub

Private Sub PercPIS_GotFocus()
    Call MaskEdBox_TrataGotFocus(PercPis, iAlterado)
End Sub

Private Sub PercCOFINS_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPerc As Double

On Error GoTo Erro_PercCOFINS_Validate

    'Verifica se PercCOFINS está preenchida
    If Len(Trim(PercCofins.Text)) <> 0 Then

       'Critica a PercCOFINS
       lErro = Porcentagem_Critica(PercCofins.Text)
       If lErro <> SUCESSO Then gError 192815
       
       dPerc = StrParaDbl(PercCofins.Text) / 100
       
       ValorCofins.Text = Format(dPerc * Abs(StrParaDbl(ValorOcrAtual.Caption)), "STANDARD")
       
    Else
    
        ValorCofins.Text = ""

    End If
    
    Call Calcula_Valor_Novo

    Exit Sub

Erro_PercCOFINS_Validate:

    Cancel = True

    Select Case gErr

        Case 192815

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192816)

    End Select

    Exit Sub

End Sub

Private Sub PercCOFINS_GotFocus()
    Call MaskEdBox_TrataGotFocus(PercCofins, iAlterado)
End Sub

Private Sub PercISS_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPerc As Double

On Error GoTo Erro_PercISS_Validate

    'Verifica se PercISS está preenchida
    If Len(Trim(PercIss.Text)) <> 0 Then

       'Critica a PercISS
       lErro = Porcentagem_Critica(PercIss.Text)
       If lErro <> SUCESSO Then gError 192817
       
       dPerc = StrParaDbl(PercIss.Text) / 100
       
       ValorISS.Text = Format(dPerc * Abs(StrParaDbl(ValorOcrAtual.Caption)), "STANDARD")
       
    Else
    
        ValorISS.Text = ""

    End If
    
    Call Calcula_Valor_Novo

    Exit Sub

Erro_PercISS_Validate:

    Cancel = True

    Select Case gErr

        Case 192817

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192818)

    End Select

    Exit Sub

End Sub

Private Sub PercISS_GotFocus()
    Call MaskEdBox_TrataGotFocus(PercIss, iAlterado)
End Sub

Private Function Calcula_Valor_Novo() As Long

Dim dValorNovo As Double

On Error GoTo Erro_Calcula_Valor_Novo

    dValorNovo = StrParaDbl(ValorOcrAtual.Caption)
    
    If optTarifa.Value = vbChecked Then
        dValorNovo = dValorNovo + StrParaDbl(ValorTarifa.Text)
    End If
    
    If optImpostos.Value = vbChecked Then
        dValorNovo = dValorNovo + StrParaDbl(ValorPis.Text) + StrParaDbl(ValorCofins.Text) + StrParaDbl(ValorISS.Text)
    End If

    ValorOcrNovo.Caption = Format(dValorNovo, "STANDARD")
    
    Calcula_Valor_Novo = SUCESSO
    
    Exit Function
    
Erro_Calcula_Valor_Novo:

    Calcula_Valor_Novo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192819)

    End Select

    Exit Function

End Function

Private Sub BotaoCancela_Click()
    
    'Nao mexer no obj da tela
    giRetornoTela = vbCancel
    
    Unload Me
    
    Exit Sub

End Sub

Private Sub BotaoOk_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_BotaoOk_Click

    Call Recolhe_Dados(iLinhaAnt)

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 192820
    
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbOK
    
    iAlterado = 0
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoOk_Click:

    Select Case gErr

        Case 192820
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192821)

    End Select

    Exit Sub
    
End Sub
