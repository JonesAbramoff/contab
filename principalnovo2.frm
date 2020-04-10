VERSION 5.00
Begin VB.MDIForm PrincipalNovo 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Forprint - Sistema de Gestão Empresarial"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   795
   ClientWidth     =   14535
   Icon            =   "principalnovo2.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Left            =   810
      Top             =   540
   End
   Begin VB.Timer Timer1 
      Left            =   315
      Top             =   555
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
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
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   14535
      TabIndex        =   0
      Top             =   0
      Width           =   14535
      Begin VB.CommandButton BotaoAvisos 
         BackColor       =   &H00C0C0C0&
         Caption         =   " ! "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11925
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Existem 0 avisos novos"
         Top             =   30
         Width           =   405
      End
      Begin VB.CommandButton BotaoFacebook 
         Height          =   360
         Left            =   11490
         Picture         =   "principalnovo2.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Use o facebook para manter-se atualizado com o Corporator"
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton SuporteOnline 
         Height          =   315
         Left            =   9975
         Picture         =   "principalnovo2.frx":05AB
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fale com nossa equipe de suporte..."
         Top             =   75
         Width           =   1425
      End
      Begin VB.CommandButton BotaoFilial 
         Caption         =   "Filiais"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   9975
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   15
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton BotaoBrowseCria 
         Height          =   315
         Left            =   13260
         Picture         =   "principalnovo2.frx":0C3E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Chama a tela que cria Browse"
         Top             =   75
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton BotaoNumLock 
         Caption         =   "NL"
         Height          =   315
         Left            =   9495
         TabIndex        =   14
         ToolTipText     =   "NUM LOCK"
         Top             =   60
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton BotaoAnotacao 
         Height          =   315
         Left            =   4710
         Picture         =   "principalnovo2.frx":1178
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Editar Observação"
         Top             =   60
         Width           =   375
      End
      Begin VB.CommandButton BotaoConsulta 
         Height          =   360
         Left            =   907
         Picture         =   "principalnovo2.frx":16B2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Trazer para a tela o registro corrente"
         Top             =   60
         Width           =   360
      End
      Begin VB.CommandButton Ultimo 
         Height          =   360
         Left            =   1740
         Picture         =   "principalnovo2.frx":19C4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Trazer para a tela o último registro"
         Top             =   60
         Width           =   360
      End
      Begin VB.CommandButton Proximo 
         Height          =   360
         Left            =   1323
         Picture         =   "principalnovo2.frx":1CD6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Trazer para a tela o registro seguinte"
         Top             =   60
         Width           =   360
      End
      Begin VB.CommandButton Anterior 
         Height          =   360
         Left            =   491
         Picture         =   "principalnovo2.frx":1E80
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Trazer para a tela o registro anterior"
         Top             =   60
         Width           =   360
      End
      Begin VB.CommandButton Primeiro 
         Height          =   360
         Left            =   75
         Picture         =   "principalnovo2.frx":202A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Trazer para a tela o 1o registro"
         Top             =   60
         Width           =   360
      End
      Begin VB.CommandButton BotaoTestaInt 
         Height          =   315
         Left            =   12450
         Picture         =   "principalnovo2.frx":233C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "NAO USE"
         Top             =   105
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton BotaoGeraCodigoLight 
         Height          =   315
         Left            =   12870
         Picture         =   "principalnovo2.frx":2876
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gera Código Modificado Para a Versao Light"
         Top             =   90
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton BotaoData 
         Height          =   315
         Left            =   8130
         TabIndex        =   4
         Top             =   75
         Width           =   1200
      End
      Begin VB.ComboBox ComboModulo 
         Height          =   315
         ItemData        =   "principalnovo2.frx":2A00
         Left            =   5730
         List            =   "principalnovo2.frx":2A02
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   75
         Width           =   1905
      End
      Begin VB.ComboBox Indice 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   60
         Width           =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   7695
         TabIndex        =   5
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Módulo:"
         Height          =   195
         Left            =   5130
         TabIndex        =   3
         Top             =   120
         Width           =   570
      End
      Begin VB.Line LinhaDivisao 
         BorderColor     =   &H80000005&
         X1              =   -270
         X2              =   10060
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuArqSub 
         Caption         =   "&Empresa e Filial"
         Index           =   1
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "&Data"
         Index           =   2
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "Feriado&s"
         Index           =   3
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "Cotação de &Moeda"
         Index           =   4
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "Cadastro de Moedas"
         Index           =   5
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "Outras Configurações"
         Index           =   7
         Begin VB.Menu mnuArqSubOC 
            Caption         =   "Impressoras"
            Index           =   1
         End
         Begin VB.Menu mnuArqSubOC 
            Caption         =   "Backup"
            Index           =   2
         End
         Begin VB.Menu mnuArqSubOC 
            Caption         =   "Logo"
            Index           =   3
         End
         Begin VB.Menu mnuArqSubOC 
            Caption         =   "Telas c/Tamanho Variável"
            Index           =   4
         End
         Begin VB.Menu mnuArqSubOC 
            Caption         =   "Arquivamento de dados"
            Index           =   5
         End
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "Modo de Edição"
         Index           =   9
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "Administração do &Sistema"
         Index           =   11
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "EC&F"
         Index           =   13
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "&WorkFlow"
         Index           =   15
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "Programa de Email Padrão"
         Index           =   16
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuArqSub 
         Caption         =   "&Sair"
         Index           =   18
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   1
      Begin VB.Menu mnuCTBMov 
         Caption         =   "&Lançamentos em Lote"
         Index           =   1
         Tag             =   "1"
      End
      Begin VB.Menu mnuCTBMov 
         Caption         =   "L&ançamentos"
         Index           =   2
         Tag             =   "1"
      End
      Begin VB.Menu mnuCTBMov 
         Caption         =   "&Estorno de Lote Contabilizado"
         Index           =   3
         Tag             =   "1"
      End
      Begin VB.Menu mnuCTBMov 
         Caption         =   "E&storno de Documento Contábil"
         Index           =   4
         Tag             =   "1"
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   2
      Begin VB.Menu mnuTESMov 
         Caption         =   "&Saque"
         Index           =   1
      End
      Begin VB.Menu mnuTESMov 
         Caption         =   "&Depósito"
         Index           =   2
      End
      Begin VB.Menu mnuTESMov 
         Caption         =   "&Transferência"
         Index           =   3
      End
      Begin VB.Menu mnuTESMov 
         Caption         =   "&Aplicação"
         Index           =   4
      End
      Begin VB.Menu mnuTESMov 
         Caption         =   "&Resgate"
         Index           =   5
      End
      Begin VB.Menu mnuTESMov 
         Caption         =   "&Conciliação"
         Index           =   6
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   3
      Begin VB.Menu mnuCPMov 
         Caption         =   "Notas Fiscais &Fatura"
         Index           =   1
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "&Notas Fiscais"
         Index           =   2
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "Fa&turas"
         Index           =   3
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "&Confirmação de Cobrança"
         Index           =   6
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "&Borderô de Pagamento"
         Index           =   8
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "Cheques &Automáticos"
         Index           =   9
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "Cheques &Manuais"
         Index           =   10
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "Cancelar &Pagamento"
         Index           =   12
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "&Devolução / Créditos"
         Index           =   13
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "&Adiantamento à Fornecedor"
         Index           =   14
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "Baixa"
         Index           =   15
         Begin VB.Menu mnuCPMovBx 
            Caption         =   "Bai&xa Manual"
            Index           =   1
         End
         Begin VB.Menu mnuCPMovBx 
            Caption         =   "Baixa de Adiantamentos / Créditos Fornecedores"
            Index           =   2
         End
         Begin VB.Menu mnuCPMovBx 
            Caption         =   "Baixa Manual com Cheques de Terceiros"
            Index           =   3
         End
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "&Cance&lar Baixa"
         Index           =   16
         Begin VB.Menu mnuCPMovBaixa 
            Caption         =   "Manual"
            Index           =   1
         End
         Begin VB.Menu mnuCPMovBaixa 
            Caption         =   "Por Seleção"
            Index           =   2
         End
         Begin VB.Menu mnuCPMovBaixa 
            Caption         =   "Adiantamento"
            Index           =   3
         End
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "Compensação de Cheque-Pré"
         Index           =   17
      End
      Begin VB.Menu mnuCPMov 
         Caption         =   "Liberação de Pagamento"
         Index           =   18
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   4
      Begin VB.Menu mnuCRMov 
         Caption         =   "&Títulos a Receber"
         Index           =   1
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "&Cheque Pré-Datados"
         Index           =   2
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "&Borderô de Cheques Pré"
         Index           =   3
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "Bord&erô de Desconto de Cheques"
         Index           =   4
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "E&xcluir Borderô de Desconto de Cheques"
         Index           =   5
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "B&orderô de Cobrança"
         Index           =   7
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "&Instruções para Cobrança Eletrônica"
         Index           =   8
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "Cancelar Borderô de Cobrança"
         Index           =   9
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "T&ransferência Manual"
         Index           =   11
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "&Devoluções / Entradas"
         Index           =   12
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "&Adiantamento de Cliente"
         Index           =   13
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "Bai&xa Manual"
         Index           =   14
         Begin VB.Menu mnuCRMovBxM 
            Caption         =   "Por Digitação"
            Index           =   1
         End
         Begin VB.Menu mnuCRMovBxM 
            Caption         =   "Por Seleção"
            Index           =   2
         End
         Begin VB.Menu mnuCRMovBxM 
            Caption         =   "Cancelar"
            Index           =   3
         End
         Begin VB.Menu mnuCRMovBxM 
            Caption         =   "Adiantamentos / Devoluções"
            Index           =   4
         End
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "&Devolução de Cheque"
         Index           =   15
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "Excluir Bordero de Cheque Pré"
         Index           =   16
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "Cobrança"
         Index           =   17
      End
      Begin VB.Menu mnuCRMov 
         Caption         =   "Histórico de Cobrança dos Clientes"
         Index           =   18
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   5
      Begin VB.Menu mnuESTMov 
         Caption         =   "Recebimento de Material de Fornecedor"
         Index           =   1
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Recebimento de &Material de Cliente"
         Index           =   2
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Recebimento de Material de Fornecedor / Compras"
         Index           =   3
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Requisição para Produção"
         Index           =   5
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Requisição de &Consumo"
         Index           =   6
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "&Movimento de Materiais"
         Index           =   7
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "&Transferência"
         Index           =   8
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "&Ordem de Produção"
         Index           =   10
         Begin VB.Menu mnuESTMovOP 
            Caption         =   "Manual"
            Index           =   1
         End
         Begin VB.Menu mnuESTMovOP 
            Caption         =   "Automático"
            Index           =   2
         End
         Begin VB.Menu mnuESTMovOP 
            Caption         =   "Ajuste de Empenho"
            Index           =   3
         End
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Produção - Entrada"
         Index           =   11
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "R&eservas"
         Index           =   12
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "&Inventário"
         Index           =   14
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Inventário &Lote"
         Index           =   15
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Inventário Em/De Terceiros"
         Index           =   16
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Entrada - Nota Fiscal Simples "
         Index           =   18
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Entrada - N. F. Simples / Compras"
         Index           =   19
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Entrada - Nota Fiscal Fatura "
         Index           =   20
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Entrada - N. F. Fatura / Compras"
         Index           =   21
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Entrada - Nota Fiscal Remessa "
         Index           =   22
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Entrada - Nota Fiscal Devolução"
         Index           =   23
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Cancelar Nota Fiscal"
         Index           =   24
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Medição de Contratos a Pagar"
         Index           =   25
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "-"
         Index           =   26
      End
      Begin VB.Menu mnuESTMov 
         Caption         =   "Desmembramento"
         Index           =   27
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   6
      Begin VB.Menu mnuFATMov2 
         Caption         =   "&Orçamento de Venda"
         Index           =   5
      End
      Begin VB.Menu mnuFATMov2 
         Caption         =   "Acompanhamento - Orçamento"
         Index           =   10
      End
      Begin VB.Menu mnuFATMov2 
         Caption         =   "&Pedido de Vendas"
         Index           =   15
      End
      Begin VB.Menu mnuFATMov2 
         Caption         =   "&Liberação de Bloqueios"
         Index           =   20
      End
      Begin VB.Menu mnuFATMov2 
         Caption         =   "&Baixa Manual de Pedido"
         Index           =   25
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "&Notas Fiscais - a Partir de vários Pedidos"
         Index           =   6
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Nota Fiscal &Simples - a Partir de um Pedido"
         Index           =   7
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Nota Fiscal &Fatura - a Partir de um Pedido"
         Index           =   8
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Nota Fiscal"
         Index           =   9
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Nota Fiscal Fatura"
         Index           =   10
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Nota Fiscal Remessa - a Partir de um Pedido"
         Index           =   12
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Nota Fiscal de Remessa"
         Index           =   13
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Nota Fiscal de Devolução"
         Index           =   14
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Cancelar Nota Fiscal"
         Index           =   15
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "&Geração de Faturas"
         Index           =   17
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "&Comissões"
         Index           =   18
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Para Transportadoras"
         Index           =   19
         Begin VB.Menu mnuFATMovTransp 
            Caption         =   "Pedido de Cotação"
            Index           =   1
         End
         Begin VB.Menu mnuFATMovTransp 
            Caption         =   "Proposta / Cotação"
            Index           =   2
         End
         Begin VB.Menu mnuFATMovTransp 
            Caption         =   "Solicitação de Serviço"
            Index           =   3
         End
         Begin VB.Menu mnuFATMovTransp 
            Caption         =   "Comprovante de Serviço"
            Index           =   4
         End
         Begin VB.Menu mnuFATMovTransp 
            Caption         =   "Conhecimento de &Transporte"
            Index           =   5
         End
         Begin VB.Menu mnuFATMovTransp 
            Caption         =   "Conhecimento de Transporte Fatura"
            Index           =   6
         End
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Medição de Contratos a Receber"
         Index           =   21
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Travel Ace"
         Index           =   22
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "Voucher - Comissão"
            Index           =   5
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "Ocorrências"
            Index           =   15
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "Liberação de Ocorrências"
            Index           =   20
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "-"
            Index           =   25
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "Acordos"
            Index           =   30
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "-"
            Index           =   35
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "Aportes"
            Index           =   40
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "Liberação de Aportes"
            Index           =   45
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "-"
            Index           =   50
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "Faturamento - Normal"
            Index           =   55
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "Faturamento - Cartão"
            Index           =   60
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "Cancelar Fatura\Nota de Crédito"
            Index           =   65
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "-"
            Index           =   70
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "Geração de Notas Fiscais"
            Index           =   75
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "-"
            Index           =   80
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "Acompanhamento da Assistência"
            Index           =   85
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "Liberação Pagto Cobertura"
            Index           =   90
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "Liberação Pagto Judicial"
            Index           =   95
         End
         Begin VB.Menu mnuTRVFATMov 
            Caption         =   "Reembolso de antecipação de pagto de seguro"
            Index           =   100
         End
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "-"
         Index           =   23
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Aportes"
         Index           =   24
         Begin VB.Menu mnuFATMovAporte 
            Caption         =   "Emissão"
            Index           =   1
         End
         Begin VB.Menu mnuFATMovAporte 
            Caption         =   "Liberação"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Vouchers"
         Index           =   25
         Begin VB.Menu mnuFATMovVou 
            Caption         =   "Emissão"
            Index           =   1
         End
         Begin VB.Menu mnuFATMovVou 
            Caption         =   "Manutenção"
            Index           =   2
         End
         Begin VB.Menu mnuFATMovVou 
            Caption         =   "Alteração de Comissão"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Ocorrências"
         Index           =   26
         Begin VB.Menu mnuFATMovOcr 
            Caption         =   "Emissão"
            Index           =   1
         End
         Begin VB.Menu mnuFATMovOcr 
            Caption         =   "Liberação"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Faturamento"
         Index           =   27
         Begin VB.Menu mnuFATMovFat 
            Caption         =   "Geração"
            Index           =   1
         End
         Begin VB.Menu mnuFATMovFat 
            Caption         =   "Cancelamento"
            Index           =   2
         End
         Begin VB.Menu mnuFATMovFat 
            Caption         =   "Cartão"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Geração de NFs Por Faturas"
         Index           =   28
      End
      Begin VB.Menu mnuFATMov 
         Caption         =   "Mapa de Entrega"
         Index           =   35
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   7
      Begin VB.Menu mnuCOMMov 
         Caption         =   "&Requisição"
         Index           =   1
         Begin VB.Menu mnuCOMMovReq 
            Caption         =   "&Não Enviada"
            Index           =   1
         End
         Begin VB.Menu mnuCOMMovReq 
            Caption         =   "&Baixar"
            Index           =   2
         End
         Begin VB.Menu mnuCOMMovReq 
            Caption         =   "&Enviar"
            Index           =   3
         End
         Begin VB.Menu mnuCOMMovReq 
            Caption         =   "&Aprovar"
            Index           =   4
         End
      End
      Begin VB.Menu mnuCOMMov 
         Caption         =   "G&eração de Requisição"
         Index           =   2
         Begin VB.Menu mnuCOMMovGerReq 
            Caption         =   "Por &Ponto de Pedido"
            Index           =   1
         End
         Begin VB.Menu mnuCOMMovGerReq 
            Caption         =   "Por Pedido de &Venda"
            Index           =   2
         End
      End
      Begin VB.Menu mnuCOMMov 
         Caption         =   "&Cotação"
         Index           =   3
         Begin VB.Menu mnuCOMMovCotacao 
            Caption         =   "Gerar &Pedido de Cotação"
            Index           =   1
         End
         Begin VB.Menu mnuCOMMovCotacao 
            Caption         =   "Gerar Pedido de &Cotação Avulso"
            Index           =   2
         End
         Begin VB.Menu mnuCOMMovCotacao 
            Caption         =   "Atualizar Cotação"
            Index           =   3
         End
         Begin VB.Menu mnuCOMMovCotacao 
            Caption         =   "Baixar Pedidos de Cotação"
            Index           =   4
         End
         Begin VB.Menu mnuCOMMovCotacao 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuCOMMovCotacao 
            Caption         =   "Mapa de Cotação"
            Index           =   6
         End
      End
      Begin VB.Menu mnuCOMMov 
         Caption         =   "&Geração de Pedido de Compra"
         Index           =   4
         Begin VB.Menu mnuCOMMovGerPC 
            Caption         =   "Por &Geração de Cotação"
            Index           =   1
         End
         Begin VB.Menu mnuCOMMovGerPC 
            Caption         =   "Por &Requisição"
            Index           =   2
         End
         Begin VB.Menu mnuCOMMovGerPC 
            Caption         =   "Por &Concorrência"
            Index           =   3
         End
         Begin VB.Menu mnuCOMMovGerPC 
            Caption         =   "&Avulsa"
            Index           =   4
         End
         Begin VB.Menu mnuCOMMovGerPC 
            Caption         =   "Por &Orçamento de Venda"
            Index           =   5
         End
      End
      Begin VB.Menu mnuCOMMov 
         Caption         =   "&Pedido de Compra"
         Index           =   5
         Begin VB.Menu mnuCOMMovPC 
            Caption         =   "&Avulso"
            Index           =   1
         End
         Begin VB.Menu mnuCOMMovPC 
            Caption         =   "&Gerado"
            Index           =   2
         End
         Begin VB.Menu mnuCOMMovPC 
            Caption         =   "&Liberar"
            Index           =   3
         End
         Begin VB.Menu mnuCOMMovPC 
            Caption         =   "&Baixar"
            Index           =   4
         End
         Begin VB.Menu mnuCOMMovPC 
            Caption         =   "Aprovar"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   8
      Begin VB.Menu mnuFISMov 
         Caption         =   "Registro de Entrada"
         Index           =   1
      End
      Begin VB.Menu mnuFISMov 
         Caption         =   "Registro de Saida"
         Index           =   2
      End
      Begin VB.Menu mnuFISMov 
         Caption         =   "ICMS"
         Index           =   3
         Begin VB.Menu mnuFISMovICMS 
            Caption         =   "Apuração de ICMS"
            Index           =   1
         End
         Begin VB.Menu mnuFISMovICMS 
            Caption         =   "Lançamentos para Apuração"
            Index           =   5
         End
         Begin VB.Menu mnuFISMovICMS 
            Caption         =   "Dados de Recolhimento GNR"
            Index           =   6
         End
         Begin VB.Menu mnuFISMovICMS 
            Caption         =   "Registro de Inventário"
            Index           =   7
         End
         Begin VB.Menu mnuFISMovICMS 
            Caption         =   "Guias de Recolhimento"
            Index           =   8
         End
         Begin VB.Menu mnuFISMovICMS 
            Caption         =   "Guias de Recolhimento ST"
            Index           =   9
         End
      End
      Begin VB.Menu mnuFISMov 
         Caption         =   "IPI"
         Index           =   4
         Begin VB.Menu mnuFISMovIPI 
            Caption         =   "Apuração de IPI"
            Index           =   1
         End
         Begin VB.Menu mnuFISMovIPI 
            Caption         =   "Lançamentos para Apuração"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFISMov 
         Caption         =   "PIS"
         Index           =   5
         Begin VB.Menu mnuFISMovPIS 
            Caption         =   "Apuração de PIS"
            Index           =   1
         End
      End
      Begin VB.Menu mnuFISMov 
         Caption         =   "COFINS"
         Index           =   6
         Begin VB.Menu mnuFISMovCOFINS 
            Caption         =   "Apuração de COFINS"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   9
      Begin VB.Menu MnuLjMov 
         Caption         =   "Borderô Boletos Manuais"
         Index           =   1
      End
      Begin VB.Menu MnuLjMov 
         Caption         =   "Borderô Cheques"
         Index           =   2
      End
      Begin VB.Menu MnuLjMov 
         Caption         =   "Borderô Vales / Tickets"
         Index           =   3
      End
      Begin VB.Menu MnuLjMov 
         Caption         =   "Borderô Outros"
         Index           =   4
      End
      Begin VB.Menu MnuLjMov 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MnuLjMov 
         Caption         =   "Depósito Bancário"
         Index           =   6
      End
      Begin VB.Menu MnuLjMov 
         Caption         =   "Depósito em Caixa"
         Index           =   7
      End
      Begin VB.Menu MnuLjMov 
         Caption         =   "Saque de Caixa"
         Index           =   8
      End
      Begin VB.Menu MnuLjMov 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu MnuLjMov 
         Caption         =   "Especificação de Cheques"
         Index           =   10
      End
      Begin VB.Menu MnuLjMov 
         Caption         =   "Transferência de Meios de Pagamento"
         Index           =   11
      End
      Begin VB.Menu MnuLjMov 
         Caption         =   "Recebimento de Carnê"
         Index           =   12
      End
      Begin VB.Menu MnuLjMov 
         Caption         =   "Cancela Recebimento de Carnê"
         Index           =   13
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   10
      Begin VB.Menu mnuPCPMov 
         Caption         =   "Ordem de Produção"
         Index           =   1
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "OPs por Pedidos de Venda"
         Index           =   3
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "Ajuste de Empenho"
         Index           =   4
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "Liberação de Bloqueio"
         Index           =   5
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "Requisição para Produção"
         Index           =   6
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "Transferência"
         Index           =   7
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "Produção - Entrada"
         Index           =   9
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "Apontamento da Produção"
         Index           =   11
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "Plano Mestre de Produção"
         Index           =   12
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "Ordem de Corte"
         Index           =   14
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "Requisição para produção D-Pack"
         Index           =   15
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "Ordem de Corte"
         Index           =   18
      End
      Begin VB.Menu mnuPCPMov 
         Caption         =   "Ordem de Corte - Manual"
         Index           =   19
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   11
      Begin VB.Menu mnuCRMMov 
         Caption         =   "Relacionamentos com Clientes"
         Index           =   1
      End
      Begin VB.Menu mnuCRMMov 
         Caption         =   "Follow-Up"
         Index           =   2
      End
      Begin VB.Menu mnuCRMMov 
         Caption         =   "Relacionamentos com Clientes Futuros"
         Index           =   3
      End
      Begin VB.Menu mnuCRMMov 
         Caption         =   "Follow-Up - Clientes Futuros"
         Index           =   4
      End
      Begin VB.Menu mnuCRMMov 
         Caption         =   "Contatos com Clientes - Call Center"
         Index           =   5
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   12
      Begin VB.Menu mnuQUAMov 
         Caption         =   "Resultados dos Testes"
         Index           =   1
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   13
      Begin VB.Menu mnuPRJMov 
         Caption         =   "Propostas"
         Index           =   1
      End
      Begin VB.Menu mnuPRJMov 
         Caption         =   "Contratos"
         Index           =   2
      End
      Begin VB.Menu mnuPRJMov 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPRJMov 
         Caption         =   "Pagamentos"
         Index           =   4
      End
      Begin VB.Menu mnuPRJMov 
         Caption         =   "Recebimentos"
         Index           =   5
      End
      Begin VB.Menu mnuPRJMov 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuPRJMov 
         Caption         =   "Apontamento"
         Index           =   7
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   14
      Begin VB.Menu mnuSRVMov 
         Caption         =   "&Solicitação de Serviço"
         Index           =   1
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "&Orçamento"
         Index           =   10
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "Liberação de Orçamento"
         Index           =   11
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "&Pedido de Serviço"
         Index           =   20
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "Liberação de Pedido"
         Index           =   21
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "Baixa Manual de Pedido de Serviço"
         Index           =   22
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "-"
         Index           =   25
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "&Ordem de Serviço"
         Index           =   30
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "&Apontamento"
         Index           =   31
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "&Movimentos de Materiais"
         Index           =   32
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "-"
         Index           =   35
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "&Nota Fiscal"
         Index           =   40
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "&Nota Fiscal Fatura"
         Index           =   41
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "&Nota Fiscal - A Partir de um Pedido"
         Index           =   42
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "&Nota Fiscal Fatura - A Partir de um Pedido"
         Index           =   43
      End
      Begin VB.Menu mnuSRVMov 
         Caption         =   "&Nota Fiscal Fatura - Itens em Garantia"
         Index           =   46
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   1
      Begin VB.Menu mnuCTBCon 
         Caption         =   "&Plano de Contas"
         Index           =   1
      End
      Begin VB.Menu mnuCTBCon 
         Caption         =   "&Centro de Custo Lucro"
         Index           =   2
      End
      Begin VB.Menu mnuCTBCon 
         Caption         =   "C&onta x Centro de Custo/Lucro"
         Index           =   3
      End
      Begin VB.Menu mnuCTBCon 
         Caption         =   "&Histórico Padrao"
         Index           =   4
      End
      Begin VB.Menu mnuCTBCon 
         Caption         =   "&Lotes"
         Index           =   5
      End
      Begin VB.Menu mnuCTBCon 
         Caption         =   "Lotes P&endentes"
         Index           =   6
      End
      Begin VB.Menu mnuCTBCon 
         Caption         =   "L&ançamentos"
         Index           =   7
      End
      Begin VB.Menu mnuCTBCon 
         Caption         =   "Lançamentos Pe&ndentes"
         Index           =   8
      End
      Begin VB.Menu mnuCTBCon 
         Caption         =   "&Documentos Automáticos"
         Index           =   9
      End
      Begin VB.Menu mnuCTBCon 
         Caption         =   "&Orçamento"
         Index           =   10
      End
      Begin VB.Menu mnuCTBCon 
         Caption         =   "Rateio On-Line"
         Index           =   11
      End
      Begin VB.Menu mnuCTBCon 
         Caption         =   "RateioOff"
         Index           =   12
      End
      Begin VB.Menu mnuCTBCon 
         Caption         =   "Mapeamento do Plano de Contas Referencial"
         Index           =   13
      End
      Begin VB.Menu mnuCTBCon 
         Caption         =   "Associação com as contas referenciais"
         Index           =   14
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   2
      Begin VB.Menu mnuTESCon 
         Caption         =   "&Fluxo de Caixa"
         Index           =   1
      End
      Begin VB.Menu mnuTESCon 
         Caption         =   "&Aplicações"
         Index           =   2
      End
      Begin VB.Menu mnuTESCon 
         Caption         =   "&Bancos"
         Index           =   3
      End
      Begin VB.Menu mnuTESCon 
         Caption         =   "&Contas Correntes"
         Index           =   4
      End
      Begin VB.Menu mnuTESCon 
         Caption         =   "Contas Corrente de Todas as &Filiais"
         Index           =   5
      End
      Begin VB.Menu mnuTESCon 
         Caption         =   "&Depósitos"
         Index           =   6
      End
      Begin VB.Menu mnuTESCon 
         Caption         =   "&Saques"
         Index           =   7
      End
      Begin VB.Menu mnuTESCon 
         Caption         =   "Ti&pos de Aplicações"
         Index           =   8
      End
      Begin VB.Menu mnuTESCon 
         Caption         =   "Transferências"
         Index           =   9
      End
      Begin VB.Menu mnuTESCon 
         Caption         =   "Movimentos de Conta Corrente"
         Index           =   10
      End
      Begin VB.Menu mnuTESCon 
         Caption         =   "Movimento C.C. - Empresa Toda"
         Index           =   11
      End
      Begin VB.Menu mnuTESCon 
         Caption         =   "Fluxo de Caixa Contábil"
         Index           =   12
      End
      Begin VB.Menu mnuTESCon 
         Caption         =   "Fluxo de Caixa Contábil no Excel"
         Index           =   13
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   3
      Begin VB.Menu mnuCPCon 
         Caption         =   "Cadastros"
         Index           =   1
         Begin VB.Menu mnuCPConCad 
            Caption         =   "F&ornecedores"
            Index           =   1
         End
         Begin VB.Menu mnuCPConCad 
            Caption         =   "&Tipos de Fornecedores"
            Index           =   2
         End
         Begin VB.Menu mnuCPConCad 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuCPConCad 
            Caption         =   "Bancos"
            Index           =   4
         End
         Begin VB.Menu mnuCPConCad 
            Caption         =   "Contas Correntes"
            Index           =   5
         End
         Begin VB.Menu mnuCPConCad 
            Caption         =   "Portadores"
            Index           =   6
         End
         Begin VB.Menu mnuCPConCad 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuCPConCad 
            Caption         =   "&Condições de Pagamento"
            Index           =   8
         End
      End
      Begin VB.Menu mnuCPCon 
         Caption         =   "À Partir de Fornecedor"
         Index           =   2
      End
      Begin VB.Menu mnuCPCon 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCPCon 
         Caption         =   "Títulos a Pagar"
         Index           =   4
         Begin VB.Menu mnuCPConTP 
            Caption         =   "Aberto"
            Index           =   1
         End
         Begin VB.Menu mnuCPConTP 
            Caption         =   "Todos"
            Index           =   3
         End
         Begin VB.Menu mnuCPConTP 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuCPConTP 
            Caption         =   "Todos - Empresa Toda"
            Index           =   5
         End
         Begin VB.Menu mnuCPConTP 
            Caption         =   "Atrasados"
            Index           =   6
         End
         Begin VB.Menu mnuCPConTP 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuCPConTP 
            Caption         =   "Baixas"
            Index           =   8
         End
         Begin VB.Menu mnuCPConTP 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnuCPConTP 
            Caption         =   "Naka"
            Index           =   10
         End
      End
      Begin VB.Menu mnuCPCon 
         Caption         =   "&NFs Simples à Pagar"
         Index           =   6
         Begin VB.Menu mnuCPConNF 
            Caption         =   "Abertas"
            Index           =   1
         End
         Begin VB.Menu mnuCPConNF 
            Caption         =   "Todas"
            Index           =   2
         End
         Begin VB.Menu mnuCPConNF 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuCPConNF 
            Caption         =   "Todas - Empresa Toda"
            Index           =   4
         End
         Begin VB.Menu mnuCPConNF 
            Caption         =   "Todas -> Fatura -> Baixas"
            Index           =   5
         End
         Begin VB.Menu mnuCPConNF 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuCPConNF 
            Caption         =   "NFs de CMCC sem Fatura"
            Index           =   7
         End
      End
      Begin VB.Menu mnuCPCon 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuCPCon 
         Caption         =   "C&réditos"
         Index           =   9
      End
      Begin VB.Menu mnuCPCon 
         Caption         =   "Adiantamentos"
         Index           =   10
      End
      Begin VB.Menu mnuCPCon 
         Caption         =   "P&agamentos"
         Index           =   11
      End
      Begin VB.Menu mnuCPCon 
         Caption         =   "Cheques"
         Index           =   12
         Begin VB.Menu mnuCPConCHQ 
            Caption         =   "Emitidos"
            Index           =   1
         End
         Begin VB.Menu mnuCPConCHQ 
            Caption         =   "A compensar"
            Index           =   2
         End
         Begin VB.Menu mnuCPConCHQ 
            Caption         =   "Parcelas"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   4
      Begin VB.Menu mnuCRCon 
         Caption         =   "Cadastros"
         Index           =   1
         Begin VB.Menu mnuCRConCad 
            Caption         =   "&Clientes"
            Index           =   1
         End
         Begin VB.Menu mnuCRConCad 
            Caption         =   "C&obradores"
            Index           =   2
         End
         Begin VB.Menu mnuCRConCad 
            Caption         =   "V&endedores"
            Index           =   3
         End
         Begin VB.Menu mnuCRConCad 
            Caption         =   "Transportadoras"
            Index           =   4
         End
         Begin VB.Menu mnuCRConCad 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuCRConCad 
            Caption         =   "Co&ndições de Pagamento"
            Index           =   6
         End
      End
      Begin VB.Menu mnuCRCon 
         Caption         =   "À Partir de Cliente"
         Index           =   2
      End
      Begin VB.Menu mnuCRCon 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCRCon 
         Caption         =   "Tít&ulos a Receber"
         Index           =   4
         Begin VB.Menu mnuCRConTR 
            Caption         =   "Abertos"
            Index           =   1
         End
         Begin VB.Menu mnuCRConTR 
            Caption         =   "Baixados"
            Index           =   2
         End
         Begin VB.Menu mnuCRConTR 
            Caption         =   "Todos"
            Index           =   3
         End
         Begin VB.Menu mnuCRConTR 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuCRConTR 
            Caption         =   "Abertos - Empresa toda"
            Index           =   5
         End
         Begin VB.Menu mnuCRConTR 
            Caption         =   "Baixados - Empresa toda"
            Index           =   6
         End
         Begin VB.Menu mnuCRConTR 
            Caption         =   "Todos - Empresa toda"
            Index           =   7
         End
         Begin VB.Menu mnuCRConTR 
            Caption         =   "Atrasados"
            Index           =   8
         End
         Begin VB.Menu mnuCRConTR 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnuCRConTR 
            Caption         =   "Baixas"
            Index           =   10
         End
         Begin VB.Menu mnuCRConTR 
            Caption         =   "-"
            Index           =   11
         End
         Begin VB.Menu mnuCRConTR 
            Caption         =   "Títulos de Cartão x Vouchers"
            Index           =   12
         End
         Begin VB.Menu mnuCRConTR 
            Caption         =   "Boletos Impressos e Cancelados"
            Index           =   13
         End
         Begin VB.Menu mnuCRConTR 
            Caption         =   "Comissões na baixa"
            Index           =   14
         End
      End
      Begin VB.Menu mnuCRCon 
         Caption         =   "&Devoluções"
         Index           =   7
      End
      Begin VB.Menu mnuCRCon 
         Caption         =   "&Adiantamentos"
         Index           =   8
      End
      Begin VB.Menu mnuCRCon 
         Caption         =   "&Borderô de Cobrança"
         Index           =   9
      End
      Begin VB.Menu mnuCRCon 
         Caption         =   "Cheques Pré"
         Index           =   10
         Begin VB.Menu mnuCRConCHQ 
            Caption         =   "Emitidos"
            Index           =   1
         End
         Begin VB.Menu mnuCRConCHQ 
            Caption         =   "Parcelas"
            Index           =   2
         End
      End
      Begin VB.Menu mnuCRCon 
         Caption         =   "A Receber/Recebido por Produto"
         Index           =   11
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   5
      Begin VB.Menu mnuESTCon 
         Caption         =   "Cadastros"
         Index           =   1
         Begin VB.Menu mnuESTConCad 
            Caption         =   "Fornecedores"
            Index           =   1
         End
         Begin VB.Menu mnuESTConCad 
            Caption         =   "P&rodutos"
            Index           =   2
         End
         Begin VB.Menu mnuESTConCad 
            Caption         =   "&Almoxarifado"
            Index           =   3
         End
         Begin VB.Menu mnuESTConCad 
            Caption         =   "&Kit"
            Index           =   4
         End
         Begin VB.Menu mnuESTConCad 
            Caption         =   "&Tipos de Produtos"
            Index           =   5
         End
         Begin VB.Menu mnuESTConCad 
            Caption         =   "&Classes de Unidades de Medidas"
            Index           =   6
         End
         Begin VB.Menu mnuESTConCad 
            Caption         =   "Cate&gorias de Produtos"
            Index           =   7
         End
         Begin VB.Menu mnuESTConCad 
            Caption         =   "Itens de Categorias de Produto"
            Index           =   8
         End
         Begin VB.Menu mnuESTConCad 
            Caption         =   "Produtos x Itens de Categorias de Produto"
            Index           =   9
         End
      End
      Begin VB.Menu mnuESTCon 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuESTCon 
         Caption         =   "Movimentos"
         Index           =   3
         Begin VB.Menu mnuESTConMov 
            Caption         =   "Recebimento de Material "
            Index           =   1
         End
         Begin VB.Menu mnuESTConMov 
            Caption         =   "Recebimento de Material do Fornecedor"
            Index           =   2
         End
         Begin VB.Menu mnuESTConMov 
            Caption         =   "Recebimento de Material de Cliente"
            Index           =   3
         End
         Begin VB.Menu mnuESTConMov 
            Caption         =   "Reser&vas"
            Index           =   4
         End
         Begin VB.Menu mnuESTConMov 
            Caption         =   "Movimentos Estoque"
            Index           =   5
         End
         Begin VB.Menu mnuESTConMov 
            Caption         =   "Movimentos de Estoque Interno"
            Index           =   6
         End
         Begin VB.Menu mnuESTConMov 
            Caption         =   "Movimentos de Estoque - Transferência"
            Index           =   7
         End
         Begin VB.Menu mnuESTConMov 
            Caption         =   "Requisição para Consumo"
            Index           =   8
         End
      End
      Begin VB.Menu mnuESTCon 
         Caption         =   "Inventário"
         Index           =   4
         Begin VB.Menu mnuESTConInv 
            Caption         =   "&Inventários"
            Index           =   1
         End
         Begin VB.Menu mnuESTConInv 
            Caption         =   "Inventários Pendentes"
            Index           =   2
         End
         Begin VB.Menu mnuESTConInv 
            Caption         =   "Lotes de Inventário"
            Index           =   3
         End
      End
      Begin VB.Menu mnuESTCon 
         Caption         =   "Produção"
         Index           =   5
         Begin VB.Menu mnuESTConPro 
            Caption         =   "&Ordens de Produção"
            Index           =   1
         End
         Begin VB.Menu mnuESTConPro 
            Caption         =   "&Ordens de Produção Baixadas"
            Index           =   2
         End
         Begin VB.Menu mnuESTConPro 
            Caption         =   "Itens &Empenhados"
            Index           =   3
         End
         Begin VB.Menu mnuESTConPro 
            Caption         =   "Itens em Produção"
            Index           =   4
         End
         Begin VB.Menu mnuESTConPro 
            Caption         =   "Requisição para Produção"
            Index           =   5
         End
         Begin VB.Menu mnuESTConPro 
            Caption         =   "Entrada de Produção"
            Index           =   6
         End
         Begin VB.Menu mnuESTConPro 
            Caption         =   "Itens Requisitados p/ Produção"
            Index           =   7
         End
         Begin VB.Menu mnuESTConPro 
            Caption         =   "Itens Produzidos"
            Index           =   8
         End
      End
      Begin VB.Menu mnuESTCon 
         Caption         =   "Controle de Estoque"
         Index           =   6
         Begin VB.Menu mnuESTConEst 
            Caption         =   "Estoque da Empresa"
            Index           =   1
         End
         Begin VB.Menu mnuESTConEst 
            Caption         =   "Estoque da Filial"
            Index           =   2
         End
         Begin VB.Menu mnuESTConEst 
            Caption         =   "Estoque em Terceiros"
            Index           =   3
         End
         Begin VB.Menu mnuESTConEst 
            Caption         =   "Saldo Disponível"
            Index           =   4
         End
         Begin VB.Menu mnuESTConEst 
            Caption         =   "Estoque da Empresa - Média de Vendas"
            Index           =   5
         End
         Begin VB.Menu mnuESTConEst 
            Caption         =   "Produtos em Falta"
            Index           =   6
         End
      End
      Begin VB.Menu mnuESTCon 
         Caption         =   "Notas Fiscais - Entrada"
         Index           =   7
         Begin VB.Menu mnuESTConNF 
            Caption         =   "Todas as Notas Fiscais"
            Index           =   1
         End
         Begin VB.Menu mnuESTConNF 
            Caption         =   "Notas Fiscais Simples"
            Index           =   2
         End
         Begin VB.Menu mnuESTConNF 
            Caption         =   "Notas Fiscais Fatura"
            Index           =   3
         End
         Begin VB.Menu mnuESTConNF 
            Caption         =   "Notas Fiscais de Remessa"
            Index           =   4
         End
         Begin VB.Menu mnuESTConNF 
            Caption         =   "Notas Fiscais de Devolução"
            Index           =   5
         End
         Begin VB.Menu mnuESTConNF 
            Caption         =   "Itens das Notas Fiscais"
            Index           =   6
         End
      End
      Begin VB.Menu mnuESTCon 
         Caption         =   "Notas Fiscais - Saída"
         Index           =   8
         Begin VB.Menu mnuESTConNFSai 
            Caption         =   "Todas as Notas Fiscais"
            Index           =   1
         End
         Begin VB.Menu mnuESTConNFSai 
            Caption         =   "Notas Fiscais Simples"
            Index           =   2
         End
         Begin VB.Menu mnuESTConNFSai 
            Caption         =   "Notas Fiscais Fatura"
            Index           =   3
         End
         Begin VB.Menu mnuESTConNFSai 
            Caption         =   "Notas Fiscais Simples gerada por Pedido"
            Index           =   4
         End
         Begin VB.Menu mnuESTConNFSai 
            Caption         =   "Notas Fiscais Fatura gerada por Pedido"
            Index           =   5
         End
         Begin VB.Menu mnuESTConNFSai 
            Caption         =   "Notas Fiscais de Remessa"
            Index           =   6
         End
         Begin VB.Menu mnuESTConNFSai 
            Caption         =   "Notas Fiscais de Devolução"
            Index           =   7
         End
         Begin VB.Menu mnuESTConNFSai 
            Caption         =   "Itens das Notas Fiscais"
            Index           =   8
         End
      End
      Begin VB.Menu mnuESTCon 
         Caption         =   "Rastreamento"
         Index           =   9
         Begin VB.Menu mnuESTConRastro 
            Caption         =   "Lotes"
            Index           =   1
         End
         Begin VB.Menu mnuESTConRastro 
            Caption         =   "Saldos nos Lotes"
            Index           =   2
         End
         Begin VB.Menu mnuESTConRastro 
            Caption         =   "Movimentos"
            Index           =   3
         End
         Begin VB.Menu mnuESTConRastro 
            Caption         =   "Saldos x Preço"
            Index           =   10
         End
         Begin VB.Menu mnuESTConRastro 
            Caption         =   "Movimentos x Preço"
            Index           =   11
         End
         Begin VB.Menu mnuESTConRastro 
            Caption         =   "Itens de NF"
            Index           =   12
         End
         Begin VB.Menu mnuESTConRastro 
            Caption         =   "Separação LM Log"
            Index           =   13
         End
         Begin VB.Menu mnuESTConRastro 
            Caption         =   "Saldos nos Lotes - Custo Médio"
            Index           =   14
         End
      End
      Begin VB.Menu mnuESTCon 
         Caption         =   "Travel Ace"
         Index           =   10
         Begin VB.Menu mnuESTConTRV 
            Caption         =   "Movimentos de Estoque - Todos"
            Index           =   1
         End
         Begin VB.Menu mnuESTConTRV 
            Caption         =   "Movimentos de Estoque - Transferências"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   6
      Begin VB.Menu mnuFATCon 
         Caption         =   "Cadastros"
         Index           =   1
         Begin VB.Menu mnuFATConCad 
            Caption         =   "Clientes"
            Index           =   1
         End
         Begin VB.Menu mnuFATConCad 
            Caption         =   "Clientes - Consultas"
            Index           =   2
         End
         Begin VB.Menu mnuFATConCad 
            Caption         =   "Vendedores"
            Index           =   3
         End
         Begin VB.Menu mnuFATConCad 
            Caption         =   "Transportadoras"
            Index           =   4
         End
         Begin VB.Menu mnuFATConCad 
            Caption         =   "Categorias de Produtos"
            Index           =   5
         End
         Begin VB.Menu mnuFATConCad 
            Caption         =   "Itens de Categorias de Produto"
            Index           =   6
         End
         Begin VB.Menu mnuFATConCad 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuFATConCad 
            Caption         =   "Emissores"
            Index           =   8
         End
         Begin VB.Menu mnuFATConCad 
            Caption         =   "Acordos"
            Index           =   9
         End
         Begin VB.Menu mnuFATConCad 
            Caption         =   "Itens de Contrato à Faturar"
            Index           =   10
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Vendas"
         Index           =   3
         Begin VB.Menu mnuFATConVen 
            Caption         =   "Pedidos Ativos"
            Index           =   1
         End
         Begin VB.Menu mnuFATConVen 
            Caption         =   "Pedidos Baixados"
            Index           =   2
         End
         Begin VB.Menu mnuFATConVen 
            Caption         =   "Tabelas de Preço"
            Index           =   3
         End
         Begin VB.Menu mnuFATConVen 
            Caption         =   "Tabelas de Preço - Itens"
            Index           =   4
         End
         Begin VB.Menu mnuFATConVen 
            Caption         =   "&Previsão de Venda"
            Index           =   5
         End
         Begin VB.Menu mnuFATConVen 
            Caption         =   "Orçamentos"
            Index           =   6
         End
         Begin VB.Menu mnuFATConVen 
            Caption         =   "Tabelas de Preço - Itens - Atual"
            Index           =   7
         End
         Begin VB.Menu mnuFATConVen 
            Caption         =   "Rotas de Venda"
            Index           =   10
         End
         Begin VB.Menu mnuFATConVen 
            Caption         =   "Mapa de Entrega"
            Index           =   11
         End
         Begin VB.Menu mnuFATConVen 
            Caption         =   "Acompanhamento"
            Index           =   12
         End
         Begin VB.Menu mnuFATConVen 
            Caption         =   "Itens dos Pedidos de Venda"
            Index           =   13
         End
         Begin VB.Menu mnuFATConVen 
            Caption         =   "Itens dos Orçamentos de Venda"
            Index           =   14
         End
         Begin VB.Menu mnuFATConVen 
            Caption         =   "Protheus - Histórico de Vendas"
            Index           =   15
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Estoque"
         Index           =   4
         Begin VB.Menu mnuFATConEST 
            Caption         =   "Produto"
            Index           =   1
         End
         Begin VB.Menu mnuFATConEST 
            Caption         =   "Estoque &Produto Filial"
            Index           =   2
         End
         Begin VB.Menu mnuFATConEST 
            Caption         =   "&Estoque Produto Empresa"
            Index           =   3
         End
         Begin VB.Menu mnuFATConEST 
            Caption         =   "Estoque Produto Em &Terceiros"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Notas Fiscais - Saída"
         Index           =   5
         Begin VB.Menu mnuFATConNF 
            Caption         =   "Todas as Notas Fiscais"
            Index           =   1
         End
         Begin VB.Menu mnuFATConNF 
            Caption         =   "Notas Fiscais Simples"
            Index           =   2
         End
         Begin VB.Menu mnuFATConNF 
            Caption         =   "Notas Fiscais Fatura"
            Index           =   3
         End
         Begin VB.Menu mnuFATConNF 
            Caption         =   "Notas Fiscais Simples gerada por Pedido"
            Index           =   4
         End
         Begin VB.Menu mnuFATConNF 
            Caption         =   "Notas Fiscais Fatura gerada por Pedido"
            Index           =   5
         End
         Begin VB.Menu mnuFATConNF 
            Caption         =   "Notas Fiscais Remessa gerada por Pedido"
            Index           =   6
         End
         Begin VB.Menu mnuFATConNF 
            Caption         =   "Notas Fiscais de Remessa"
            Index           =   7
         End
         Begin VB.Menu mnuFATConNF 
            Caption         =   "Notas Fiscais de Devolução"
            Index           =   8
         End
         Begin VB.Menu mnuFATConNF 
            Caption         =   "&Séries de Notas Fiscais"
            Index           =   9
         End
         Begin VB.Menu mnuFATConNF 
            Caption         =   "Itens das Notas Fiscais"
            Index           =   10
         End
         Begin VB.Menu mnuFATConNF 
            Caption         =   "Conhecimentos de Transporte"
            Index           =   11
         End
         Begin VB.Menu mnuFATConNF 
            Caption         =   "Notas Fiscais - Transportadoras"
            Index           =   12
         End
         Begin VB.Menu mnuFATConNF 
            Caption         =   "Notas Fiscais a Faturar"
            Index           =   13
         End
         Begin VB.Menu mnuFATConNF 
            Caption         =   "Notas Fiscais - BI"
            Index           =   14
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Notas Fiscais - Entrada"
         Index           =   6
         Begin VB.Menu mnuFATConNFEnt 
            Caption         =   "Todas as Notas Fiscais"
            Index           =   1
         End
         Begin VB.Menu mnuFATConNFEnt 
            Caption         =   "Notas Fiscais Simples"
            Index           =   2
         End
         Begin VB.Menu mnuFATConNFEnt 
            Caption         =   "Notas Fiscais Fatura"
            Index           =   3
         End
         Begin VB.Menu mnuFATConNFEnt 
            Caption         =   "Notas Fiscais Remessa"
            Index           =   4
         End
         Begin VB.Menu mnuFATConNFEnt 
            Caption         =   "Notas Fiscais de Devolução"
            Index           =   5
         End
         Begin VB.Menu mnuFATConNFEnt 
            Caption         =   "Itens das Notas Fiscais"
            Index           =   6
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Gráficos"
         Index           =   8
         Begin VB.Menu mnuFATConGraf 
            Caption         =   "Faturamento por Área"
            Index           =   1
         End
         Begin VB.Menu mnuFATConGraf 
            Caption         =   "Faturamento por Cliente"
            Index           =   2
         End
         Begin VB.Menu mnuFATConGraf 
            Caption         =   "Faturamento - Comparativo Mensal em Dolar"
            Index           =   3
         End
         Begin VB.Menu mnuFATConGraf 
            Caption         =   "Faturamento - Comparativo Mensal"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Log"
         Index           =   10
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Controle de Notas Fiscais"
         Index           =   12
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Recibo Provisório de Serviço (RPS)"
         Index           =   14
         Begin VB.Menu mnuFATConRPS 
            Caption         =   "Recibos - Todos"
            Index           =   1
         End
         Begin VB.Menu mnuFATConRPS 
            Caption         =   "Recibos não Enviados"
            Index           =   2
         End
         Begin VB.Menu mnuFATConRPS 
            Caption         =   "Recibos não Convertidos em NFe"
            Index           =   3
         End
         Begin VB.Menu mnuFATConRPS 
            Caption         =   "Arquivos Gerados"
            Index           =   4
         End
         Begin VB.Menu mnuFATConRPS 
            Caption         =   "Histórico de Recibos Enviados"
            Index           =   5
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Notas Fiscais Eletrônica Municipal (NFe)"
         Index           =   15
         Begin VB.Menu mnuFATConNFe 
            Caption         =   "Notas Fiscais Eletrônicas - Todas"
            Index           =   1
         End
         Begin VB.Menu mnuFATConNFe 
            Caption         =   "Histórico de Notas Recebidas"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "TRV - Faturamento"
         Index           =   19
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Vouchers - Tela"
            Index           =   1
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Vouchers - Todos"
            Index           =   2
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Vouches não Faturados"
            Index           =   3
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Vouchers Cancelados - Previsão Reembolso"
            Index           =   4
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Vouchers por Representante"
            Index           =   5
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Vouchers por Correntista"
            Index           =   6
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Vouchers por Emissor"
            Index           =   7
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Ocorrências - Todas"
            Index           =   9
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Ocorrências Bloqueadas"
            Index           =   10
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Ocorrências Liberadas e não Faturadas"
            Index           =   11
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Ocorrências de Inativação"
            Index           =   12
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "-"
            Index           =   13
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Aportes"
            Index           =   25
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Pagamentos de Aportes - Todos"
            Index           =   26
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Pagamentos de Aportes não Faturados"
            Index           =   27
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Pagamentos de Aportes - Sobre Faturamento"
            Index           =   28
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Pagamentos de Aportes - Créditos a Faturar"
            Index           =   29
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Histórico de utilização dos Pagtos sobre Fat."
            Index           =   30
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Histórico de utilização dos Pagtos via crédito"
            Index           =   31
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "-"
            Index           =   32
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Títulos a Receber Sem Nota Fiscal"
            Index           =   33
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "-"
            Index           =   34
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Documentos a Serem Faturados (VOU, OCR, NVL, CMC, CMCC, CMR e OVER)"
            Index           =   35
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Documentos Faturados (VOU, OCR, NVL, CMC, CMCC, CMR e OVER)"
            Index           =   36
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "-"
            Index           =   39
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Faturas\Notas de Crédito - Todas"
            Index           =   40
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Faturas\Notas de Crédito - Canceladas"
            Index           =   41
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Faturas Cartão - Tarifas - Pagamento"
            Index           =   42
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "-"
            Index           =   45
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Vendas - Call Center"
            Index           =   50
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Posição Geral - Vouchers"
            Index           =   51
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Estatística de Venda"
            Index           =   55
         End
         Begin VB.Menu mnuTRVFATCon 
            Caption         =   "Vouchers x Faturas x Baixas"
            Index           =   60
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "TRV - Assistência"
         Index           =   23
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Todas Enviadas"
            Index           =   5
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Todas Abertas"
            Index           =   10
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Aguardando Documentos"
            Index           =   15
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Todas Autorizadas"
            Index           =   20
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Autorizadas e Não Faturadas"
            Index           =   25
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Todas Faturadas por Cobertura"
            Index           =   27
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Faturado Cobertura e Não Paga"
            Index           =   28
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Todas com Processo"
            Index           =   30
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Processos em Aberto"
            Index           =   35
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Todas Condenadas"
            Index           =   40
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Condenadas e Não Faturadas"
            Index           =   45
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Todas Faturadas por Processo"
            Index           =   47
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Faturado Processo e Não Pago"
            Index           =   50
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Todas com Reembolso"
            Index           =   55
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Ocorrências - Reembolsos não recebidos"
            Index           =   60
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "-"
            Index           =   65
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Titulos a Pagar - Cobertura"
            Index           =   70
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Titulos a Pagar - Condenação"
            Index           =   75
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Titulos a Receber - Reembolso"
            Index           =   80
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "-"
            Index           =   85
         End
         Begin VB.Menu mnuTRVFATCon1 
            Caption         =   "Valores Liberados"
            Index           =   90
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "TRV - Comissionamento"
         Index           =   25
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Acordos"
            Index           =   1
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Comissões do Representante - Todas"
            Index           =   14
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Comissões do Representante - Bloqueadas"
            Index           =   15
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Comissões do Representante - Liberadas e não Faturadas"
            Index           =   16
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "-"
            Index           =   19
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Comissões do Correntista - Todas"
            Index           =   22
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Comissões do Correntista - Bloqueadas"
            Index           =   23
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Comissões do Correntista - Liberadas e não Faturadas"
            Index           =   24
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "-"
            Index           =   28
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Comissões de Cartão de Crédito- Todas"
            Index           =   30
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Comissões de Cartão de Crédito- Bloqueadas"
            Index           =   31
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Comissões de Cartão de Crédito- Liberadas e não Faturadas"
            Index           =   32
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "-"
            Index           =   34
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Comissões Retidas nas Agências"
            Index           =   35
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "-"
            Index           =   36
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Over - Todos"
            Index           =   37
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Over - Bloqueado"
            Index           =   38
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Over - Liberados e não Faturados"
            Index           =   39
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Over - Faturados"
            Index           =   40
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Over - Por Emissor- Emitidos Hoje"
            Index           =   41
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "-"
            Index           =   43
         End
         Begin VB.Menu mnuTRVFATCon2 
            Caption         =   "Comissões aguardando envio de NF"
            Index           =   45
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "-"
         Index           =   26
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Aportes"
         Index           =   27
         Begin VB.Menu mnuFATConAporte 
            Caption         =   "Pagamentos - Todos"
            Index           =   1
         End
         Begin VB.Menu mnuFATConAporte 
            Caption         =   "Pagamentos - Não Faturados"
            Index           =   2
         End
         Begin VB.Menu mnuFATConAporte 
            Caption         =   "Pagamentos - Sobre Fatura"
            Index           =   3
         End
         Begin VB.Menu mnuFATConAporte 
            Caption         =   "Pagamentos - Créditos a Faturar"
            Index           =   4
         End
         Begin VB.Menu mnuFATConAporte 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuFATConAporte 
            Caption         =   "Histórico de Utilização -Desconto em Fatura"
            Index           =   6
         End
         Begin VB.Menu mnuFATConAporte 
            Caption         =   "Histórico de Utilização - Créditos Efetuados"
            Index           =   7
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Vouchers"
         Index           =   28
         Begin VB.Menu mnuFATConVou 
            Caption         =   "Tela"
            Index           =   1
         End
         Begin VB.Menu mnuFATConVou 
            Caption         =   "Todos"
            Index           =   2
         End
         Begin VB.Menu mnuFATConVou 
            Caption         =   "Não faturados"
            Index           =   3
         End
         Begin VB.Menu mnuFATConVou 
            Caption         =   "Cancelados com Previsão de Reembolso"
            Index           =   4
         End
         Begin VB.Menu mnuFATConVou 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuFATConVou 
            Caption         =   "Passageiros"
            Index           =   6
         End
         Begin VB.Menu mnuFATConVou 
            Caption         =   "Pagtos com Cartão sem Autorização"
            Index           =   7
         End
         Begin VB.Menu mnuFATConVou 
            Caption         =   "Posição Geral"
            Index           =   8
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Comissões"
         Index           =   29
         Begin VB.Menu mnuFATConComis 
            Caption         =   "CMR - Todas"
            Index           =   1
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "CMR - Bloqueadas"
            Index           =   2
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "CMR - Liberadas"
            Index           =   3
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "CMC - Todas"
            Index           =   5
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "CMC - Bloqueadas"
            Index           =   6
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "CMC - Liberadas"
            Index           =   7
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "CMCC - Todas"
            Index           =   9
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "CMCC - Bloqueadas"
            Index           =   10
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "CMCC - Liberadas"
            Index           =   11
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "-"
            Index           =   12
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "CME - Todas"
            Index           =   13
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "CME - Bloqueadas"
            Index           =   14
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "CME - Liberadas"
            Index           =   15
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "-"
            Index           =   16
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "CMA"
            Index           =   17
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "-"
            Index           =   18
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "Comissões aguardando envio de NF"
            Index           =   19
         End
         Begin VB.Menu mnuFATConComis 
            Caption         =   "Todas"
            Index           =   20
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Ocorrências"
         Index           =   30
         Begin VB.Menu mnuFATConOcr 
            Caption         =   "Todas"
            Index           =   1
         End
         Begin VB.Menu mnuFATConOcr 
            Caption         =   "Bloqueadas"
            Index           =   2
         End
         Begin VB.Menu mnuFATConOcr 
            Caption         =   "Liberadas"
            Index           =   3
         End
         Begin VB.Menu mnuFATConOcr 
            Caption         =   "Inativações"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Faturas"
         Index           =   31
         Begin VB.Menu mnuFATConFat 
            Caption         =   "Todas"
            Index           =   1
         End
         Begin VB.Menu mnuFATConFat 
            Caption         =   "Canceladas"
            Index           =   2
         End
         Begin VB.Menu mnuFATConFat 
            Caption         =   "À Receber sem NF vinculada"
            Index           =   3
         End
         Begin VB.Menu mnuFATConFat 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuFATConFat 
            Caption         =   "Itens a Serem Faturados"
            Index           =   5
         End
         Begin VB.Menu mnuFATConFat 
            Caption         =   "Itens Faturados"
            Index           =   6
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "-"
         Index           =   32
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Notas Fiscais Eletrônicas Federal"
         Index           =   33
         Begin VB.Menu mnuFATConNFeFed 
            Caption         =   "Lotes"
            Index           =   1
         End
         Begin VB.Menu mnuFATConNFeFed 
            Caption         =   "Log de Envio dos Lotes"
            Index           =   2
         End
         Begin VB.Menu mnuFATConNFeFed 
            Caption         =   "Retorno de Envio dos Lotes"
            Index           =   3
         End
         Begin VB.Menu mnuFATConNFeFed 
            Caption         =   "Status das Notas Fiscais"
            Index           =   4
         End
         Begin VB.Menu mnuFATConNFeFed 
            Caption         =   "Retorno da Consulta de Lotes"
            Index           =   5
         End
         Begin VB.Menu mnuFATConNFeFed 
            Caption         =   "Retorno do Cancelamento de Nota Fiscal"
            Index           =   6
         End
         Begin VB.Menu mnuFATConNFeFed 
            Caption         =   "Retorno de Inutilização de Faixas"
            Index           =   7
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "-"
         Index           =   34
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Notas Fiscais de Serviço Eletrônicas"
         Index           =   35
         Begin VB.Menu mnuFATConNFSE 
            Caption         =   "Lotes"
            Index           =   1
         End
         Begin VB.Menu mnuFATConNFSE 
            Caption         =   "Log de Envio dos Lotes"
            Index           =   2
         End
         Begin VB.Menu mnuFATConNFSE 
            Caption         =   "Retorno de Envio dos Lotes"
            Index           =   3
         End
         Begin VB.Menu mnuFATConNFSE 
            Caption         =   "Notas Fiscais Autorizadas"
            Index           =   4
         End
         Begin VB.Menu mnuFATConNFSE 
            Caption         =   "Retorno da Consulta de Lotes"
            Index           =   5
         End
         Begin VB.Menu mnuFATConNFSE 
            Caption         =   "Retorno do Cancelamento de Nota Fiscal"
            Index           =   6
         End
         Begin VB.Menu mnuFATConNFSE 
            Caption         =   "Retono da Consulta da Situaçao do Lote"
            Index           =   7
         End
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "-"
         Index           =   36
      End
      Begin VB.Menu mnuFATCon 
         Caption         =   "Beneficiamento"
         Index           =   37
         Begin VB.Menu mnuFATConBenef 
            Caption         =   "Saldos nas Remessas"
            Index           =   1
         End
         Begin VB.Menu mnuFATConBenef 
            Caption         =   "Devoluções"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   7
      Begin VB.Menu mnuCOMCon 
         Caption         =   "Cadastros"
         Index           =   1
         Begin VB.Menu mnuCOMConCad 
            Caption         =   "Produtos"
            Index           =   1
         End
         Begin VB.Menu mnuCOMConCad 
            Caption         =   "Fornecedores"
            Index           =   2
         End
         Begin VB.Menu mnuCOMConCad 
            Caption         =   "Produtos x Fornecedores"
            Index           =   3
         End
         Begin VB.Menu mnuCOMConCad 
            Caption         =   "Requisitantes"
            Index           =   4
         End
         Begin VB.Menu mnuCOMConCad 
            Caption         =   "Compradores"
            Index           =   5
         End
         Begin VB.Menu mnuCOMConCad 
            Caption         =   "Alçada"
            Index           =   6
         End
      End
      Begin VB.Menu mnuCOMCon 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCOMCon 
         Caption         =   "Requisições de Compra"
         Index           =   3
         Begin VB.Menu mnuCOMConReqCompra 
            Caption         =   "&Todas"
            Index           =   1
         End
         Begin VB.Menu mnuCOMConReqCompra 
            Caption         =   "&Enviadas"
            Index           =   2
         End
         Begin VB.Menu mnuCOMConReqCompra 
            Caption         =   "&Não Enviadas"
            Index           =   3
         End
         Begin VB.Menu mnuCOMConReqCompra 
            Caption         =   "Itens de Requisições de Compra"
            Index           =   4
         End
      End
      Begin VB.Menu mnuCOMCon 
         Caption         =   "Pedidos de Cotação"
         Index           =   4
         Begin VB.Menu mnuCOMConPedCotacao 
            Caption         =   "&Todas"
            Index           =   1
         End
         Begin VB.Menu mnuCOMConPedCotacao 
            Caption         =   "&Baixadas"
            Index           =   2
         End
         Begin VB.Menu mnuCOMConPedCotacao 
            Caption         =   "&Ativas"
            Index           =   3
         End
         Begin VB.Menu mnuCOMConPedCotacao 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuCOMConPedCotacao 
            Caption         =   "Cotações Pendentes"
            Index           =   5
         End
         Begin VB.Menu mnuCOMConPedCotacao 
            Caption         =   "Cotações Atualizadas"
            Index           =   6
         End
      End
      Begin VB.Menu mnuCOMCon 
         Caption         =   "Concorrências"
         Index           =   5
         Begin VB.Menu mnuCOMConcorrencias 
            Caption         =   "&Todas"
            Index           =   1
         End
         Begin VB.Menu mnuCOMConcorrencias 
            Caption         =   "&Baixadas"
            Index           =   2
         End
         Begin VB.Menu mnuCOMConcorrencias 
            Caption         =   "&Ativas"
            Index           =   3
         End
      End
      Begin VB.Menu mnuCOMCon 
         Caption         =   "Pedidos de Compra"
         Index           =   6
         Begin VB.Menu mnuCOMConPedCompra 
            Caption         =   "To&dos"
            Index           =   1
         End
         Begin VB.Menu mnuCOMConPedCompra 
            Caption         =   "En&viados"
            Index           =   2
         End
         Begin VB.Menu mnuCOMConPedCompra 
            Caption         =   "Nã&o Enviados"
            Index           =   3
         End
         Begin VB.Menu mnuCOMConPedCompra 
            Caption         =   "Abe&rtos"
            Index           =   4
         End
         Begin VB.Menu mnuCOMConPedCompra 
            Caption         =   "Itens de Pedidos de Compra"
            Index           =   5
         End
         Begin VB.Menu mnuCOMConPedCompra 
            Caption         =   "Última Compra por Produto"
            Index           =   6
         End
         Begin VB.Menu mnuCOMConPedCompra 
            Caption         =   "&Pedidos de Compra x Requisições"
            Index           =   7
         End
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   8
      Begin VB.Menu mnuFISCon 
         Caption         =   "Exceções ICMS"
         Index           =   1
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Exceções IPI"
         Index           =   2
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Naturezas de Operação"
         Index           =   3
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Produtos"
         Index           =   4
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Tipos de Tributação"
         Index           =   5
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Tipos de Registro p/ Apuração de ICMS"
         Index           =   6
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Tipos de Registro p/ Apuração de IPI"
         Index           =   7
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Livros Abertos"
         Index           =   9
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Livros Fechados"
         Index           =   10
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Apuração ICMS"
         Index           =   12
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Apuração IPI"
         Index           =   13
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Lançamentos para Apuração ICMS"
         Index           =   14
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Lançamentos para Apuração IPI"
         Index           =   15
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Registros de Entrada"
         Index           =   16
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Registros de Saída"
         Index           =   17
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "ICMS Crédito\Débito - Itens NF"
         Index           =   19
      End
      Begin VB.Menu mnuFISCon 
         Caption         =   "Tributação - Itens NF"
         Index           =   20
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   9
      Begin VB.Menu mnuLJCon 
         Caption         =   "Cadastros"
         Index           =   1
         Begin VB.Menu mnuLJConCad 
            Caption         =   "Clientes"
            Index           =   1
         End
         Begin VB.Menu mnuLJConCad 
            Caption         =   "Produtos"
            Index           =   2
         End
         Begin VB.Menu mnuLJConCad 
            Caption         =   "Operadores"
            Index           =   3
         End
         Begin VB.Menu mnuLJConCad 
            Caption         =   "Caixas"
            Index           =   4
         End
         Begin VB.Menu mnuLJConCad 
            Caption         =   "Vendedores"
            Index           =   5
         End
         Begin VB.Menu mnuLJConCad 
            Caption         =   "Emissores de Cupons Fiscais"
            Index           =   6
         End
         Begin VB.Menu mnuLJConCad 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuLJConCad 
            Caption         =   "Preços"
            Index           =   8
         End
         Begin VB.Menu mnuLJConCad 
            Caption         =   "Meios de Pagamento"
            Index           =   9
         End
         Begin VB.Menu mnuLJConCad 
            Caption         =   "Redes"
            Index           =   10
         End
      End
      Begin VB.Menu mnuLJCon 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuLJCon 
         Caption         =   "Movimentos de Caixa"
         Index           =   3
      End
      Begin VB.Menu mnuLJCon 
         Caption         =   "Movimentos de Caixa - Sessões"
         Index           =   4
      End
      Begin VB.Menu mnuLJCon 
         Caption         =   "Movimentos de Caixa - Itens de CF"
         Index           =   5
      End
      Begin VB.Menu mnuLJCon 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuLJCon 
         Caption         =   "Cupom Fiscal"
         Index           =   7
      End
      Begin VB.Menu mnuLJCon 
         Caption         =   "Itens de Cupom Fiscal"
         Index           =   8
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   10
      Begin VB.Menu mnuPCPCon 
         Caption         =   "Cadastros"
         Index           =   1
         Begin VB.Menu mnuPCPConCad 
            Caption         =   "Produtos"
            Index           =   1
         End
         Begin VB.Menu mnuPCPConCad 
            Caption         =   "Almoxarifados"
            Index           =   2
         End
         Begin VB.Menu mnuPCPConCad 
            Caption         =   "Kit"
            Index           =   3
         End
         Begin VB.Menu mnuPCPConCad 
            Caption         =   "Tipos de Produtos"
            Index           =   4
         End
         Begin VB.Menu mnuPCPConCad 
            Caption         =   "Classes de Unidade de Medida"
            Index           =   5
         End
         Begin VB.Menu mnuPCPConCad 
            Caption         =   "Competências"
            Index           =   6
         End
         Begin VB.Menu mnuPCPConCad 
            Caption         =   "Máquinas, Habilidades e Processos"
            Index           =   7
         End
         Begin VB.Menu mnuPCPConCad 
            Caption         =   "Centros de Trabalho"
            Index           =   8
         End
         Begin VB.Menu mnuPCPConCad 
            Caption         =   "Tipos de Mão de Obra"
            Index           =   9
         End
      End
      Begin VB.Menu mnuPCPCon 
         Caption         =   "Produção"
         Index           =   2
         Begin VB.Menu mnuPCPConPro 
            Caption         =   "Ordens de Produção"
            Index           =   1
         End
         Begin VB.Menu mnuPCPConPro 
            Caption         =   "Ordens de Produção Baixadas"
            Index           =   2
         End
         Begin VB.Menu mnuPCPConPro 
            Caption         =   "Itens Empenhados"
            Index           =   3
         End
         Begin VB.Menu mnuPCPConPro 
            Caption         =   "Itens em Produção"
            Index           =   4
         End
         Begin VB.Menu mnuPCPConPro 
            Caption         =   "Requisição para Produção"
            Index           =   5
         End
         Begin VB.Menu mnuPCPConPro 
            Caption         =   "Entrada de Produção"
            Index           =   6
         End
         Begin VB.Menu mnuPCPConPro 
            Caption         =   "Itens Requisitados p/ Produção"
            Index           =   7
         End
         Begin VB.Menu mnuPCPConPro 
            Caption         =   "Itens Produzidos"
            Index           =   8
         End
         Begin VB.Menu mnuPCPConPro 
            Caption         =   "Roteiros de Fabricação"
            Index           =   9
         End
         Begin VB.Menu mnuPCPConPro 
            Caption         =   "Taxas de Produção"
            Index           =   10
         End
         Begin VB.Menu mnuPCPConPro 
            Caption         =   "Plano Mestre de Produção"
            Index           =   11
         End
      End
      Begin VB.Menu mnuPCPCon 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPCPCon 
         Caption         =   "Controle de Estoque"
         Index           =   4
         Begin VB.Menu mnuPCPConEst 
            Caption         =   "Estoque da Empresa"
            Index           =   1
         End
         Begin VB.Menu mnuPCPConEst 
            Caption         =   "Estoque da Filial"
            Index           =   2
         End
         Begin VB.Menu mnuPCPConEst 
            Caption         =   "Estoque em Terceiros"
            Index           =   3
         End
      End
      Begin VB.Menu mnuPCPCon 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPCPCon 
         Caption         =   "Rastreamento"
         Index           =   6
         Begin VB.Menu mnuPCPConRastro 
            Caption         =   "Lotes"
            Index           =   1
         End
         Begin VB.Menu mnuPCPConRastro 
            Caption         =   "Saldos nos Lotes"
            Index           =   2
         End
         Begin VB.Menu mnuPCPConRastro 
            Caption         =   "Movimentos"
            Index           =   3
         End
      End
      Begin VB.Menu mnuPCPCon 
         Caption         =   "Certificados"
         Index           =   7
         Begin VB.Menu mnuPCPConCur 
            Caption         =   "Certificados"
            Index           =   1
         End
         Begin VB.Menu mnuPCPConCur 
            Caption         =   "Cursos/Exames"
            Index           =   2
         End
         Begin VB.Menu mnuPCPConCur 
            Caption         =   "Participantes x Cursos"
            Index           =   3
         End
         Begin VB.Menu mnuPCPConCur 
            Caption         =   "Participantes x Certificados"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   11
      Begin VB.Menu mnuCRMCon 
         Caption         =   "Cadastros"
         Index           =   1
         Begin VB.Menu mnuCRMConCad 
            Caption         =   "Clientes"
            Index           =   1
         End
         Begin VB.Menu mnuCRMConCad 
            Caption         =   "Atendentes"
            Index           =   2
         End
         Begin VB.Menu mnuCRMConCad 
            Caption         =   "Vendedores"
            Index           =   3
         End
         Begin VB.Menu mnuCRMConCad 
            Caption         =   "Clientes x Contatos"
            Index           =   4
         End
         Begin VB.Menu mnuCRMConCad 
            Caption         =   "Clientes x Telefones e Emails"
            Index           =   5
         End
      End
      Begin VB.Menu mnuCRMCon 
         Caption         =   "À Partir de Cliente"
         Index           =   2
      End
      Begin VB.Menu mnuCRMCon 
         Caption         =   "Relacionamentos"
         Index           =   3
         Begin VB.Menu mnuCRMConRelac 
            Caption         =   "Pendentes"
            Index           =   1
         End
         Begin VB.Menu mnuCRMConRelac 
            Caption         =   "Encerrados"
            Index           =   2
         End
         Begin VB.Menu mnuCRMConRelac 
            Caption         =   "Todos"
            Index           =   3
         End
         Begin VB.Menu mnuCRMConRelac 
            Caption         =   "Call Center"
            Index           =   4
         End
         Begin VB.Menu mnuCRMConRelac 
            Caption         =   "Solicitações de Srv"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   12
      Begin VB.Menu mnuQUACon 
         Caption         =   "Cadastros"
         Index           =   1
         Begin VB.Menu mnuQUAConCad 
            Caption         =   "Produtos"
            Index           =   1
         End
         Begin VB.Menu mnuQUAConCad 
            Caption         =   "Testes"
            Index           =   2
         End
         Begin VB.Menu mnuQUAConCad 
            Caption         =   "Produtos x Testes"
            Index           =   3
         End
      End
      Begin VB.Menu mnuQUACon 
         Caption         =   "Resultados dos Testes"
         Index           =   2
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   13
      Begin VB.Menu mnuPRJCon 
         Caption         =   "Cadastros"
         Index           =   1
         Begin VB.Menu mnuPRJConCad 
            Caption         =   "Produtos"
            Index           =   1
         End
         Begin VB.Menu mnuPRJConCad 
            Caption         =   "Clientes"
            Index           =   2
         End
         Begin VB.Menu mnuPRJConCad 
            Caption         =   "Fornecedores"
            Index           =   3
         End
      End
      Begin VB.Menu mnuPRJCon 
         Caption         =   "Projetos"
         Index           =   2
      End
      Begin VB.Menu mnuPRJCon 
         Caption         =   "Etapas"
         Index           =   3
      End
      Begin VB.Menu mnuPRJCon 
         Caption         =   "Propostas"
         Index           =   4
      End
      Begin VB.Menu mnuPRJCon 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPRJCon 
         Caption         =   "Pagamentos"
         Index           =   6
      End
      Begin VB.Menu mnuPRJCon 
         Caption         =   "Recebimentos"
         Index           =   7
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Index           =   14
      Begin VB.Menu mnuSRVCon 
         Caption         =   "&Cadastros"
         Index           =   1
         Begin VB.Menu mnuSRVConCad 
            Caption         =   "&Produtos"
            Index           =   1
         End
         Begin VB.Menu mnuSRVConCad 
            Caption         =   "&Garantias"
            Index           =   2
         End
         Begin VB.Menu mnuSRVConCad 
            Caption         =   "&Tipos de Garantia"
            Index           =   3
         End
         Begin VB.Menu mnuSRVConCad 
            Caption         =   "&Contratos de Manutenção"
            Index           =   4
         End
         Begin VB.Menu mnuSRVConCad 
            Caption         =   "&Contratos de Manutenção - Itens"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSRVConCad 
            Caption         =   "&Mão de Obra"
            Index           =   6
         End
         Begin VB.Menu mnuSRVConCad 
            Caption         =   "&Máquinas"
            Index           =   7
         End
         Begin VB.Menu mnuSRVConCad 
            Caption         =   "&Competências"
            Index           =   8
         End
         Begin VB.Menu mnuSRVConCad 
            Caption         =   "&Centros de Trabalho"
            Index           =   9
         End
      End
      Begin VB.Menu mnuSRVCon 
         Caption         =   "&Solicitações"
         Index           =   2
         Begin VB.Menu mnuSRVConSolic 
            Caption         =   "&Abertas"
            Index           =   1
         End
         Begin VB.Menu mnuSRVConSolic 
            Caption         =   "&Baixadas"
            Index           =   2
         End
         Begin VB.Menu mnuSRVConSolic 
            Caption         =   "&Todas"
            Index           =   3
         End
         Begin VB.Menu mnuSRVConSolic 
            Caption         =   "&CRM"
            Index           =   4
         End
         Begin VB.Menu mnuSRVConSolic 
            Caption         =   "&Itens"
            Index           =   5
         End
      End
      Begin VB.Menu mnuSRVCon 
         Caption         =   "&Orçamentos"
         Index           =   3
      End
      Begin VB.Menu mnuSRVCon 
         Caption         =   "&Pedidos"
         Index           =   4
         Begin VB.Menu mnuSRVConPed 
            Caption         =   "&Abertos"
            Index           =   1
         End
         Begin VB.Menu mnuSRVConPed 
            Caption         =   "&Baixados"
            Index           =   2
         End
         Begin VB.Menu mnuSRVConPed 
            Caption         =   "&Todos"
            Index           =   3
         End
         Begin VB.Menu mnuSRVConPed 
            Caption         =   "&Itens"
            Index           =   4
         End
      End
      Begin VB.Menu mnuSRVCon 
         Caption         =   "&Itens de Pedido"
         Index           =   5
      End
      Begin VB.Menu mnuSRVCon 
         Caption         =   "&Ordens de Serviço"
         Index           =   6
         Begin VB.Menu mnuSRVConOS 
            Caption         =   "&Abertas"
            Index           =   1
         End
         Begin VB.Menu mnuSRVConOS 
            Caption         =   "&Baixadas"
            Index           =   2
         End
         Begin VB.Menu mnuSRVConOS 
            Caption         =   "&Todas"
            Index           =   3
         End
      End
      Begin VB.Menu mnuSRVCon 
         Caption         =   "&Itens Ordens de Serviço"
         Index           =   7
      End
      Begin VB.Menu mnuSRVCon 
         Caption         =   "&Notas Fiscais de Serviço Simples"
         Index           =   8
      End
      Begin VB.Menu mnuSRVCon 
         Caption         =   "&Itens Notas Fiscais de Serviço"
         Index           =   9
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   1
      Begin VB.Menu mnuCTBRel 
         Caption         =   "&Cadastros"
         Index           =   1
         Begin VB.Menu mnuCTBRelCad 
            Caption         =   "Plano de Contas"
            Index           =   1
         End
         Begin VB.Menu mnuCTBRelCad 
            Caption         =   "Centro de Custos/Lucro"
            Index           =   2
         End
         Begin VB.Menu mnuCTBRelCad 
            Caption         =   "Histórico Padrão"
            Index           =   3
         End
         Begin VB.Menu mnuCTBRelCad 
            Caption         =   "Lotes Contabilizados"
            Index           =   4
         End
         Begin VB.Menu mnuCTBRelCad 
            Caption         =   "Lotes Pendentes"
            Index           =   5
         End
         Begin VB.Menu mnuCTBRelCad 
            Caption         =   "Lançamentos por Data"
            Index           =   6
         End
         Begin VB.Menu mnuCTBRelCad 
            Caption         =   "Lançamentos por Centro de Custo/Lucro"
            Index           =   7
         End
         Begin VB.Menu mnuCTBRelCad 
            Caption         =   "Lançamentos por Lote"
            Index           =   8
         End
         Begin VB.Menu mnuCTBRelCad 
            Caption         =   "Lançamentos Pendentes"
            Index           =   9
         End
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "&Balancete de Verificação"
         Index           =   3
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "&Razão"
         Index           =   4
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "Razão &Auxiliar"
         Index           =   5
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "Razão Aglutinado"
         Index           =   6
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "&Diário"
         Index           =   7
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "Diário Auxiliar"
         Index           =   8
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "Diário Aglutinado"
         Index           =   9
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "D&emostrativos"
         Index           =   10
         Begin VB.Menu mnuCTBRelDRE 
            Caption         =   "Resultado do Exercício (DRE)"
         End
         Begin VB.Menu mnuCTBRelDRP 
            Caption         =   "Resultado do Período (DRP)"
         End
         Begin VB.Menu mnuCTBRelDemMutPatrLiq 
            Caption         =   "Mutações do Patrimônio Líquido (DMPL)"
         End
         Begin VB.Menu mnuCTBRelDemOAR 
            Caption         =   "Origens e Aplicações de Recursos (DOAR)"
         End
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "&Balanço Patrimonial"
         Index           =   11
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "Outros"
         Index           =   13
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "&Gerador de Relatórios"
         Index           =   15
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuCTBRel 
         Caption         =   "Planilhas"
         Index           =   17
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   2
      Begin VB.Menu mnuTESRel 
         Caption         =   "Extrato de &Tesouraria"
         Index           =   1
      End
      Begin VB.Menu mnuTESRel 
         Caption         =   "Extrato &Bancário"
         Index           =   2
      End
      Begin VB.Menu mnuTESRel 
         Caption         =   "&Posição de Aplicações"
         Index           =   3
      End
      Begin VB.Menu mnuTESRel 
         Caption         =   "B&orderô de Pré-Datados"
         Index           =   4
      End
      Begin VB.Menu mnuTESRel 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuTESRel 
         Caption         =   "Movimento Financeiro &Reduzido"
         Index           =   7
      End
      Begin VB.Menu mnuTESRel 
         Caption         =   "Movimento Financeiro &Detalhado"
         Index           =   8
      End
      Begin VB.Menu mnuTESRel 
         Caption         =   "&Erros na Conciliação Automática"
         Index           =   9
      End
      Begin VB.Menu mnuTESRel 
         Caption         =   "&Conciliações Pendentes"
         Index           =   10
      End
      Begin VB.Menu mnuTESRel 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuTESRel 
         Caption         =   "&Outros"
         Index           =   12
      End
      Begin VB.Menu mnuTESRel 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuTESRel 
         Caption         =   "&Gerador de Relatórios"
         Index           =   14
      End
      Begin VB.Menu mnuTESRel 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuTESRel 
         Caption         =   "Planilhas"
         Index           =   16
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   3
      Begin VB.Menu mnuCPRel 
         Caption         =   "&Cadastros"
         Index           =   1
         Begin VB.Menu mnuCPRelCad 
            Caption         =   "&Fornecedores"
            Index           =   1
         End
         Begin VB.Menu mnuCPRelCad 
            Caption         =   "Tipos de Fornecedor"
            Index           =   2
         End
         Begin VB.Menu mnuCPRelCad 
            Caption         =   "Portadores"
            Index           =   3
         End
         Begin VB.Menu mnuCPRelCad 
            Caption         =   "Condições de Pagamento"
            Index           =   4
         End
      End
      Begin VB.Menu mnuCPRel 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCPRel 
         Caption         =   "&Títulos a Pagar"
         Index           =   3
      End
      Begin VB.Menu mnuCPRel 
         Caption         =   "Posição dos Fornecedores"
         Index           =   4
      End
      Begin VB.Menu mnuCPRel 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuCPRel 
         Caption         =   "Cheques Emitidos"
         Index           =   6
      End
      Begin VB.Menu mnuCPRel 
         Caption         =   "Pagamentos / Baixas"
         Index           =   7
      End
      Begin VB.Menu mnuCPRel 
         Caption         =   "Pagamentos Cancelados"
         Index           =   8
      End
      Begin VB.Menu mnuCPRel 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuCPRel 
         Caption         =   "&Outros"
         Index           =   10
      End
      Begin VB.Menu mnuCPRel 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuCPRel 
         Caption         =   "&Gerador de Relatórios"
         Index           =   12
      End
      Begin VB.Menu mnuCPRel 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuCPRel 
         Caption         =   "Planilhas"
         Index           =   14
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   4
      Begin VB.Menu mnuCRRel 
         Caption         =   "&Cadastros"
         Index           =   1
         Begin VB.Menu mnuCRRelCad 
            Caption         =   "Clientes"
            Index           =   1
         End
         Begin VB.Menu mnuCRRelCad 
            Caption         =   "Tipos de Cliente"
            Index           =   2
         End
         Begin VB.Menu mnuCRRelCad 
            Caption         =   "Categorias de Cliente"
            Index           =   3
         End
         Begin VB.Menu mnuCRRelCad 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuCRRelCad 
            Caption         =   "Cobradores"
            Index           =   5
         End
         Begin VB.Menu mnuCRRelCad 
            Caption         =   "Tipos de Carteira de Cobrança"
            Index           =   6
         End
         Begin VB.Menu mnuCRRelCad 
            Caption         =   "Padrões de Cobrança"
            Index           =   7
         End
         Begin VB.Menu mnuCRRelCad 
            Caption         =   "Tipos de Instrução de Cobrança"
            Index           =   8
         End
         Begin VB.Menu mnuCRRelCad 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnuCRRelCad 
            Caption         =   "Vendedores"
            Index           =   10
         End
         Begin VB.Menu mnuCRRelCad 
            Caption         =   "Tipos de Vendedor"
            Index           =   11
         End
         Begin VB.Menu mnuCRRelCad 
            Caption         =   "Regiões de Venda"
            Index           =   12
         End
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "&Títulos a Receber"
         Index           =   3
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "Títulos em &Atraso"
         Index           =   4
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "Posição Geral da Cobrança"
         Index           =   5
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "P&osição do Cliente"
         Index           =   6
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "Recebimentos / Baixas"
         Index           =   8
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "Comissões"
         Index           =   10
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "Resumo de Comissões"
         Index           =   11
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "&Maiores Devedores"
         Index           =   13
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "Cobrança por Telefone"
         Index           =   14
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "Cobrança Via Mala Direta"
         Index           =   15
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "&Outros"
         Index           =   17
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "&Gerador de Relatórios"
         Index           =   19
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu mnuCRRel 
         Caption         =   "Planilhas"
         Index           =   21
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   5
      Begin VB.Menu mnuESTRel 
         Caption         =   "&Cadastros"
         Index           =   1
         Begin VB.Menu mnuESTRelCad 
            Caption         =   "Relação de Fornecedores"
            Index           =   1
         End
         Begin VB.Menu mnuESTRelCad 
            Caption         =   "Relação de Produtos"
            Index           =   2
         End
         Begin VB.Menu mnuESTRelCad 
            Caption         =   "Relação de Almoxarifados"
            Index           =   3
         End
         Begin VB.Menu mnuESTRelCad 
            Caption         =   "Relação de &Kits"
            Index           =   4
         End
         Begin VB.Menu mnuESTRelCad 
            Caption         =   "Utilização do Produto"
            Index           =   5
         End
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "Movimentos"
         Index           =   3
         Begin VB.Menu mnuESTRelMov 
            Caption         =   "Relação das Movimentações &Internas"
            Index           =   1
         End
         Begin VB.Menu mnuESTRelMov 
            Caption         =   "Requisições para Consumo"
            Index           =   2
         End
         Begin VB.Menu mnuESTRelMov 
            Caption         =   "Lista dos &Movimentos - Kardex"
            Index           =   3
         End
         Begin VB.Menu mnuESTRelMov 
            Caption         =   "Lista dos Movimentos &Diários - Kardex p/ dia"
            Index           =   4
         End
         Begin VB.Menu mnuESTRelMov 
            Caption         =   "Lista dos Movimentos &Resumidos por dia"
            Index           =   5
         End
         Begin VB.Menu mnuESTRelMov 
            Caption         =   "&Boletim de Entrada"
            Index           =   6
         End
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "Inventário"
         Index           =   4
         Begin VB.Menu mnuESTRelInv 
            Caption         =   "&Etiquetas para Inventário"
            Index           =   1
         End
         Begin VB.Menu mnuESTRelInv 
            Caption         =   "Listagem para Inventário"
            Index           =   2
         End
         Begin VB.Menu mnuESTRelInv 
            Caption         =   "Demonstrativo de Apuração de Inventário"
            Index           =   3
         End
         Begin VB.Menu mnuESTRelInv 
            Caption         =   "Registro de Inventário"
            Index           =   4
         End
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "Produção"
         Index           =   5
         Begin VB.Menu mnuESTRelPro 
            Caption         =   "&Ordens de Produção"
            Index           =   1
         End
         Begin VB.Menu mnuESTRelPro 
            Caption         =   "L&ista dos Empenhos"
            Index           =   2
         End
         Begin VB.Menu mnuESTRelPro 
            Caption         =   "Lista de &Faltas"
            Index           =   3
         End
         Begin VB.Menu mnuESTRelPro 
            Caption         =   "Movimentos de Estoque para cada Ordem de Produção"
            Index           =   4
         End
         Begin VB.Menu mnuESTRelPro 
            Caption         =   "Resumo dos Produtos - Ordem de Produção"
            Index           =   5
         End
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "Relação dos Produtos Vendidos"
         Index           =   7
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "Resumo das Entradas e Saídas em Valor"
         Index           =   8
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "Consumo / Vendas Mês a Mês"
         Index           =   9
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "&Análise do Estoque"
         Index           =   11
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "A&nálise de Movimentação de Estoque"
         Index           =   12
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "Produtos que Atingiram o Ponto de Pedido"
         Index           =   13
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "Saldo em Estoque"
         Index           =   14
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "&Outros"
         Index           =   16
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "&Gerador de Relatórios"
         Index           =   18
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu mnuESTRel 
         Caption         =   "Planilhas"
         Index           =   20
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   6
      Begin VB.Menu mnuFATRel 
         Caption         =   "Cadastro"
         Index           =   1
         Begin VB.Menu mnuFATRelCad 
            Caption         =   "Relação de Clientes"
            Index           =   1
         End
         Begin VB.Menu mnuFATRelCad 
            Caption         =   "Relação de Produtos"
            Index           =   2
         End
         Begin VB.Menu mnuFATRelCad 
            Caption         =   "Relação de Vendedores"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "Documentos"
         Index           =   3
         Begin VB.Menu mnuFATRelDoc 
            Caption         =   "Pré-Nota"
            Index           =   1
         End
         Begin VB.Menu mnuFATRelDoc 
            Caption         =   "Relação de Notas Fiscais"
            Index           =   2
         End
         Begin VB.Menu mnuFATRelDoc 
            Caption         =   "Notas Fiscais de Devolução"
            Index           =   3
         End
         Begin VB.Menu mnuFATRelDoc 
            Caption         =   "Notas Fiscais para Transportadoras"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "Vendas"
         Index           =   4
         Begin VB.Menu mnuFATRelVen 
            Caption         =   "Faturamento por &Cliente"
            Index           =   1
         End
         Begin VB.Menu mnuFATRelVen 
            Caption         =   "Faturamento Cliente x Produto"
            Index           =   2
         End
         Begin VB.Menu mnuFATRelVen 
            Caption         =   "Faturamento por &Vendedor"
            Index           =   3
         End
         Begin VB.Menu mnuFATRelVen 
            Caption         =   "Faturamento por Prazo de Pagamento"
            Index           =   4
         End
         Begin VB.Menu mnuFATRelVen 
            Caption         =   "Faturamento Real x Previsto"
            Index           =   5
         End
         Begin VB.Menu mnuFATRelVen 
            Caption         =   "Resumo de Vendas"
            Index           =   6
         End
         Begin VB.Menu mnuFATRelVen 
            Caption         =   "Disponibilidade de Estoque para Venda"
            Index           =   7
         End
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "Pedidos"
         Index           =   5
         Begin VB.Menu mnuFATRelPed 
            Caption         =   "Pedidos Aptos a Faturar"
            Index           =   1
         End
         Begin VB.Menu mnuFATRelPed 
            Caption         =   "Pedidos não Entregues"
            Index           =   2
         End
         Begin VB.Menu mnuFATRelPed 
            Caption         =   "Pedidos por Produtos"
            Index           =   3
         End
         Begin VB.Menu mnuFATRelPed 
            Caption         =   "Pedidos de Vendas por Vendedor / Cliente"
            Index           =   4
         End
         Begin VB.Menu mnuFATRelPed 
            Caption         =   "Pedidos por Vendedor / Produto"
            Index           =   5
         End
         Begin VB.Menu mnuFATRelPed 
            Caption         =   "Pedidos Para Produção"
            Index           =   6
         End
         Begin VB.Menu mnuFATRelPed 
            Caption         =   "Pedidos de Vendas por Cliente"
            Index           =   7
         End
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "Comi&ssões"
         Index           =   7
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "Lista de Preços"
         Index           =   8
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "&Outros"
         Index           =   10
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "&Gerador de Relatórios"
         Index           =   12
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "Planilhas"
         Index           =   14
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "-"
         Index           =   30
      End
      Begin VB.Menu mnuFATRel 
         Caption         =   "Travel Ace"
         Index           =   35
         Begin VB.Menu mnuFATRelTRV 
            Caption         =   "Atendentes"
            Index           =   1
         End
         Begin VB.Menu mnuFATRelTRV 
            Caption         =   "Estatísticas"
            Index           =   2
         End
         Begin VB.Menu mnuFATRelTRV 
            Caption         =   "Desvio no padrão de vendas"
            Index           =   3
         End
         Begin VB.Menu mnuFATRelTRV 
            Caption         =   "Vouchers Emitidos"
            Index           =   4
         End
         Begin VB.Menu mnuFATRelTRV 
            Caption         =   "Acompanhamento de inadimplência"
            Index           =   5
         End
         Begin VB.Menu mnuFATRelTRV 
            Caption         =   "Posição de inadimplência"
            Index           =   6
         End
         Begin VB.Menu mnuFATRelTRV 
            Caption         =   "Log de Importação"
            Index           =   7
         End
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   7
      Begin VB.Menu mnuCOMRel 
         Caption         =   "Cadastro"
         Index           =   1
         Begin VB.Menu mnuCOMRelCad 
            Caption         =   "Requisitantes"
            Index           =   1
         End
         Begin VB.Menu mnuCOMRelCad 
            Caption         =   "Compradores"
            Index           =   2
         End
         Begin VB.Menu mnuCOMRelCad 
            Caption         =   "Fornecedores"
            Index           =   3
         End
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "Pedidos de Compra"
         Index           =   2
         Begin VB.Menu mnuCOMRelPC 
            Caption         =   "Pedidos de Compra em Aberto"
            Index           =   1
         End
         Begin VB.Menu mnuCOMRelPC 
            Caption         =   "Pedidos de Compra Baixados"
            Index           =   2
         End
         Begin VB.Menu mnuCOMRelPC 
            Caption         =   "Pedidos de Compra Atrasados"
            Index           =   3
         End
         Begin VB.Menu mnuCOMRelPC 
            Caption         =   "Pedidos de Compra Bloqueados"
            Index           =   4
         End
         Begin VB.Menu mnuCOMRelPC 
            Caption         =   "Pedidos de Compra Emitidos"
            Index           =   5
         End
         Begin VB.Menu mnuCOMRelPC 
            Caption         =   "Pedidos de Compra Emitidos por Concorrência"
            Index           =   6
         End
         Begin VB.Menu mnuCOMRelPC 
            Caption         =   "Pedidos de Compra x Notas Fiscais"
            Index           =   7
         End
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "Pedidos de Cotação"
         Index           =   3
         Begin VB.Menu mnuCOMRelPCot 
            Caption         =   "Pedidos de Cotação Emitidos"
            Index           =   1
         End
         Begin VB.Menu mnuCOMRelPCot 
            Caption         =   "Pedidos de Cotação Emitidos por Geração"
            Index           =   2
         End
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "Requisições de Compra"
         Index           =   4
         Begin VB.Menu mnuCOMRelRC 
            Caption         =   "Requisições de Compra em Aberto"
            Index           =   1
         End
         Begin VB.Menu mnuCOMRelRC 
            Caption         =   "Requisições de Compra Baixadas"
            Index           =   2
         End
         Begin VB.Menu mnuCOMRelRC 
            Caption         =   "Requisições de Compra Atrasadas"
            Index           =   3
         End
         Begin VB.Menu mnuCOMRelRC 
            Caption         =   "Requisições de Compra x Pedidos de Venda"
            Index           =   4
         End
         Begin VB.Menu mnuCOMRelRC 
            Caption         =   "Requisições de Compra x Ordens de Produção"
            Index           =   5
         End
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "Produtos"
         Index           =   5
         Begin VB.Menu mnuCOMRelProd 
            Caption         =   "Produtos x Cotações"
            Index           =   1
         End
         Begin VB.Menu mnuCOMRelProd 
            Caption         =   "Produtos x Fornecedores"
            Index           =   2
         End
         Begin VB.Menu mnuCOMRelProd 
            Caption         =   "Produtos x Compras"
            Index           =   3
         End
         Begin VB.Menu mnuCOMRelProd 
            Caption         =   "Produtos x Requisições"
            Index           =   4
         End
         Begin VB.Menu mnuCOMRelProd 
            Caption         =   "Produtos x Pedidos de Compra"
            Index           =   5
         End
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "Concorrências Abertas"
         Index           =   7
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "Análise de Cotações Recebidas"
         Index           =   8
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "Previsão de Entrega de Produtos"
         Index           =   10
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "Agenda de Previsão de Entrega"
         Index           =   11
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "&Outros"
         Index           =   13
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "&Gerador de Relatórios"
         Index           =   15
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuCOMRel 
         Caption         =   "Planilhas"
         Index           =   17
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   8
      Begin VB.Menu mnuFISRel 
         Caption         =   "Termo de Abertura e Fechamento"
         Index           =   1
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Livros Fiscais ISS"
         Index           =   2
         Begin VB.Menu mnuFISRelLivroISS 
            Caption         =   "Registro de Apuracao do ISS (mod 3)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Livros Fiscais ICMS / IPI"
         Index           =   4
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Registro de Entradas (mod 1 e 1a)"
            Index           =   1
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Registro de Saídas (mod 2 e 2a)"
            Index           =   2
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Registro de Controle de Produção e Estoque (mod 3)"
            Index           =   3
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Registro de Inventário (mod 7)"
            Index           =   4
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Registro de Apuração do ICMS (mod 9)"
            Index           =   5
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Resumo de Apuração ICMS (mod 9)"
            Index           =   6
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Registro de Apuração do IPI (mod 8)"
            Index           =   7
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Resumo de Apuração IPI (mod 8)"
            Index           =   8
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Lista de Códigos de Emitentes (mod 10)"
            Index           =   9
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Tabela de Códigos de Mercadorias (mod 11)"
            Index           =   10
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Listagem de Operações por UF (mod 12)"
            Index           =   11
         End
         Begin VB.Menu mnuFISRelLivICMSIPI 
            Caption         =   "Dados de Recolhimentos - GNR (mod 14)"
            Index           =   12
         End
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Gerenciais ICMS / IPI"
         Index           =   5
         Begin VB.Menu mnuFISRelGerenciais 
            Caption         =   "Registro de E/S por Natureza de Operação"
            Index           =   1
         End
         Begin VB.Menu mnuFISRelGerenciais 
            Caption         =   "Registro de E/S por Estado"
            Index           =   2
         End
         Begin VB.Menu mnuFISRelGerenciais 
            Caption         =   "Registro de E/S por Clientes"
            Index           =   4
         End
         Begin VB.Menu mnuFISRelGerenciais 
            Caption         =   "Registro de E/S por Fornecedor"
            Index           =   5
         End
         Begin VB.Menu mnuFISRelGerenciais 
            Caption         =   "Resumo do Registro de E/S "
            Index           =   6
         End
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Apuração de PIS"
         Index           =   7
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Apuração do COFINS"
         Index           =   8
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Retenção de IR"
         Index           =   9
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "INSS"
         Index           =   10
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Listagem IN 068/95"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Resumo para o preenchimento do Declan"
         Index           =   12
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Retenção de PIS, COFINS, CSLL"
         Index           =   13
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Outros"
         Index           =   15
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Gerador de Relatórios"
         Index           =   17
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu mnuFISRel 
         Caption         =   "Planilhas"
         Index           =   19
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   9
      Begin VB.Menu mnuLJRel 
         Caption         =   "Cadastros"
         Index           =   1
         Begin VB.Menu mnuLJRelCad 
            Caption         =   "Relação de Clientes"
            Index           =   1
         End
         Begin VB.Menu mnuLJRelCad 
            Caption         =   "Relação de Produtos"
            Index           =   2
         End
         Begin VB.Menu mnuLJRelCad 
            Caption         =   "Relação de Operadores"
            Index           =   3
         End
         Begin VB.Menu mnuLJRelCad 
            Caption         =   "Relação de Caixas"
            Index           =   4
         End
         Begin VB.Menu mnuLJRelCad 
            Caption         =   "Relação de ECF's"
            Index           =   5
         End
      End
      Begin VB.Menu mnuLJRel 
         Caption         =   "Caixa"
         Index           =   2
         Begin VB.Menu mnuLJRelCx 
            Caption         =   "Movimentos de Caixa"
            Index           =   1
         End
         Begin VB.Menu mnuLJRelCx 
            Caption         =   "Painel de Caixas"
            Index           =   2
         End
         Begin VB.Menu mnuLJRelCx 
            Caption         =   "Orçamentos Emitidos"
            Index           =   3
         End
         Begin VB.Menu mnuLJRelCx 
            Caption         =   "Cupons Emitidos"
            Index           =   4
         End
      End
      Begin VB.Menu mnuLJRel 
         Caption         =   "Vendas"
         Index           =   3
         Begin VB.Menu mnuLJRelVen 
            Caption         =   "Evolução de Vendas"
            Index           =   1
         End
         Begin VB.Menu mnuLJRelVen 
            Caption         =   "Mapa de Vendas"
            Index           =   2
         End
         Begin VB.Menu mnuLJRelVen 
            Caption         =   "Flash de Vendas"
            Index           =   3
         End
         Begin VB.Menu mnuLJRelVen 
            Caption         =   "Vendas x Meios de Pagamento"
            Index           =   4
         End
         Begin VB.Menu mnuLJRelVen 
            Caption         =   "Ranking de Produtos Vendidos"
            Index           =   5
         End
         Begin VB.Menu mnuLJRelVen 
            Caption         =   "Produtos Devolvidos em Trocas"
            Index           =   6
         End
      End
      Begin VB.Menu mnuLJRel 
         Caption         =   "Borderôs"
         Index           =   4
         Begin VB.Menu mnuLJRelBord 
            Caption         =   "Borderôs de Boletos"
            Index           =   1
         End
         Begin VB.Menu mnuLJRelBord 
            Caption         =   "Borderôs de Ticket's"
            Index           =   2
         End
         Begin VB.Menu mnuLJRelBord 
            Caption         =   "Borderôs de Outros"
            Index           =   3
         End
      End
      Begin VB.Menu mnuLJRel 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuLJRel 
         Caption         =   "Outros"
         Index           =   6
      End
      Begin VB.Menu mnuLJRel 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuLJRel 
         Caption         =   "Gerador de Relatórios"
         Index           =   8
      End
      Begin VB.Menu mnuLJRel 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuLJRel 
         Caption         =   "Planilhas"
         Index           =   10
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   10
      Begin VB.Menu mnuPCPRel 
         Caption         =   "Cadastro"
         Index           =   1
         Begin VB.Menu mnuPCPRelCad 
            Caption         =   "Relação de Produtos"
            Index           =   1
         End
         Begin VB.Menu mnuPCPRelCad 
            Caption         =   "Relação de Almoxarifados"
            Index           =   2
         End
         Begin VB.Menu mnuPCPRelCad 
            Caption         =   "Relação de Kits"
            Index           =   3
         End
         Begin VB.Menu mnuPCPRelCad 
            Caption         =   "Utilização do Produto"
            Index           =   4
         End
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "Ordens de Produção"
         Index           =   3
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "Lista de Empenhos"
         Index           =   4
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "Lista de Faltas"
         Index           =   5
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "Distribuição de Matéria-Prima por Máquina"
         Index           =   7
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "Produtos x Ordens de Produção"
         Index           =   8
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "Movimentos de Estoque para cada Ordem de Produção"
         Index           =   9
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "Ordem de Produção x Requisição p/ Produção"
         Index           =   10
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "Análise de Rendimento por Ordem de Produção"
         Index           =   12
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "Previsão de Vendas x Previsão de Consumo"
         Index           =   13
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "Fórmula Padrão para Custo"
         Index           =   14
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "Outros"
         Index           =   16
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "Gerador de Relatórios"
         Index           =   18
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu mnuPCPRel 
         Caption         =   "Planilhas"
         Index           =   20
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   11
      Begin VB.Menu mnuCRMRel 
         Caption         =   "Cadastros"
         Index           =   1
         Begin VB.Menu mnuCRMRelCad 
            Caption         =   "Clientes"
            Index           =   1
         End
         Begin VB.Menu mnuCRMRelCad 
            Caption         =   "Atendentes"
            Index           =   2
         End
         Begin VB.Menu mnuCRMRelCad 
            Caption         =   "Vendedores"
            Index           =   3
         End
         Begin VB.Menu mnuCRMRelCad 
            Caption         =   "Clientes x Contatos"
            Index           =   4
         End
      End
      Begin VB.Menu mnuCRMRel 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCRMRel 
         Caption         =   "Follow-Up"
         Index           =   3
      End
      Begin VB.Menu mnuCRMRel 
         Caption         =   "Relacionamentos x Estatísticas"
         Index           =   4
      End
      Begin VB.Menu mnuCRMRel 
         Caption         =   "Clientes sem Relacionamentos"
         Index           =   5
      End
      Begin VB.Menu mnuCRMRel 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuCRMRel 
         Caption         =   "Outros"
         Index           =   7
      End
      Begin VB.Menu mnuCRMRel 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuCRMRel 
         Caption         =   "Gerador de Relatórios"
         Index           =   9
      End
      Begin VB.Menu mnuCRMRel 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuCRMRel 
         Caption         =   "Planilhas"
         Index           =   11
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   12
      Begin VB.Menu mnuQUARel 
         Caption         =   "Cadastros"
         Index           =   1
         Begin VB.Menu mnuQUARelCad 
            Caption         =   "Produtos"
            Index           =   1
         End
         Begin VB.Menu mnuQUARelCad 
            Caption         =   "Testes"
            Index           =   2
         End
         Begin VB.Menu mnuQUARelCad 
            Caption         =   "Produto x Testes"
            Index           =   3
         End
      End
      Begin VB.Menu mnuQUARel 
         Caption         =   "Ficha de Controle de Qualidade"
         Index           =   2
      End
      Begin VB.Menu mnuQUARel 
         Caption         =   "Laudos de Notas Fiscais"
         Index           =   3
      End
      Begin VB.Menu mnuQUARel 
         Caption         =   "Não Conformidade"
         Index           =   4
      End
      Begin VB.Menu mnuQUARel 
         Caption         =   "Resultados Pendentes"
         Index           =   5
      End
      Begin VB.Menu mnuQUARel 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuQUARel 
         Caption         =   "Outros"
         Index           =   7
      End
      Begin VB.Menu mnuQUARel 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuQUARel 
         Caption         =   "Gerador de Relatórios"
         Index           =   9
      End
      Begin VB.Menu mnuQUARel 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuQUARel 
         Caption         =   "Planilhas"
         Index           =   11
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   13
      Begin VB.Menu mnuPRJRel 
         Caption         =   "Previsto x Realizado"
         Index           =   1
      End
      Begin VB.Menu mnuPRJRel 
         Caption         =   "Fluxo Financeiro"
         Index           =   2
      End
      Begin VB.Menu mnuPRJRel 
         Caption         =   "Materiais Utilizados"
         Index           =   3
      End
      Begin VB.Menu mnuPRJRel 
         Caption         =   "Mãos-de-obra Utilizadas"
         Index           =   4
      End
      Begin VB.Menu mnuPRJRel 
         Caption         =   "Acompanhamento"
         Index           =   5
      End
      Begin VB.Menu mnuPRJRel 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuPRJRel 
         Caption         =   "Outros"
         Index           =   7
      End
      Begin VB.Menu mnuPRJRel 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuPRJRel 
         Caption         =   "Gerador de Relatórios"
         Index           =   9
      End
      Begin VB.Menu mnuPRJRel 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuPRJRel 
         Caption         =   "Planilhas"
         Index           =   11
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Index           =   14
      Begin VB.Menu mnuSRVRel 
         Caption         =   "Solicitações"
         Index           =   1
      End
      Begin VB.Menu mnuSRVRel 
         Caption         =   "Pedidos de Serviço"
         Index           =   2
         Begin VB.Menu mnuSRVRelPS 
            Caption         =   "Aptos a Faturar"
            Index           =   1
         End
         Begin VB.Menu mnuSRVRelPS 
            Caption         =   "Não Entregues"
            Index           =   2
         End
         Begin VB.Menu mnuSRVRelPS 
            Caption         =   "Por Serviço"
            Index           =   3
         End
         Begin VB.Menu mnuSRVRelPS 
            Caption         =   "Por Cliente"
            Index           =   4
         End
      End
      Begin VB.Menu mnuSRVRel 
         Caption         =   "Ordens de Serviço"
         Index           =   3
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   1
      Begin VB.Menu mnuCTBCad 
         Caption         =   "&Plano de Contas"
         Index           =   1
      End
      Begin VB.Menu mnuCTBCad 
         Caption         =   "&Categoria"
         Index           =   2
      End
      Begin VB.Menu mnuCTBCad 
         Caption         =   "C&entro de Custo/Lucro"
         Index           =   3
      End
      Begin VB.Menu mnuCTBCad 
         Caption         =   "Conta x Centro de Custo/Lucro Extra-Contábil"
         Index           =   4
      End
      Begin VB.Menu mnuCTBCad 
         Caption         =   "Conta x Centro Custo/Lucro Contábil"
         Index           =   5
      End
      Begin VB.Menu mnuCTBCad 
         Caption         =   "&Histórico Padrão"
         Index           =   6
      End
      Begin VB.Menu mnuCTBCad 
         Caption         =   "&Lotes"
         Index           =   7
      End
      Begin VB.Menu mnuCTBCad 
         Caption         =   "&Documento Automático"
         Index           =   8
      End
      Begin VB.Menu mnuCTBCad 
         Caption         =   "Rateio O&n-Line"
         Index           =   9
      End
      Begin VB.Menu mnuCTBCad 
         Caption         =   "Rateio O&ff-Line"
         Index           =   10
      End
      Begin VB.Menu mnuCTBCad 
         Caption         =   "&Orçamento"
         Index           =   11
      End
      Begin VB.Menu mnuCTBCad 
         Caption         =   "Saldos Iniciais - Centro de Custo/Lucro"
         Index           =   12
      End
      Begin VB.Menu mnuCTBCad 
         Caption         =   "&Padrão de Contabilização"
         Index           =   13
      End
      Begin VB.Menu mnuCTBCad 
         Caption         =   "Plano de Contas Referencial"
         Index           =   14
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   2
      Begin VB.Menu mnuTESCad 
         Caption         =   "&Bancos"
         Index           =   1
      End
      Begin VB.Menu mnuTESCad 
         Caption         =   "&Contas Correntes"
         Index           =   2
      End
      Begin VB.Menu mnuTESCad 
         Caption         =   "&Favorecidos"
         Index           =   3
      End
      Begin VB.Menu mnuTESCad 
         Caption         =   "&Lotes"
         Index           =   4
      End
      Begin VB.Menu mnuTESCad 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuTESCad 
         Caption         =   "&Tabelas Auxiliares"
         Index           =   6
         Begin VB.Menu mnuTESCadTA 
            Caption         =   "Tipos de Aplicação"
            Index           =   1
         End
         Begin VB.Menu mnuTESCadTA 
            Caption         =   "Histórico para Extrato"
            Index           =   2
         End
         Begin VB.Menu mnuTESCadTA 
            Caption         =   "Naturezas"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   3
      Begin VB.Menu mnuCPCad 
         Caption         =   "&Fornecedores"
         Index           =   1
      End
      Begin VB.Menu mnuCPCad 
         Caption         =   "&Portadores"
         Index           =   2
      End
      Begin VB.Menu mnuCPCad 
         Caption         =   "&Lotes"
         Index           =   3
      End
      Begin VB.Menu mnuCPCad 
         Caption         =   "Bancos"
         Index           =   4
      End
      Begin VB.Menu mnuCPCad 
         Caption         =   "Contas Correntes"
         Index           =   5
      End
      Begin VB.Menu mnuCPCad 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuCPCad 
         Caption         =   "&Tabelas Auxiliares"
         Index           =   7
         Begin VB.Menu mnuCPCadTA 
            Caption         =   "Tipos de Fornecedor"
            Index           =   1
         End
         Begin VB.Menu mnuCPCadTA 
            Caption         =   "Condições de Pagamento"
            Index           =   2
         End
         Begin VB.Menu mnuCPCadTA 
            Caption         =   "Categorias de Fornecedor"
            Index           =   7
         End
         Begin VB.Menu mnuCPCadTA 
            Caption         =   "Pagamentos Periódicos"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCPCadTA 
            Caption         =   "Modelos de Email"
            Index           =   9
         End
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   4
      Begin VB.Menu mnuCRCad 
         Caption         =   "&Clientes"
         Index           =   1
      End
      Begin VB.Menu mnuCRCad 
         Caption         =   "C&obradores"
         Index           =   2
      End
      Begin VB.Menu mnuCRCad 
         Caption         =   "&Vendedores"
         Index           =   3
      End
      Begin VB.Menu mnuCRCad 
         Caption         =   "&Transportadoras"
         Index           =   4
      End
      Begin VB.Menu mnuCRCad 
         Caption         =   "&Lotes"
         Index           =   5
      End
      Begin VB.Menu mnuCRCad 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuCRCad 
         Caption         =   "&Tabelas Auxiliares"
         Index           =   7
         Begin VB.Menu mnuCRCadTA 
            Caption         =   "Categorias de Cliente"
            Index           =   1
         End
         Begin VB.Menu mnuCRCadTA 
            Caption         =   "Tipos de Cliente"
            Index           =   2
         End
         Begin VB.Menu mnuCRCadTA 
            Caption         =   "Tipos de Vendedor"
            Index           =   3
         End
         Begin VB.Menu mnuCRCadTA 
            Caption         =   "Regiões de Venda"
            Index           =   4
         End
         Begin VB.Menu mnuCRCadTA 
            Caption         =   "Condições de Pagamento"
            Index           =   5
         End
         Begin VB.Menu mnuCRCadTA 
            Caption         =   "Carteiras de Cobrança"
            Index           =   6
         End
         Begin VB.Menu mnuCRCadTA 
            Caption         =   "Padrões de Cobrança"
            Index           =   7
         End
         Begin VB.Menu mnuCRCadTA 
            Caption         =   "Comissões Avulsas"
            Index           =   11
         End
         Begin VB.Menu mnuCRCadTA 
            Caption         =   "Recebimentos Periódicos"
            Index           =   12
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCRCadTA 
            Caption         =   "Arquivo de Cobrança - Detalhe"
            Index           =   13
         End
         Begin VB.Menu mnuCRCadTA 
            Caption         =   "Arquivo de Cobrança - Tipos de Diferença"
            Index           =   14
         End
         Begin VB.Menu mnuCRCadTA 
            Caption         =   "Modelos de Email"
            Index           =   15
         End
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   5
      Begin VB.Menu mnuESTCad 
         Caption         =   "&Fornecedores"
         Index           =   1
      End
      Begin VB.Menu mnuESTCad 
         Caption         =   "&Almoxarifados"
         Index           =   2
      End
      Begin VB.Menu mnuESTCad 
         Caption         =   "&Produtos"
         Index           =   3
      End
      Begin VB.Menu mnuESTCad 
         Caption         =   "&Kit"
         Index           =   4
      End
      Begin VB.Menu mnuESTCad 
         Caption         =   "&Centro de Custo/Lucro"
         Index           =   5
      End
      Begin VB.Menu mnuESTCad 
         Caption         =   "&Lotes"
         Index           =   6
         Begin VB.Menu mnuESTCadLoteInventario 
            Caption         =   "Lote de Inventário"
         End
         Begin VB.Menu mnuESTCadLoteContabil 
            Caption         =   "Lote Contábil"
         End
         Begin VB.Menu mnuESTCadLoteRastro 
            Caption         =   "Lote de Rastreamento"
         End
         Begin VB.Menu mnuESTCadLoteRastroLoc 
            Caption         =   "Localização dos Lotes"
         End
      End
      Begin VB.Menu mnuESTCad 
         Caption         =   "Transportadoras"
         Index           =   7
      End
      Begin VB.Menu mnuESTCad 
         Caption         =   "Produtos x Almoxarifados"
         Index           =   8
      End
      Begin VB.Menu mnuESTCad 
         Caption         =   "Embalagens"
         Index           =   9
      End
      Begin VB.Menu mnuESTCad 
         Caption         =   "Produto X Desconto"
         Index           =   10
      End
      Begin VB.Menu mnuESTCad 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuESTCad 
         Caption         =   "&Tabelas Auxiliares"
         Index           =   12
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Condições de Pagamento"
            Index           =   1
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Tipos de Fornecedor"
            Index           =   2
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Custo - Estoque"
            Index           =   3
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Custo de Produção"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Controle de Estoque"
            Index           =   5
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Tipos de Produto"
            Index           =   6
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Classe de Unidades de Medida"
            Index           =   7
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Estoque Inicial"
            Index           =   8
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Categorias de Produto"
            Index           =   9
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Categorias de Fornecedor"
            Index           =   10
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Produto x Embalagens"
            Index           =   11
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "-"
            Index           =   12
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Estados"
            Index           =   13
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Séries de Notas Fiscais"
            Index           =   14
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Naturezas de Operação"
            Index           =   15
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Exceções ICMS"
            Index           =   16
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Exceções IPI"
            Index           =   17
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Tipos de Tributação"
            Index           =   18
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Tributação Fornecedores"
            Index           =   19
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Tributação Clientes"
            Index           =   20
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Contratos a Pagar"
            Index           =   21
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Inventário por Terceiro"
            Enabled         =   0   'False
            Index           =   22
            Visible         =   0   'False
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Cores e Variações"
            Index           =   23
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Pinturas"
            Index           =   24
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Coleções"
            Index           =   25
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Declaração de Importação"
            Index           =   26
         End
         Begin VB.Menu mnuESTCadTA 
            Caption         =   "Produto - Grade"
            Index           =   27
         End
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   6
      Begin VB.Menu MnuFATCad 
         Caption         =   "&Clientes"
         Index           =   1
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "&Produtos"
         Index           =   2
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "&Vendedores"
         Index           =   3
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "&Transportadoras"
         Index           =   4
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "&Lotes"
         Index           =   5
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "Formação de preços"
         Index           =   6
         Begin VB.Menu mnuFATCadFP 
            Caption         =   "Custos de MPs e Embalagens"
            Index           =   1
         End
         Begin VB.Menu mnuFATCadFP 
            Caption         =   "Custos Diretos por Produto"
            Index           =   2
         End
         Begin VB.Menu mnuFATCadFP 
            Caption         =   "Custos Fixos por Produto"
            Index           =   3
         End
         Begin VB.Menu mnuFATCadFP 
            Caption         =   "Despesas Variáveis por Cliente"
            Index           =   4
         End
         Begin VB.Menu mnuFATCadFP 
            Caption         =   "Tipos de Frete"
            Index           =   5
         End
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "Tabelas &Auxiliares"
         Index           =   8
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Tipos de Clientes"
            Index           =   1
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Tipos de Produto"
            Index           =   2
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Tipos de Vendedores"
            Index           =   3
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Tipos de Bloqueio"
            Index           =   7
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Formação de Preço"
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Mnemônicos para Formação de Preço"
            Index           =   10
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Séries de Notas Fiscais"
            Index           =   11
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Classe de Unidades de Medida"
            Index           =   13
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Categorias de Cliente"
            Index           =   14
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Categorias de Produto"
            Index           =   15
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Mensagens"
            Index           =   16
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "-"
            Index           =   17
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Estados"
            Index           =   18
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Campos Genéricos"
            Index           =   26
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Contratos a Receber"
            Index           =   29
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "-"
            Index           =   30
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Projetos"
            Index           =   31
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Contrato de Propaganda"
            Index           =   33
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Cliente - Expresso"
            Index           =   39
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Modelos de Email"
            Index           =   40
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Planilha de Comissão"
            Index           =   41
         End
         Begin VB.Menu mnuFATCadTA 
            Caption         =   "Declaração de Exportação"
            Index           =   42
         End
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "Liberação de &Senha"
         Index           =   9
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "Famílias"
         Index           =   10
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "Para Transportadoras"
         Index           =   11
         Begin VB.Menu MnuFATCadTransp 
            Caption         =   "Despachante"
            Index           =   1
         End
         Begin VB.Menu MnuFATCadTransp 
            Caption         =   "Prog. Navio"
            Index           =   2
         End
         Begin VB.Menu MnuFATCadTransp 
            Caption         =   "Tabela de Preço"
            Index           =   3
         End
         Begin VB.Menu MnuFATCadTransp 
            Caption         =   "Itens de Serviço"
            Index           =   4
         End
         Begin VB.Menu MnuFATCadTransp 
            Caption         =   "Tipo de Embalagem"
            Index           =   5
         End
         Begin VB.Menu MnuFATCadTransp 
            Caption         =   "Origem / Destino"
            Index           =   6
         End
         Begin VB.Menu MnuFATCadTransp 
            Caption         =   "Tipo de Container"
            Index           =   7
         End
         Begin VB.Menu MnuFATCadTransp 
            Caption         =   "Serviço x Item de Serviço"
            Index           =   8
         End
         Begin VB.Menu MnuFATCadTransp 
            Caption         =   "Documento"
            Index           =   9
         End
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "Vendas"
         Index           =   13
         Begin VB.Menu MnuFATCadVend 
            Caption         =   "Regiões de Venda"
            Index           =   1
         End
         Begin VB.Menu MnuFATCadVend 
            Caption         =   "Canais de Venda"
            Index           =   2
         End
         Begin VB.Menu MnuFATCadVend 
            Caption         =   "Tabelas de Preço"
            Index           =   3
         End
         Begin VB.Menu MnuFATCadVend 
            Caption         =   "Previsão de Vendas"
            Index           =   4
         End
         Begin VB.Menu MnuFATCadVend 
            Caption         =   "Kit de Venda"
            Index           =   5
         End
         Begin VB.Menu MnuFATCadVend 
            Caption         =   "Condições de Pagamento"
            Index           =   6
         End
         Begin VB.Menu MnuFATCadVend 
            Caption         =   "Rota de Vendas"
            Index           =   17
         End
         Begin VB.Menu MnuFATCadVend 
            Caption         =   "Veículos de Entrega"
            Index           =   18
         End
         Begin VB.Menu MnuFATCadVend 
            Caption         =   "Status do Andamento do PV"
            Index           =   19
         End
         Begin VB.Menu MnuFATCadVend 
            Caption         =   "Tabelas de Preço por Grupo de Produtos"
            Index           =   20
         End
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "Tributação"
         Index           =   14
         Begin VB.Menu MnuFATCadFIS 
            Caption         =   "Naturezas de Operação"
            Index           =   1
         End
         Begin VB.Menu MnuFATCadFIS 
            Caption         =   "Classificação Fiscal"
            Index           =   2
         End
         Begin VB.Menu MnuFATCadFIS 
            Caption         =   "Exceções de ICMS"
            Index           =   3
         End
         Begin VB.Menu MnuFATCadFIS 
            Caption         =   "Exceções de IPI"
            Index           =   4
         End
         Begin VB.Menu MnuFATCadFIS 
            Caption         =   "Tipos de Tributação"
            Index           =   5
         End
         Begin VB.Menu MnuFATCadFIS 
            Caption         =   "Tributação de Fornecedores"
            Index           =   6
         End
         Begin VB.Menu MnuFATCadFIS 
            Caption         =   "Tributação de Clientes"
            Index           =   7
         End
         Begin VB.Menu MnuFATCadFIS 
            Caption         =   "Alíquotas da DAS"
            Index           =   8
         End
         Begin VB.Menu MnuFATCadFIS 
            Caption         =   "Exceções de Pis e Cofins"
            Index           =   9
         End
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "Travel Ace"
         Index           =   16
         Begin VB.Menu MnuFATCadTRV 
            Caption         =   "Tipos de Ocorrências"
            Index           =   1
         End
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "Emissores"
         Index           =   18
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "Tipos de Ocorrências"
         Index           =   19
      End
      Begin VB.Menu MnuFATCad 
         Caption         =   "Acordos"
         Index           =   20
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   7
      Begin VB.Menu mnuCOMCad 
         Caption         =   "&Produtos"
         Index           =   1
      End
      Begin VB.Menu mnuCOMCad 
         Caption         =   "&Fornecedores"
         Index           =   2
      End
      Begin VB.Menu mnuCOMCad 
         Caption         =   "Produto &x Fornecedor"
         Index           =   3
      End
      Begin VB.Menu mnuCOMCad 
         Caption         =   "&Requisitantes"
         Index           =   4
      End
      Begin VB.Menu mnuCOMCad 
         Caption         =   "&Compradores"
         Index           =   5
      End
      Begin VB.Menu mnuCOMCad 
         Caption         =   "Alça&das"
         Index           =   6
      End
      Begin VB.Menu mnuCOMCad 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuCOMCad 
         Caption         =   "Tabelas &Auxiliares"
         Index           =   8
         Begin VB.Menu mnuCOMCadTA 
            Caption         =   "&Tipos de Produto"
            Index           =   1
         End
         Begin VB.Menu mnuCOMCadTA 
            Caption         =   "&Modelos de Requisicões"
            Index           =   2
         End
         Begin VB.Menu mnuCOMCadTA 
            Caption         =   "T&ransportadoras"
            Index           =   3
         End
         Begin VB.Menu mnuCOMCadTA 
            Caption         =   "Con&dicões de Pagamento"
            Index           =   4
         End
         Begin VB.Menu mnuCOMCadTA 
            Caption         =   "Tipos de &Bloqueio"
            Index           =   5
         End
         Begin VB.Menu mnuCOMCadTA 
            Caption         =   "Controle de Esto&que"
            Index           =   6
         End
         Begin VB.Menu mnuCOMCadTA 
            Caption         =   "&Categorias de Fornecedor"
            Index           =   7
         End
         Begin VB.Menu mnuCOMCadTA 
            Caption         =   "&Notas de Pedidos de Compra"
            Index           =   8
         End
         Begin VB.Menu mnuCOMCadTA 
            Caption         =   "&Produto x Fornecedor - Expresso"
            Index           =   9
         End
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   8
      Begin VB.Menu mnuFISCad 
         Caption         =   "Naturezas de Operação"
         Index           =   1
      End
      Begin VB.Menu mnuFISCad 
         Caption         =   "Tipos de Tributação"
         Index           =   2
      End
      Begin VB.Menu mnuFISCad 
         Caption         =   "Exceções ICMS"
         Index           =   3
      End
      Begin VB.Menu mnuFISCad 
         Caption         =   "Exceções IPI"
         Index           =   4
      End
      Begin VB.Menu mnuFISCad 
         Caption         =   "Tributação Fornecedores"
         Index           =   5
      End
      Begin VB.Menu mnuFISCad 
         Caption         =   "Tributação Clientes"
         Index           =   6
      End
      Begin VB.Menu mnuFISCad 
         Caption         =   "Tipos de Registro p/ Apuração de ICMS"
         Index           =   7
      End
      Begin VB.Menu mnuFISCad 
         Caption         =   "Tipos de Registro p/ Apuração de IPI"
         Index           =   8
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   9
      Begin VB.Menu mnuLJCad 
         Caption         =   "Clientes"
         Index           =   1
      End
      Begin VB.Menu mnuLJCad 
         Caption         =   "Produtos"
         Index           =   2
      End
      Begin VB.Menu mnuLJCad 
         Caption         =   "Operadores"
         Index           =   3
      End
      Begin VB.Menu mnuLJCad 
         Caption         =   "Caixas"
         Index           =   4
      End
      Begin VB.Menu mnuLJCad 
         Caption         =   "Vendedores"
         Index           =   5
      End
      Begin VB.Menu mnuLJCad 
         Caption         =   "Emissores de Cupons Fiscais"
         Index           =   6
      End
      Begin VB.Menu mnuLJCad 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuLJCad 
         Caption         =   "Tabelas Auxiliares"
         Index           =   8
         Begin VB.Menu mnuLJCadTA 
            Caption         =   "Tabelas de Preços"
            Index           =   1
         End
         Begin VB.Menu mnuLJCadTA 
            Caption         =   "Produto x Desconto"
            Index           =   2
         End
         Begin VB.Menu mnuLJCadTA 
            Caption         =   "Meios de Pagamento"
            Index           =   3
         End
         Begin VB.Menu mnuLJCadTA 
            Caption         =   "Redes"
            Index           =   4
         End
         Begin VB.Menu mnuLJCadTA 
            Caption         =   "Teclados"
            Index           =   5
         End
         Begin VB.Menu mnuLJCadTA 
            Caption         =   "Impressoras Fiscais"
            Index           =   6
         End
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   10
      Begin VB.Menu mnuPCPCad 
         Caption         =   "Produtos"
         Index           =   1
      End
      Begin VB.Menu mnuPCPCad 
         Caption         =   "Almoxarifados"
         Index           =   2
      End
      Begin VB.Menu mnuPCPCad 
         Caption         =   "Kit"
         Index           =   3
      End
      Begin VB.Menu mnuPCPCad 
         Caption         =   "Lote de Rastreamento"
         Index           =   4
      End
      Begin VB.Menu mnuPCPCad 
         Caption         =   "Produtos x Almoxarifados"
         Index           =   5
      End
      Begin VB.Menu mnuPCPCad 
         Caption         =   "Embalagens"
         Index           =   6
      End
      Begin VB.Menu mnuPCPCad 
         Caption         =   "Máquinas - Equipamentos"
         Index           =   7
      End
      Begin VB.Menu mnuPCPCad 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuPCPCad 
         Caption         =   "Tabelas Auxiliares"
         Index           =   9
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Controle de Estoque"
            Index           =   1
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Tipos de Produto"
            Index           =   2
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Classe de Unidades de Medida"
            Index           =   3
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Estoque Inicial"
            Index           =   4
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Categorias de Produto"
            Index           =   5
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Produto x Embalagem"
            Index           =   6
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Testes de Controle de Qualidade"
            Index           =   7
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Competências"
            Index           =   8
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Centros de Trabalho"
            Index           =   9
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Taxas de Produção"
            Index           =   10
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Roteiros de Fabricação"
            Index           =   11
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "-"
            Index           =   12
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Tipos de Mão-de-Obra"
            Index           =   13
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Usuários da Produção"
            Index           =   18
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Certificados"
            Index           =   19
         End
         Begin VB.Menu mnuPCPCadTA 
            Caption         =   "Cursos/Exames"
            Index           =   20
         End
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   11
      Begin VB.Menu mnuCRMCad 
         Caption         =   "Clientes"
         Index           =   1
      End
      Begin VB.Menu mnuCRMCad 
         Caption         =   "Atendentes"
         Index           =   2
      End
      Begin VB.Menu mnuCRMCad 
         Caption         =   "Campos Genéricos"
         Index           =   3
      End
      Begin VB.Menu mnuCRMCad 
         Caption         =   "Vendedores"
         Index           =   4
      End
      Begin VB.Menu mnuCRMCad 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuCRMCad 
         Caption         =   "Tabelas Auxiliares"
         Index           =   6
         Begin VB.Menu mnuCRMCadTA 
            Caption         =   "Categorias de Clientes"
            Index           =   1
         End
         Begin VB.Menu mnuCRMCadTA 
            Caption         =   "Tipos de Clientes"
            Index           =   2
         End
         Begin VB.Menu mnuCRMCadTA 
            Caption         =   "Clientes x Contatos"
            Index           =   3
         End
         Begin VB.Menu mnuCRMCadTA 
            Caption         =   "Tipos de Vendedores"
            Index           =   4
         End
         Begin VB.Menu mnuCRMCadTA 
            Caption         =   "Regiões de Venda"
            Index           =   5
         End
         Begin VB.Menu mnuCRMCadTA 
            Caption         =   "Clientes Futuros"
            Index           =   6
         End
         Begin VB.Menu mnuCRMCadTA 
            Caption         =   "Clientes Futuros x Contatos"
            Index           =   7
         End
         Begin VB.Menu mnuCRMCadTA 
            Caption         =   "Modelos de Email"
            Index           =   8
         End
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   12
      Begin VB.Menu mnuQUACad 
         Caption         =   "Produtos"
         Index           =   1
      End
      Begin VB.Menu mnuQUACad 
         Caption         =   "Testes de Controle de Qualidade"
         Index           =   2
      End
      Begin VB.Menu mnuQUACad 
         Caption         =   "Produto x Testes"
         Index           =   3
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   13
      Begin VB.Menu mnuPRJCad 
         Caption         =   "Produtos"
         Index           =   1
      End
      Begin VB.Menu mnuPRJCad 
         Caption         =   "Clientes"
         Index           =   2
      End
      Begin VB.Menu mnuPRJCad 
         Caption         =   "Fornecedores"
         Index           =   3
      End
      Begin VB.Menu mnuPRJCad 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuPRJCad 
         Caption         =   "Projetos"
         Index           =   5
      End
      Begin VB.Menu mnuPRJCad 
         Caption         =   "Organogramas"
         Index           =   6
      End
      Begin VB.Menu mnuPRJCad 
         Caption         =   "Etapas"
         Index           =   7
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Ca&dastros"
      Index           =   14
      Begin VB.Menu mnuSRVCad 
         Caption         =   "&Clientes"
         Index           =   1
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "&Produtos"
         Index           =   2
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "&Tipo de Garantia"
         Index           =   10
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "&Garantia"
         Index           =   11
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "&Contratos a Receber"
         Index           =   20
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "&Contrato de Manutenção"
         Index           =   21
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "-"
         Index           =   25
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "&Tipo de Mão de Obra"
         Index           =   30
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "&Mão de Obra"
         Index           =   31
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "-"
         Index           =   35
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "&Máquinas"
         Index           =   40
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "&Competências"
         Index           =   41
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "&Centros de Trabalho"
         Index           =   42
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "&Roteiros"
         Index           =   43
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "-"
         Index           =   45
      End
      Begin VB.Menu mnuSRVCad 
         Caption         =   "&Tabelas Auxiliares"
         Index           =   50
         Begin VB.Menu mnuSRVCadTA 
            Caption         =   "Tipos de Bloqueios"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   1
      Begin VB.Menu mnuCTBRot 
         Caption         =   "&Atualização de Lote"
         Index           =   1
      End
      Begin VB.Menu mnuCTBRot 
         Caption         =   "Apuração de &Períodos"
         Index           =   2
      End
      Begin VB.Menu mnuCTBRot 
         Caption         =   "Apuração de &Exercício"
         Index           =   3
      End
      Begin VB.Menu mnuCTBRot 
         Caption         =   "&Fechamento de Exercício"
         Index           =   4
      End
      Begin VB.Menu mnuCTBRot 
         Caption         =   "&Reprocessamento"
         Index           =   5
      End
      Begin VB.Menu mnuCTBRot 
         Caption         =   "Re&abertura de Exercício"
         Index           =   6
      End
      Begin VB.Menu mnuCTBRot 
         Caption         =   "Rateio O&ff-Line"
         Index           =   7
      End
      Begin VB.Menu mnuCTBRot 
         Caption         =   "Importação"
         Index           =   8
      End
      Begin VB.Menu mnuCTBRot 
         Caption         =   "Geração de DRE e DRP em Excel"
         Index           =   9
      End
      Begin VB.Menu mnuCTBRot 
         Caption         =   "&Desapuração de Exercício"
         Index           =   10
      End
      Begin VB.Menu mnuCTBRot 
         Caption         =   "&Importação de Rateio Off-Line"
         Index           =   11
      End
      Begin VB.Menu mnuCTBRot 
         Caption         =   "&Geração de Novos Rateios"
         Index           =   12
      End
      Begin VB.Menu mnuCTBRot 
         Caption         =   "SPED"
         Index           =   13
         Begin VB.Menu mnuCTBRotSped 
            Caption         =   "Diário"
            Index           =   1
         End
         Begin VB.Menu mnuCTBRotSped 
            Caption         =   "FCont"
            Index           =   2
         End
      End
      Begin VB.Menu mnuCTBRot 
         Caption         =   "Importação de Lçtos da Folha"
         Index           =   14
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   2
      Begin VB.Menu mnuTESRot 
         Caption         =   "&Receber Extrato para Conciliação"
         Index           =   2
      End
      Begin VB.Menu mnuTESRot 
         Caption         =   "&Conciliar Extrato Bancário"
         Index           =   3
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   3
      Begin VB.Menu mnuCPRot 
         Caption         =   "Geração de Arquivo de Pagamentos"
         Index           =   1
      End
      Begin VB.Menu mnuCPRot 
         Caption         =   "Envio de email de cobrança de envio de Fatura"
         Index           =   2
      End
      Begin VB.Menu mnuCPRot 
         Caption         =   "Envio de email de aviso de pagamento"
         Index           =   3
      End
      Begin VB.Menu mnuCPRot 
         Caption         =   "Importação de Cartas Frete"
         Index           =   4
      End
      Begin VB.Menu mnuCPRot 
         Caption         =   "Retorno do Arquivo de Pagamento"
         Index           =   5
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   4
      Begin VB.Menu mnuCRRot 
         Caption         =   "&Atualizar Pagtos de Comissões"
         Index           =   1
      End
      Begin VB.Menu mnuCRRot 
         Caption         =   "&Comunicação Bancária"
         Index           =   3
         Begin VB.Menu mnuCRRotTituloCobranca 
            Caption         =   "Remessa de Títulos em Cobrança"
         End
         Begin VB.Menu mnuCRRotRetornoTitulos 
            Caption         =   "Retorno de Títulos em Cobrança"
         End
      End
      Begin VB.Menu mnuCRRot 
         Caption         =   "&Emissão de Boletos"
         Index           =   4
      End
      Begin VB.Menu mnuCRRot 
         Caption         =   "Emissão de Duplicatas"
         Index           =   5
      End
      Begin VB.Menu mnuCRRot 
         Caption         =   "Reajuste de Títulos"
         Index           =   6
      End
      Begin VB.Menu mnuCRRot 
         Caption         =   "Envio de Emails"
         Index           =   7
         Begin VB.Menu mnuCRRotEmail 
            Caption         =   "Email de Aviso de Cobrança"
            Index           =   1
         End
         Begin VB.Menu mnuCRRotEmail 
            Caption         =   "Email de Cobrança de Atrasados"
            Index           =   2
         End
         Begin VB.Menu mnuCRRotEmail 
            Caption         =   "Email de Agradecimento por Pagamento"
            Index           =   3
         End
      End
      Begin VB.Menu mnuCRRot 
         Caption         =   "Importação de Títulos - Após Furnas"
         Index           =   9
      End
      Begin VB.Menu mnuCRRot 
         Caption         =   "Exportação de Associados - Após Furnas"
         Index           =   10
      End
      Begin VB.Menu mnuCRRot 
         Caption         =   "Importação de Faturas"
         Index           =   11
      End
      Begin VB.Menu mnuCRRot 
         Caption         =   "Importação de Extratos de Redes de Cartão"
         Index           =   13
      End
      Begin VB.Menu mnuCRRot 
         Caption         =   "Baixas de Títulos de Cartão pelos Extratos"
         Index           =   14
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   5
      Begin VB.Menu mnuESTRot 
         Caption         =   "&Atualização de Lotes"
         Index           =   1
      End
      Begin VB.Menu mnuESTRot 
         Caption         =   "Atualização de &Rastreamento"
         Index           =   2
      End
      Begin VB.Menu mnuESTRot 
         Caption         =   "&Custo Médio de Produção"
         Index           =   3
      End
      Begin VB.Menu mnuESTRot 
         Caption         =   "Abertura do &Mês"
         Index           =   4
      End
      Begin VB.Menu mnuESTRot 
         Caption         =   "C&lassificação ABC"
         Index           =   5
      End
      Begin VB.Menu mnuESTRot 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuESTRot 
         Caption         =   "Emissão de Notas Fiscais"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuESTRot 
         Caption         =   "Emissão de Notas Fiscais Fatura"
         Index           =   8
      End
      Begin VB.Menu mnuESTRot 
         Caption         =   "Emissão de Notas de Recebimento"
         Index           =   9
      End
      Begin VB.Menu mnuESTRot 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuESTRot 
         Caption         =   "Reprocessamento"
         Index           =   11
      End
      Begin VB.Menu mnuESTRot 
         Caption         =   "Importação de Inventário"
         Index           =   12
      End
      Begin VB.Menu mnuESTRot 
         Caption         =   "Importação de Notas de Raiz de Mandioca"
         Index           =   13
      End
      Begin VB.Menu mnuESTRot 
         Caption         =   "Importação de xmls de NFe/CTe/DI/Pedido"
         Index           =   14
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   6
      Begin VB.Menu mnuFATRot 
         Caption         =   "&Atualização de Preços"
         Index           =   1
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Atualização de &Rastreamento"
         Index           =   2
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Romaneio de Separação"
         Index           =   4
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Emissão de Notas Fiscais"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Emissão de Notas Fiscais Fatura"
         Index           =   6
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Emissão de Faturas"
         Index           =   7
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Emissão de Duplicatas"
         Index           =   8
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Faturamento de Contratos"
         Index           =   9
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Formação de Preços"
         Index           =   10
         Begin VB.Menu mnuFATRotFP 
            Caption         =   "Análise de Margem de Contribuição"
            Index           =   1
         End
         Begin VB.Menu mnuFATRotFP 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuFATRotFP 
            Caption         =   "Rateio de Custos Diretos"
            Index           =   3
         End
         Begin VB.Menu mnuFATRotFP 
            Caption         =   "Rateio de Custos Fixos"
            Index           =   4
         End
         Begin VB.Menu mnuFATRotFP 
            Caption         =   "Cálculo de Preços"
            Index           =   5
         End
         Begin VB.Menu mnuFATRotFP 
            Caption         =   "Ajuste de Preços"
            Index           =   6
         End
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Exportação de Projetos"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Geração de Arquivo de Pedidos com Lote"
         Index           =   12
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Exportar Notas Fiscais"
         Index           =   14
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Importar Notas Fiscais"
         Index           =   15
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Exportação de dados"
         Index           =   17
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Importação de dados"
         Index           =   18
         Begin VB.Menu mnuFATRotImp 
            Caption         =   "Automática"
            Index           =   1
         End
         Begin VB.Menu mnuFATRotImp 
            Caption         =   "Manual"
            Index           =   2
         End
         Begin VB.Menu mnuFATRotImp 
            Caption         =   "Xmls de NFe/CTe/DI/Pedido"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Geração de Arquivos de Lote de RPS"
         Index           =   20
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Importação de NFe Municipal"
         Index           =   21
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Travel Ace"
         Index           =   22
         Begin VB.Menu mnuTRVFATRot 
            Caption         =   "Extração de Dados do Sigav"
            Index           =   1
         End
         Begin VB.Menu mnuTRVFATRot 
            Caption         =   "Regerar Faturas em html"
            Index           =   2
         End
         Begin VB.Menu mnuTRVFATRot 
            Caption         =   "Geração de comissões internas"
            Index           =   3
         End
         Begin VB.Menu mnuTRVFATRot 
            Caption         =   "Geração de relatórios em excel"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Exportar Notas Fiscais - Harmonia"
         Index           =   23
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Nota Fiscal Paulista"
         Index           =   24
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Importar Pedidos de Venda - SET"
         Index           =   25
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "-"
         Index           =   26
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Regerar Fatura .html"
         Index           =   27
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Gerar Comissão"
         Index           =   28
         Begin VB.Menu mnuFATRotComis 
            Caption         =   "Externa"
            Index           =   1
         End
         Begin VB.Menu mnuFATRotComis 
            Caption         =   "Interna"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Recálculo de vendedores e seus % nos vouchers"
         Index           =   29
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "-"
         Index           =   30
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Nota Fiscal Eletrônica Federal"
         Index           =   31
         Begin VB.Menu mnuFATRotNFe 
            Caption         =   "Geração de Lote de Envio"
            Index           =   1
         End
         Begin VB.Menu mnuFATRotNFe 
            Caption         =   "Consulta de Lote"
            Index           =   2
         End
         Begin VB.Menu mnuFATRotNFe 
            Caption         =   "Envio por Email"
            Index           =   3
         End
         Begin VB.Menu mnuFATRotNFe 
            Caption         =   "Inutilização de Faixa"
            Index           =   4
         End
         Begin VB.Menu mnuFATRotNFe 
            Caption         =   "Consulta de NFe"
            Index           =   5
         End
         Begin VB.Menu mnuFATRotNFe 
            Caption         =   "Ocorrência de Contingência"
            Index           =   6
         End
         Begin VB.Menu mnuFATRotNFe 
            Caption         =   "Exportação dos Xml"
            Index           =   7
         End
         Begin VB.Menu mnuFATRotNFe 
            Caption         =   "Carta de Correção"
            Index           =   8
         End
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "-"
         Index           =   32
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Arquivo para o Sistema de Comissões"
         Index           =   33
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "-"
         Index           =   34
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Nota Fiscal de Serviço Eletrônica"
         Index           =   35
         Begin VB.Menu mnuFATRotNFSE 
            Caption         =   "Geração de Lote de Envio"
            Index           =   1
         End
         Begin VB.Menu mnuFATRotNFSE 
            Caption         =   "Consulta de Lote"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "-"
         Index           =   50
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Exportação - L'oreal"
         Index           =   55
      End
      Begin VB.Menu mnuFATRot 
         Caption         =   "Exportação - Gerenciador de Vendas"
         Index           =   56
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   7
      Begin VB.Menu mnuComRot 
         Caption         =   "Parâmetros de Ponto de Pedido"
         Index           =   1
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   8
      Begin VB.Menu mnuFISRot 
         Caption         =   "Fechamento de Livro"
         Index           =   1
      End
      Begin VB.Menu mnuFISRot 
         Caption         =   "Geração Arq. Sintegra"
         Index           =   2
      End
      Begin VB.Menu mnuFISRot 
         Caption         =   "Geração Arq. IPI (DIPI)"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFISRot 
         Caption         =   "Geração IN 068/95"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFISRot 
         Caption         =   "Geração de Arquivos IN86"
         Index           =   5
      End
      Begin VB.Menu mnuFISRot 
         Caption         =   "Reabertura de Livro"
         Index           =   6
      End
      Begin VB.Menu mnuFISRot 
         Caption         =   "Livros de E/S - Atualização de Cadastros"
         Index           =   7
      End
      Begin VB.Menu mnuFISRot 
         Caption         =   "EFD - ICMS/IPI"
         Index           =   8
      End
      Begin VB.Menu mnuFISRot 
         Caption         =   "EFD - Contribuições"
         Index           =   9
      End
      Begin VB.Menu mnuFISRot 
         Caption         =   "DFC - Declaração Fisco Contábil"
         Index           =   10
      End
      Begin VB.Menu mnuFISRot 
         Caption         =   "ECF - Escrituração Contábil Fiscal"
         Index           =   11
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   9
      Begin VB.Menu mnuLJRot 
         Caption         =   "Operação de Arquivos para Caixa Central"
         Index           =   1
      End
      Begin VB.Menu mnuLJRot 
         Caption         =   "Operação de Arquivos para Backoffice"
         Index           =   2
      End
      Begin VB.Menu mnuLJRot 
         Caption         =   "Carga Balança"
         Index           =   3
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   10
      Begin VB.Menu mnuPCPRot 
         Caption         =   "Custo Médio de Produção"
         Index           =   1
      End
      Begin VB.Menu mnuPCPRot 
         Caption         =   "Planejamento da Necessidade de Materias (MRP)"
         Index           =   2
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   11
      Begin VB.Menu mnuCRMRot 
         Caption         =   "Envio de Email para Clientes"
         Index           =   1
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   12
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   13
      Begin VB.Menu mnuPRJRot 
         Caption         =   "Exportação de relatórios para Excel"
         Index           =   1
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "&Rotinas"
      Index           =   14
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   1
      Begin VB.Menu mnuCTBConfig 
         Caption         =   "&Exercício"
         Index           =   1
      End
      Begin VB.Menu mnuCTBConfig 
         Caption         =   "Exercício - Períodos"
         Index           =   2
      End
      Begin VB.Menu mnuCTBConfig 
         Caption         =   "&Configurações"
         Index           =   3
      End
      Begin VB.Menu mnuCTBConfig 
         Caption         =   "&Segmentos"
         Index           =   4
      End
      Begin VB.Menu mnuCTBConfig 
         Caption         =   "Campos &Globais"
         Index           =   5
      End
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   2
      Begin VB.Menu mnuTESConfig 
         Caption         =   "Con&figurações"
         Index           =   1
      End
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   3
      Begin VB.Menu mnuCPConfig 
         Caption         =   "&Configuração"
         Index           =   1
      End
      Begin VB.Menu mnuCPConfig 
         Caption         =   "&Segmentos"
         Index           =   2
      End
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   4
      Begin VB.Menu mnuCRConfig 
         Caption         =   "&Configuração"
         Index           =   1
      End
      Begin VB.Menu mnuCRConfig 
         Caption         =   "Co&brança Eletrônica"
         Index           =   2
      End
      Begin VB.Menu mnuCRConfig 
         Caption         =   "&Segmentos"
         Index           =   3
      End
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   5
      Begin VB.Menu mnuESTConfig 
         Caption         =   "&Segmentos"
         Index           =   1
      End
      Begin VB.Menu mnuESTConfig 
         Caption         =   "&Configuração"
         Index           =   2
      End
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   6
      Begin VB.Menu mnuFATConfig 
         Caption         =   "Configurações"
         Index           =   1
      End
      Begin VB.Menu mnuFATConfig 
         Caption         =   "&Habilitação de Autorização de Crédito"
         Index           =   2
      End
      Begin VB.Menu mnuFATConfig 
         Caption         =   "&Segmentos"
         Index           =   3
      End
      Begin VB.Menu mnuFATConfig 
         Caption         =   "&Regras de Comissões"
         Index           =   4
      End
      Begin VB.Menu mnuFATConfig 
         Caption         =   "Formação de Preços"
         Index           =   5
         Begin VB.Menu mnuFATConfigFP 
            Caption         =   "Mnemônicos"
            Index           =   1
         End
         Begin VB.Menu mnuFATConfigFP 
            Caption         =   "Planilhas"
            Index           =   2
         End
         Begin VB.Menu mnuFATConfigFP 
            Caption         =   "Análise de Margem de Contribuição"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFATConfig 
         Caption         =   "Travel Ace"
         Index           =   6
         Begin VB.Menu mnuTRVFATConfig 
            Caption         =   "Configurações"
            Index           =   1
         End
      End
      Begin VB.Menu mnuFATConfig 
         Caption         =   "Outras Configurações"
         Index           =   7
      End
      Begin VB.Menu mnuFATConfig 
         Caption         =   "Envio de Email"
         Index           =   8
      End
      Begin VB.Menu mnuFATConfig 
         Caption         =   "Regras de Mensagens"
         Index           =   9
      End
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   7
      Begin VB.Menu mnuCOMConfig 
         Caption         =   "Configurações"
         Index           =   1
      End
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   8
      Begin VB.Menu mnuFISConfig 
         Caption         =   "Configuração Geral"
         Index           =   1
      End
      Begin VB.Menu mnuFISConfig 
         Caption         =   "Configuração por Tributo"
         Index           =   2
      End
      Begin VB.Menu mnuFISConfig 
         Caption         =   "Alíquotas ICMS"
         Index           =   3
      End
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   9
      Begin VB.Menu mnuLJConfig 
         Caption         =   "Configuração"
         Index           =   1
      End
      Begin VB.Menu mnuLJConfig 
         Caption         =   "Teclado"
         Index           =   2
      End
      Begin VB.Menu mnuLJConfig 
         Caption         =   "Associação Vendedor X Loja"
         Index           =   3
      End
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   10
      Begin VB.Menu mnuPCPConfig 
         Caption         =   "Segmentos"
         Index           =   1
      End
      Begin VB.Menu mnuPCPConfig 
         Caption         =   "Configuração"
         Index           =   2
      End
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   11
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   12
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   13
      Begin VB.Menu mnuPRJConfig 
         Caption         =   "Segmentos"
         Index           =   1
      End
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Con&figurações"
      Index           =   14
      Begin VB.Menu mnuSRVConfig 
         Caption         =   "&Configuração"
         Index           =   1
      End
   End
   Begin VB.Menu mnuJanelas 
      Caption         =   "&Janelas"
      WindowList      =   -1  'True
      Begin VB.Menu mnuJanelasCascade 
         Caption         =   "Cascata"
         Index           =   2
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "Aju&da"
      Begin VB.Menu mnuAjud 
         Caption         =   "&Índice"
         Index           =   1
      End
      Begin VB.Menu mnuAjud 
         Caption         =   "&Sobre o Corporator"
         Index           =   2
      End
      Begin VB.Menu mnuAjud 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuAjud 
         Caption         =   "Suporte Online"
         Index           =   4
      End
      Begin VB.Menu mnuAjud 
         Caption         =   "Procurar por Atualizações"
         Index           =   5
      End
   End
End
Attribute VB_Name = "PrincipalNovo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function FindExecutable Lib "shell32.dll" Alias _
   "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
   String, ByVal lpResult As String) As Long
   
Private Declare Function WinHelp Lib "user32" Alias _
        "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile _
        As String, ByVal wCommand As Long, ByVal dwData _
        As Long) As Long
        
Private Const HELP_CONTENTS As Long = &H3&
Private Const HELP_QUIT As Long = &H2

Const BOTAO_FILIAL_ATIVA_COR = -2147483632 '-2147483648# '-2147483629
Const BOTAO_FILIAL_NAO_ATIVA_COR = -2147483633

Private Declare Function SetKeyboardState Lib "user32" _
    (lppbKeyState As Byte) As Long
Private Declare Function GetKeyState Lib "user32" _
    (ByVal nVirtKey As Long) As Integer
  
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Const VK_NUMLOCK = &H90
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

Private Const SGE_DATA_SIMULADO = DATA_NULA

Public objAdmSeta As New AdmSeta
Dim giModuloAnterior As Integer

'Funcao que retorna ordenacao lexicografica
Private Declare Function Conexao_ObterTipoOrdenacaoInt Lib "ADSQLMN.DLL" Alias "AD_Conexao_ObterTipoOrdenacao" (ByVal lConexao As Long, iTipo As Integer) As Long

'Public GL_objKeepAlive As AdmKeepAlive

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Const REG_SZ = 1
Private Const ERROR_SUCCESS = 0&
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002

Private Const MENU_FOLHA = "1"
Private Const MENU_JA_VERIFICADO = "2"
Const NENHUM_MODULO = -1

Private Const SEPARADOR = "-"

Dim iCountBkp As Integer
'Dim iExeBkp As Integer

'Constantes de Indices para o Menu:

'**************************************************TRIP*******************************************
Const MENU_FAT_MOV_APORTE_EMISSAO = 1
Const MENU_FAT_MOV_APORTE_LIBERACAO = 2

Const MENU_FAT_MOV_VOU_EMISSAO = 1
Const MENU_FAT_MOV_VOU_MANUTENCAO = 2
Const MENU_FAT_MOV_VOU_COMISSAO = 3

Const MENU_FAT_MOV_OCR_EMISSAO = 1
Const MENU_FAT_MOV_OCR_LIBERACAO = 2

Const MENU_FAT_MOV_FAT_FATURAMENTO = 1
Const MENU_FAT_MOV_FAT_CANCELAMENTO = 2
Const MENU_FAT_MOV_FAT_FATCARTAO = 3

Const MENU_FAT_ROT_COMIS_EXT = 1
Const MENU_FAT_ROT_COMIS_INT = 2

Const MENU_FAT_CON_APORTE_TODOS = 1
Const MENU_FAT_CON_APORTE_NAOFAT = 2
Const MENU_FAT_CON_APORTE_SF = 3
Const MENU_FAT_CON_APORTE_CREDITO = 4
Const MENU_FAT_CON_APORTE_HIST_SF = 6
Const MENU_FAT_CON_APORTE_HIST_CRED = 7

Const MENU_FAT_CON_VOU_TELA = 1
Const MENU_FAT_CON_VOU_TODOS = 2
Const MENU_FAT_CON_VOU_NAOFAT = 3
Const MENU_FAT_CON_VOU_CANCPREV = 4
Const MENU_FAT_CON_VOU_PAX = 6
Const MENU_FAT_CON_VOU_CCSEMAUTO = 7
Const MENU_FAT_CON_VOU_PG = 8

Const MENU_FAT_CON_COMIS_CMRTODAS = 1
Const MENU_FAT_CON_COMIS_CMRBLOQ = 2
Const MENU_FAT_CON_COMIS_CMRLIB = 3
Const MENU_FAT_CON_COMIS_CMCTODAS = 5
Const MENU_FAT_CON_COMIS_CMCBLOQ = 6
Const MENU_FAT_CON_COMIS_CMCLIB = 7
Const MENU_FAT_CON_COMIS_CMCCTODAS = 9
Const MENU_FAT_CON_COMIS_CMCCBLOQ = 10
Const MENU_FAT_CON_COMIS_CMCCLIB = 11
Const MENU_FAT_CON_COMIS_CMETODAS = 13
Const MENU_FAT_CON_COMIS_CMEBLOQ = 14
Const MENU_FAT_CON_COMIS_CMELIB = 15
Const MENU_FAT_CON_COMIS_CMA = 17
Const MENU_FAT_CON_COMIS_COMISSEMNF = 19
Const MENU_FAT_CON_COMIS_TODAS = 20

Const MENU_FAT_CON_OCR_TODAS = 1
Const MENU_FAT_CON_OCR_BLOQUEADAS = 2
Const MENU_FAT_CON_OCR_LIBERADAS = 3
Const MENU_FAT_CON_OCR_INATIVACAO = 4

Const MENU_FAT_CON_FAT_CONSULTAFATURA = 1
Const MENU_FAT_CON_FAT_CANCELAFATURA = 2
Const MENU_FAT_CON_FAT_TITRECSEMNF = 3
Const MENU_FAT_CON_FAT_DOCPARAFAT = 5
Const MENU_FAT_CON_FAT_DOCFATURADO = 6
'**************************************************TRIP*******************************************

Const MENU_FAT_CAD_TRANSP_DESPACHANTE = 1
Const MENU_FAT_CAD_TRANSP_NAVIO = 2
Const MENU_FAT_CAD_TRANSP_TABELAPRECO = 3
Const MENU_FAT_CAD_TRANSP_ITENSSERVICO = 4
Const MENU_FAT_CAD_TRANSP_TIPOCARGA = 5
Const MENU_FAT_CAD_TRANSP_ORIGEMDESTINO = 6
Const MENU_FAT_CAD_TRANSP_TIPOCONTAINER = 7
Const MENU_FAT_CAD_TRANSP_SERVITEMSERV = 8
Const MENU_FAT_CAD_TRANSP_DOCUMENTO = 9

Const MENU_FAT_MOV_TRANSP_PEDIDOCOTACAO = 1
Const MENU_FAT_MOV_TRANSP_PROPOSTACOTACAO = 2
Const MENU_FAT_MOV_TRANSP_SOLICITACAOSERVICO = 3
Const MENU_FAT_MOV_TRANSP_COMPROVANTESERVICO = 4
Const MENU_FAT_MOV_TRANSP_CONHECIMENTOFRETE = 5
Const MENU_FAT_MOV_TRANSP_CONHECIMENTOFRETEFAT = 6

Const MENU_TRVFAT_CAD_TIPOOCR = 1

Const MENU_TRVFAT_MOV_OCR = 15
Const MENU_TRVFAT_MOV_LIB_OCR = 20
Const MENU_TRVFAT_MOV_APORTE = 40
Const MENU_TRVFAT_MOV_LIB_APORTE = 45
Const MENU_TRVFAT_MOV_FAT = 55
Const MENU_TRVFAT_MOV_GER_NF = 75
Const MENU_TRVFAT_MOV_ACORDOS = 30
Const MENU_TRVFAT_MOV_CANCFAT = 65
Const MENU_TRVFAT_MOV_VOUCOMI = 5
Const MENU_TRVFAT_MOV_FAT_CARTAO = 60
Const MENU_TRVFAT_MOV_OCRCASO = 85
Const MENU_TRVFAT_MOV_OCRCASO_LIBCOBER = 90
Const MENU_TRVFAT_MOV_OCRCASO_LIBJUR = 95
Const MENU_TRVFAT_MOV_REEMBOLSO = 100

Const MENU_TRVFAT_CON_VOU_TELA = 1
Const MENU_TRVFAT_CON_VOU_TODOS = 2
Const MENU_TRVFAT_CON_VOU_NAO_FAT = 3
Const MENU_TRVFAT_CON_VOU_PREV_REEMB = 4
Const MENU_TRVFAT_CON_VOU_REP = 5
Const MENU_TRVFAT_CON_VOU_COR = 6
Const MENU_TRVFAT_CON_VOU_EMI = 7
Const MENU_TRVFAT_CON_OCR_TODAS = 9
Const MENU_TRVFAT_CON_OCR_BLOQ = 10
Const MENU_TRVFAT_CON_OCR_LIB = 11
Const MENU_TRVFAT_CON_NVL = 12
Const MENU_TRVFAT_CON_APORTE = 25
Const MENU_TRVFAT_CON_PAGTO_TODOS = 26
Const MENU_TRVFAT_CON_PAGTO_NAO_FAT = 27
Const MENU_TRVFAT_CON_APORTE_SF = 28
Const MENU_TRVFAT_CON_CRED_DISP = 29
Const MENU_TRVFAT_CON_HIST_APORTE_SF = 30
Const MENU_TRVFAT_CON_UTIL_CRED = 31
Const MENU_TRVFAT_CON_TITREC_SEM_NF = 33
Const MENU_TRVFAT_CON_DOCS_A_FAT = 35
Const MENU_TRVFAT_CON_DOCS_FAT = 36
Const MENU_TRVFAT_CON_FAT = 40
Const MENU_TRVFAT_CON_FAT_CANC = 41
Const MENU_TRVFAT_CON_FAT_CRT = 42
Const MENU_TRVFAT_CON_VEND_CALLCENTER = 50
Const MENU_TRVFAT_CON_POS_GERAL = 51
Const MENU_TRVFAT_CON_EST_VENDA = 55
Const MENU_TRVFAT_CON_VOU_BAIXAS = 60

Const MENU_TRVFAT_CON1_OCR_TODAS_ENV = 5
Const MENU_TRVFAT_CON1_OCR_TODAS_ABERTAS = 10
Const MENU_TRVFAT_CON1_OCR_AGUARD_DOCS = 15
Const MENU_TRVFAT_CON1_OCR_TODAS_AUTORIZADAS = 20
Const MENU_TRVFAT_CON1_OCR_AUTO_NAO_FAT = 25
Const MENU_TRVFAT_CON1_OCR_TODAS_FAT_COBR = 27
Const MENU_TRVFAT_CON1_OCR_FAT_COBR_NAO_PAGAS = 28
Const MENU_TRVFAT_CON1_OCR_TODAS_PROCESSO = 30
Const MENU_TRVFAT_CON1_OCR_PROCESSO_ABERTO = 35
Const MENU_TRVFAT_CON1_OCR_TODAS_CONDENADAS = 40
Const MENU_TRVFAT_CON1_OCR_CONDENADAS_NAO_FAT = 45
Const MENU_TRVFAT_CON1_OCR_TODAS_FAT_PROC = 47
Const MENU_TRVFAT_CON1_OCR_FAT_PROC_NAO_PAGAS = 50
Const MENU_TRVFAT_CON1_OCR_TODAS_REEMB = 55
Const MENU_TRVFAT_CON1_OCR_REEMB_NAO_REC = 60
Const MENU_TRVFAT_CON1_TP_COBERTURA = 70
Const MENU_TRVFAT_CON1_TP_JUDICIAL = 75
Const MENU_TRVFAT_CON1_TR_REEMBOLSO = 80
Const MENU_TRVFAT_CON1_VLR_AUTO = 90

Const MENU_TRVFAT_CON2_ACORDOS = 1
Const MENU_TRVFAT_CON2_CMR = 14
Const MENU_TRVFAT_CON2_CMR_BLOQ = 15
Const MENU_TRVFAT_CON2_CMR_LIB = 16
Const MENU_TRVFAT_CON2_CMC = 22
Const MENU_TRVFAT_CON2_CMC_BLOQ = 23
Const MENU_TRVFAT_CON2_CMC_LIB = 24
Const MENU_TRVFAT_CON2_CMCC = 30
Const MENU_TRVFAT_CON2_CMCC_BLOQ = 31
Const MENU_TRVFAT_CON2_CMCC_LIB = 32
Const MENU_TRVFAT_CON2_CMA = 35
Const MENU_TRVFAT_CON2_OVER = 37
Const MENU_TRVFAT_CON2_OVER_BLOQ = 38
Const MENU_TRVFAT_CON2_OVER_LIB = 39
Const MENU_TRVFAT_CON2_OVER_FAT = 40
Const MENU_TRVFAT_CON2_OVER_IH = 41
Const MENU_TRVFAT_CON2_COMI_CC = 45

Const MENU_TRVFAT_CONFIG_CONFIG = 1

Const MENU_TRVFAT_ROT_EXTRAIDADOS = 1
Const MENU_TRVFAT_ROT_REG_FAT = 2
Const MENU_TRVFAT_ROT_GER_COMI_INT = 3
Const MENU_TRVFAT_ROT_GER_RELS_EXCEL = 4

Const MENU_PRJ_REL_PREV_REAL = 1
Const MENU_PRJ_REL_FLUXO = 2
Const MENU_PRJ_REL_MAT = 3
Const MENU_PRJ_REL_MO = 4
Const MENU_PRJ_REL_ACOMP = 5

Const MENU_PRJ_REL_OUTROS = 7
Const MENU_PRJ_REL_GERREL = 9
Const MENU_PRJ_REL_PLANILHAS = 11

Const MENU_PRJ_CAD_PRODUTO = 1
Const MENU_PRJ_CAD_CLIENTE = 2
Const MENU_PRJ_CAD_FORNECEDOR = 3
Const MENU_PRJ_CAD_PROJETO = 5
Const MENU_PRJ_CAD_ORGANOGRAMA = 6
Const MENU_PRJ_CAD_ETAPA = 7

Const MENU_PRJ_MOV_PROPOSTA = 1
Const MENU_PRJ_MOV_CONTRATO = 2
Const MENU_PRJ_MOV_PAGTO = 4
Const MENU_PRJ_MOV_RECEB = 5
Const MENU_PRJ_MOV_APONT = 7

Const MENU_PRJ_CON_CAD_PRODUTO = 1
Const MENU_PRJ_CON_CAD_CLIENTE = 2
Const MENU_PRJ_CON_CAD_FORNECEDOR = 3

Const MENU_PRJ_CON_PROJETO = 2
Const MENU_PRJ_CON_ETAPA = 3
Const MENU_PRJ_CON_PROPOSTA = 4
Const MENU_PRJ_CON_PAGTO = 6
Const MENU_PRJ_CON_RECEB = 7

Const MENU_PRJ_ROT_EXPORTEXCEL = 1

Const MENU_PRJ_CONFIG_SEGMENTOS = 1

Const MENU_QUA_REL_CAD_PRODUTO = 1
Const MENU_QUA_REL_CAD_TESTE = 2
Const MENU_QUA_REL_CAD_PRODUTOTESTE = 3

Const MENU_QUA_REL_FICHA_CONTROLE = 2
Const MENU_QUA_REL_LAUDOS_NF = 3
Const MENU_QUA_REL_NAO_CONFORME = 4
Const MENU_QUA_REL_OUTROS = 7
Const MENU_QUA_REL_GERREL = 9

Const MENU_QUA_REL_PLANILHAS = 11 'Inserido por Wagner

Const MENU_QUA_CAD_PRODUTOS = 1
Const MENU_QUA_CAD_TESTES = 2
Const MENU_QUA_CAD_PRODUTOTESTES = 3

Const MENU_QUA_MOV_RESULTADOS = 1

Const MENU_QUA_CON_RESULTADO_TESTES = 2

Const MENU_QUA_CON_CAD_PRODUTOS = 1
Const MENU_QUA_CON_CAD_TESTES = 2
Const MENU_QUA_CON_CAD_PRODUTOTESTES = 3

Const MENU_CTB_CAD_PLANOCONTA = 1
Const MENU_CTB_CAD_CATEGORIA = 2
Const MENU_CTB_CAD_CCL = 3
Const MENU_CTB_CAD_ASSOCCCL = 4
Const MENU_CTB_CAD_ASSOCCCLCTB = 5
Const MENU_CTB_CAD_HISTPADRAO = 6
Const MENU_CTB_CAD_LOTES = 7
Const MENU_CTB_CAD_DOCAUTO = 8
Const MENU_CTB_CAD_RATEIOON = 9
Const MENU_CTB_CAD_RATEIOOFF = 10
Const MENU_CTB_CAD_ORCAMENTO = 11
Const MENU_CTB_CAD_SALDOINI = 12
Const MENU_CTB_CAD_PADRAOCONTAB = 13
Const MENU_CTB_CAD_PLANOCONTAREF = 14

Const MENU_TES_CAD_BANCOS = 1
Const MENU_TES_CAD_CONTACORRENTE = 2
Const MENU_TES_CAD_FAVORECIDOS = 3
Const MENU_TES_CAD_LOTES = 4
Const MENU_TES_CAD_TA_TIPOAPLIC = 1
Const MENU_TES_CAD_TA_HISTEXTRATO = 2
Const MENU_TES_CAD_TA_NATMOVCTA = 3

Const MENU_CP_CAD_FORNECEDORES = 1
Const MENU_CP_CAD_PORTADORES = 2
Const MENU_CP_CAD_LOTES = 3
Const MENU_CP_CAD_BANCOS = 4
Const MENU_CP_CAD_CONTACORRENTE = 5
            
Const MENU_CP_CAD_TA_TIPOSFORN = 1
Const MENU_CP_CAD_TA_CONDPAG = 2
Const MENU_CP_CAD_TA_FERIADOS = 3
Const MENU_CP_CAD_TA_MENSAGENS = 4
Const MENU_CP_CAD_TA_PAISES = 5
Const MENU_CP_CAD_TA_HISTEXTRATO = 6
Const MENU_CP_CAD_TA_CATFORNECEDOR = 7
Const MENU_CP_CAD_TA_PAGTO_PERIODICO = 8
Const MENU_CP_CAD_TA_COBREMAILPADRAO = 9

Const MENU_CR_CAD_CLIENTES = 1
Const MENU_CR_CAD_COBRADORES = 2
Const MENU_CR_CAD_VENDEDORES = 3
Const MENU_CR_CAD_TRANSPORTADORAS = 4
Const MENU_CR_CAD_LOTES = 5
Const MENU_CR_CAD_TA_CATEGORIACLI = 1
Const MENU_CR_CAD_TA_TIPOSCLI = 2
Const MENU_CR_CAD_TA_TIPOSVEND = 3
Const MENU_CR_CAD_TA_REGIOESVEND = 4
Const MENU_CR_CAD_TA_CONDPAG = 5
Const MENU_CR_CAD_TA_CARTEIRACOBRANCA = 6
Const MENU_CR_CAD_TA_PADROESCOBRANCA = 7
Const MENU_CR_CAD_TA_PAISES = 8
Const MENU_CR_CAD_TA_MENSAGENS = 9
Const MENU_CR_CAD_TA_HISTEXTRATO = 10
Const MENU_CR_CAD_TA_COMISSAVULSA = 11
Const MENU_CR_CAD_TA_RECEB_PERIODICO = 12
Const MENU_CR_CAD_TA_ARQRET_DETALHE = 13
Const MENU_CR_CAD_TA_ARQRET_DIF = 14
Const MENU_CR_CAD_TA_COBREMAILPADRAO = 15

Const MENU_EST_CAD_FORN = 1
Const MENU_EST_CAD_ALMOXARIFADOS = 2
Const MENU_EST_CAD_PRODUTOS = 3
Const MENU_EST_CAD_KIT = 4
Const MENU_EST_CAD_CCL = 5
Const MENU_EST_CAD_TRANSPORTADORAS = 7
Const MENU_EST_CAD_PRODALM = 8
Const MENU_EST_CAD_EMBALAGEM = 9
Const MENU_EST_CAD_PRODUTO_DESCONTO = 10

Const MENU_EST_CAD_TA_CONDPAG = 1
Const MENU_EST_CAD_TA_TIPOSFORN = 2
Const MENU_EST_CAD_TA_CUSTOS = 3
Const MENU_EST_CAD_TA_CUSTOPROD = 4
Const MENU_EST_CAD_TA_ESTOQUE = 5
Const MENU_EST_CAD_TA_TIPOPROD = 6
Const MENU_EST_CAD_TA_UNIDADEMED = 7
Const MENU_EST_CAD_TA_ESTOQUEINI = 8
Const MENU_EST_CAD_TA_CATEGORIAPROD = 9
Const MENU_EST_CAD_TA_CATEGFORNECEDOR = 10
Const MENU_EST_CAD_TA_PRODUTOEMBALAGEM = 11
Const MENU_EST_CAD_TA_ESTADOS = 13
Const MENU_EST_CAD_TA_SERIESNFISC = 14
Const MENU_EST_CAD_TA_NATUREZAOP = 15
Const MENU_EST_CAD_TA_EXCECOESICMS = 16
Const MENU_EST_CAD_TA_EXCECOESIPI = 17
Const MENU_EST_CAD_TA_TIPOTRIB = 18
Const MENU_EST_CAD_TA_TRIBUTACAOFORN = 19
Const MENU_EST_CAD_TA_TRIBUTACAOCLI = 20

'########################################
'Inserido pelo Wagner
Const MENU_EST_CAD_TA_CONTRATOPAG = 21
'########################################
Const MENU_EST_CAD_TA_INVTERC = 22
Const MENU_EST_CAD_TA_CORVAR = 23
Const MENU_EST_CAD_TA_PINTURA = 24
Const MENU_EST_CAD_TA_COLECAO = 25
Const MENU_EST_CAD_TA_DECL_IMPORTACAO = 26
Const MENU_EST_CAD_TA_PROD_GRADE = 27

Const MENU_FAT_CAD_CLIENTES = 1
Const MENU_FAT_CAD_PRODUTOS = 2
Const MENU_FAT_CAD_VENDEDORES = 3
Const MENU_FAT_CAD_TRANSPORTADORAS = 4
Const MENU_FAT_CAD_LOTES = 5
Const MENU_FAT_CAD_GERACAOSENHA = 9
Const MENU_FAT_CAD_FAMILIAS = 10
Const MENU_FAT_CAD_EMI = 18
Const MENU_FAT_CAD_TIPOOCR = 19
Const MENU_FAT_CAD_ACORDO = 20

Const MENU_FAT_CAD_FP_CUSTOEMBMP = 1
Const MENU_FAT_CAD_FP_CUSTODIRPROD = 2
Const MENU_FAT_CAD_FP_CUSTOFIXOPROD = 3
Const MENU_FAT_CAD_FP_DVVCLIENTE = 4
Const MENU_FAT_CAD_FP_TIPOFRETE = 5

Const MENU_FAT_CAD_TA_TIPOCLI = 1
Const MENU_FAT_CAD_TA_TIPOPROD = 2
Const MENU_FAT_CAD_TA_TIPOVEND = 3
Const MENU_FAT_CAD_TA_CODICAOPAG = 4
'Const MENU_FAT_CAD_TA_REGIOESVENDA = 5
'Const MENU_FAT_CAD_TA_CANAISVENDA = 6
Const MENU_FAT_CAD_TA_TIPOBLOQUEIOS = 7
'Const MENU_FAT_CAD_TA_TABPRECOS = 8
Const MENU_FAT_CAD_TA_FORMPRECO = 9
Const MENU_FAT_CAD_TA_MNEMONICOFPRECO = 10
Const MENU_FAT_CAD_TA_SERIESNFISC = 11
'Const MENU_FAT_CAD_TA_PREVISAOVENDA = 12
Const MENU_FAT_CAD_TA_UNIDADEMED = 13
Const MENU_FAT_CAD_TA_CATEGORIACLI = 14
Const MENU_FAT_CAD_TA_CATEGORIAPROD = 15
Const MENU_FAT_CAD_TA_MENSAGENS = 16
Const MENU_FAT_CAD_TA_ESTADOS = 18
'Const MENU_FAT_CAD_TA_NATUREZAOP = 19
'Const MENU_FAT_CAD_TA_CLASFISC = 20
'Const MENU_FAT_CAD_TA_EXCECOESICMS = 21
'Const MENU_FAT_CAD_TA_EXCECOESIPI = 22
'Const MENU_FAT_CAD_TA_TIPOTRIB = 23
'Const MENU_FAT_CAD_TA_TRIBUTACAOFORN = 24
'Const MENU_FAT_CAD_TA_TRIBUTACAOCLI = 25
Const MENU_FAT_CAD_TA_CAMPOSGENERICOS = 26
Const MENU_FAT_CAD_TA_CLIENTECONTATOS = 27
Const MENU_FAT_CAD_TA_ATENDENTES = 28
Const MENU_FAT_CAD_TA_CONTRATOCAD = 29
Const MENU_FAT_CAD_TA_PROJETO = 31
'Const MENU_FAT_CAD_TA_CUSTEIOPROJETO = 32
Const MENU_FAT_CAD_TA_CONTRATOPROG = 33
'Const MENU_FAT_CAD_TA_KITVENDA = 34
'Const MENU_FAT_CAD_TA_DASALIQUOTAS = 37
Const MENU_FAT_CAD_TA_CLIENTEEXPRESSO = 39
Const MENU_FAT_CAD_TA_MODELOSEMAIL = 40
Const MENU_FAT_CAD_TA_PLANILHACOMISSOES = 41
Const MENU_FAT_CAD_TA_DEINFO = 42

Const MENU_FAT_CAD_FIS_NATUREZAOP = 1
Const MENU_FAT_CAD_FIS_CLASFISC = 2
Const MENU_FAT_CAD_FIS_EXCECOESICMS = 3
Const MENU_FAT_CAD_FIS_EXCECOESIPI = 4
Const MENU_FAT_CAD_FIS_TIPOTRIB = 5
Const MENU_FAT_CAD_FIS_TRIBUTACAOFORN = 6
Const MENU_FAT_CAD_FIS_TRIBUTACAOCLI = 7
Const MENU_FAT_CAD_FIS_DASALIQUOTAS = 8
Const MENU_FAT_CAD_FIS_EXCECOESPISCOFINS = 9

Const MENU_FAT_CAD_VEND_REGIOESVENDA = 1
Const MENU_FAT_CAD_VEND_CANAISVENDA = 2
Const MENU_FAT_CAD_VEND_TABPRECOS = 3
Const MENU_FAT_CAD_VEND_PREVISAOVENDA = 4
Const MENU_FAT_CAD_VEND_KITVENDA = 5
Const MENU_FAT_CAD_VEND_CODICAOPAG = 6
Const MENU_FAT_CAD_VEND_ROTAS = 17
Const MENU_FAT_CAD_VEND_VEICULOS = 18
Const MENU_FAT_CAD_VEND_PVANDAMENTO = 19
Const MENU_FAT_CAD_VEND_TABPRECOGRUPO = 20

Const MENU_COM_CAD_PRODUTOS = 1
Const MENU_COM_CAD_FORNECEDORES = 2
Const MENU_COM_CAD_FORNECEDORPRODUTOFF = 3
Const MENU_COM_CAD_REQUISITANTES = 4
Const MENU_COM_CAD_COMPRADORES = 5
Const MENU_COM_CAD_ALCADAS = 6
Const MENU_COM_CAD_TA_TIPOSDEPRODUTO = 1
Const MENU_COM_CAD_TA_MODELOSDEREQUISICOES = 2
Const MENU_COM_CAD_TA_TRANSPORTADORAS = 3
Const MENU_COM_CAD_TA_CONDICAOPAGAMENTO = 4
Const MENU_COM_CAD_TA_TIPOSDEBLOQUEIO = 5
Const MENU_COM_CAD_TA_ESTOQUE = 6
Const MENU_COM_CAD_TA_CATEGFORNECEDOR = 7
Const MENU_COM_CAD_TA_NOTASPC = 8
Const MENU_COM_CAD_TA_PRODUTOFORNECEDOR = 9

Const MENU_CTB_MOV_LANCALOTE = 1
Const MENU_CTB_MOV_LAN = 2
Const MENU_CTB_MOV_ESTORNOLOTE = 3
Const MENU_CTB_MOV_ESTORNODOC = 4

Const MENU_TES_MOV_SAQUE = 1
Const MENU_TES_MOV_DEPOSITO = 2
Const MENU_TES_MOV_TRANFERENCIA = 3
Const MENU_TES_MOV_APLICACAO = 4
Const MENU_TES_MOV_RESGATE = 5
Const MENU_TES_MOV_CONCILIACAO = 6

Const MENU_CP_MOV_NFISCFATURA = 1
Const MENU_CP_MOV_NFISCAIS = 2
Const MENU_CP_MOV_FATURAS = 3
Const MENU_CP_MOV_OUTROSTITULOS = 4
Const MENU_CP_MOV_CONFIRMACOBRANCA = 6
Const MENU_CP_MOV_BORDEROPAG = 8
Const MENU_CP_MOV_CHEQUEAUTO = 9
Const MENU_CP_MOV_CHEQUEMANUAIS = 10
Const MENU_CP_MOV_CANCELARPAG = 12
Const MENU_CP_MOV_DEVCREDITOS = 13
Const MENU_CP_MOV_ADIANTAMENTOFORN = 14
'Const MENU_CP_MOV_BAIXAMANUAL = 15
'Const MENU_CP_MOV_CANCELARBAIXA = 16
Const MENU_CP_MOV_COMPENSAR_CHEQUE_PRE = 17
Const MENU_CP_MOV_LIBERA_PAGTO = 18

Const MENU_CP_MOV_BX_BAIXAMANUAL = 1
Const MENU_CP_MOV_BX_BAIXAADIANFORN = 2
Const MENU_CP_MOV_BX_BAIXAMANUAL_CHEQUETERC = 3

Const MENU_CP_MOV_BAIXA_CANCELARBAIXA_MANUAL = 1
Const MENU_CP_MOV_BAIXA_CANCELARBAIXA_SELECAO = 2
Const MENU_CP_MOV_BAIXA_CANCELARBAIXA_ADIANT = 3

Const MENU_CR_MOV_TITULOSREC = 1
Const MENU_CR_MOV_CHEQUEPRE = 2
Const MENU_CR_MOV_BORDEROCHEQUEPRE = 3
Const MENU_CR_MOV_BORDERODESCCHQ = 4
Const MENU_CR_MOV_BORDERODESCCHQEXCLUI = 5
Const MENU_CR_MOV_BORDEROCOBRANCA = 7
Const MENU_CR_MOV_INSTCOBELETRONICA = 8
Const MENU_CR_MOV_CANCELARBORDEROCOB = 9
Const MENU_CR_MOV_TRANFERENCIAMANUAL = 11
Const MENU_CR_MOV_DEVOLUCOESENTRADA = 12
Const MENU_CR_MOV_ADIANTAMENTOCLI = 13
Const MENU_CR_MOV_DEVOLUCAOCHEQUE = 15
Const MENU_CR_MOV_BORDEROCHQPREEXCLUI = 16
Const MENU_CR_MOV_COBRANCA = 17
Const MENU_CR_MOV_HISTORICOCLIENTE = 18

Const MENU_EST_MOV_RECMATFORN = 1
Const MENU_EST_MOV_RECMATCLI = 2
Const MENU_EST_MOV_RECMATFORNCOM = 3
Const MENU_EST_MOV_PRODUCAOSAIDA = 5
Const MENU_EST_MOV_REQCONSUMO = 6
Const MENU_EST_MOV_MOVIMENTOINTERNO = 7
Const MENU_EST_MOV_TRANSFERENCIAS = 8
Const MENU_EST_MOV_PRODUCAOENTRADA = 11
Const MENU_EST_MOV_RESERVA = 12
Const MENU_EST_MOV_INVENTARIO = 14
Const MENU_EST_MOV_INVENTARIOLOTE = 15
Const MENU_EST_MOV_INVENTARIOTERC = 16
Const MENU_EST_MOV_NFISCENT = 18
Const MENU_EST_MOV_NFISCENTCOM = 19
Const MENU_EST_MOV_NFISCENTFAT = 20
Const MENU_EST_MOV_NFISCENTFATCOM = 21
Const MENU_EST_MOV_NFISCENTREM = 22
Const MENU_EST_MOV_NFISCENTDEV = 23
Const MENU_EST_MOV_CANCELAR_NF = 24
Const MENU_EST_MOV_MEDICAO = 25
Const MENU_EST_MOV_DESMEMBRAMENTO = 27

'Daniel 29/05/2002
'Const MENU_FAT_MOV_ORCVENDA = 1
'
'Const MENU_FAT_MOV_PEDVEND = 2
'Const MENU_FAT_MOV_LIBBLOQUEIO = 3
'Const MENU_FAT_MOV_BAIXAMANPED = 4
Const MENU_FAT_MOV_GERANFISC = 6
Const MENU_FAT_MOV_NFISCPED = 7
Const MENU_FAT_MOV_NFISCFATPED = 8
Const MENU_FAT_MOV_NFISC = 9
Const MENU_FAT_MOV_NFISCFATURA = 10
'Janaina
Const MENU_FAT_MOV_NFISCREMPED = 12
'Janaina
Const MENU_FAT_MOV_NFISCREM = 13
Const MENU_FAT_MOV_NFISCDEV = 14
Const MENU_FAT_MOV_CANCELAR_NF = 15
Const MENU_FAT_MOV_GERAFAT = 17
Const MENU_FAT_MOV_COMISSOES = 18
Const MENU_FAT_MOV_CONHECIMENTOFRETE = 19
Const MENU_FAT_MOV_CONHECIMENTOFRETEFAT = 20
Const MENU_FAT_MOV_CONTRATOMED = 21
Const MENU_FAT_MOV_GERNFFAT = 28
Const MENU_FAT_MOV_VEND_MAPA = 35

Const MENU_CTB_CON_PLANOCONTA = 1
Const MENU_CTB_CON_CCL = 2
Const MENU_CTB_CON_CONTACCL = 3
Const MENU_CTB_CON_HISTPADRAO = 4
Const MENU_CTB_CON_LOTES = 5
Const MENU_CTB_CON_LOTEPEND = 6
Const MENU_CTB_CON_LAN = 7
Const MENU_CTB_CON_LANPEND = 8
Const MENU_CTB_CON_DOCAUTO = 9
Const MENU_CTB_CON_ORCAMENTO = 10
Const MENU_CTB_CON_RATEIOON = 11
Const MENU_CTB_CON_RATEIOOFF = 12
Const MENU_CTB_CON_MAPPLANCTAREF = 13
Const MENU_CTB_CON_ASSOCCTAREF = 14

Const MENU_TES_CON_FLUXOCAIXA = 1
Const MENU_TES_CON_APLICACOES = 2
Const MENU_TES_CON_BANCOS = 3
Const MENU_TES_CON_CONTASCORRENTES = 4
Const MENU_TES_CON_CONTASCORRENTESFILIAIS = 5
Const MENU_TES_CON_DEPOSITOS = 6
Const MENU_TES_CON_SAQUES = 7
'Const MENU_TES_CON_TRANSPORTADORAS = 8
Const MENU_TES_CON_TIPOAPLICACAO = 8
Const MENU_TES_CON_TRANSFERENCIA = 9
Const MENU_TES_CON_MOVCC = 10
Const MENU_TES_CON_MOVCC_TF = 11
Const MENU_TES_CON_FLUXOCAIXACTB = 12
Const MENU_TES_CON_FLUXOCAIXACTB1 = 13

Const MENU_CP_CON_CAD_FORNECEDORES = 1
Const MENU_CP_CON_CAD_TIPOFORN = 2
Const MENU_CP_CON_CAD_BANCOS = 4
Const MENU_CP_CON_CAD_CONTASCORRENTESFILIAIS = 5
Const MENU_CP_CON_CAD_PORTADORES = 6
Const MENU_CP_CON_CAD_CONDPAG = 8
        
Const MENU_CP_CON_FORNECEDOR = 2
'Const MENU_CP_CON_TITPAG = 4
'Const MENU_CP_CON_TITPAG_TODOS = 5
'Const MENU_CP_CON_NFISC = 6
'Const MENU_CP_CON_NFISCAL_TODOS = 7
Const MENU_CP_CON_CREDITOPRE = 9
Const MENU_CP_CON_PAGANTECIP = 10
Const MENU_CP_CON_PAGAMENTOS = 11
'Const MENU_CP_CON_TITPAG_TODOS_TF = 12
'Const MENU_CP_CON_NFISCAL_TODOS_TF = 13
'Const MENU_CP_CON_COMISSAO_CARTAO = 14
'Const MENU_CP_CON_BAIXASPAG = 15
'Const MENU_CP_CON_NOTAFATPAGTO = 16
'Const MENU_CP_CON_TITPAG_NAKA = 17

Const MENU_CP_CON_TP_ABERTO = 1
Const MENU_CP_CON_TP_TODOS = 3
Const MENU_CP_CON_TP_TODOS_ET = 5
Const MENU_CP_CON_TP_ATRASADOS = 6
Const MENU_CP_CON_TP_BAIXAS = 8
Const MENU_CP_CON_TP_NAKA = 10

Const MENU_CP_CON_NF_ABERTO = 1
Const MENU_CP_CON_NF_TODOS = 2
Const MENU_CP_CON_NF_TODOS_ET = 4
Const MENU_CP_CON_NF_FAT_BX = 5
Const MENU_CP_CON_NF_CMCC_SEM_FAT = 7

Const MENU_CP_CON_CHQ_EMITIDOS = 1
Const MENU_CP_CON_CHQ_COMP = 2
Const MENU_CP_CON_CHQ_PARCELAS = 3

Const MENU_CR_CON_CAD_CLIENTES = 1
Const MENU_CR_CON_CAD_COBRADORES = 2
Const MENU_CR_CON_CAD_VENDEDORES = 3
Const MENU_CR_CON_CAD_TRANSP = 4
Const MENU_CR_CON_CAD_CONDPAG = 6

Const MENU_CR_CON_CLICONS = 2
'Const MENU_CR_CON_TITULOSREC = 4
'Const MENU_CR_CON_TITULOSREC_TODOS = 5
'Const MENU_CR_CON_TITULOSREC_BAIXADOS = 6
Const MENU_CR_CON_DEBREC = 7
Const MENU_CR_CON_RECANTECIP = 8
Const MENU_CR_CON_BORDCOBRANCA = 9
'Const MENU_CR_CON_VOUCHER = 10
'Const MENU_CR_CON_TITULOSREC_TODOS_TF = 11
'Const MENU_CR_CON_TITULOSRECEBER_TF = 12
'Const MENU_CR_CON_PARCELASRECEBER_TF = 13
'Const MENU_CR_CON_BAIXASREC = 14
Const MENU_CR_CON_PRODCR = 11

Const MENU_CR_CON_TR_ABERTO = 1
Const MENU_CR_CON_TR_BAIXADOS = 2
Const MENU_CR_CON_TR_TODOS = 3
Const MENU_CR_CON_TR_ABERTO_ET = 5
Const MENU_CR_CON_TR_BAIXADOS_ET = 6
Const MENU_CR_CON_TR_TODOS_ET = 7
Const MENU_CR_CON_TR_ATRASADOS = 8
Const MENU_CR_CON_TR_BAIXAS = 10
Const MENU_CR_CON_TR_CARTAO_VOUCHER = 12
Const MENU_CR_CON_TR_BOLETO_CANC = 13
Const MENU_CR_CON_TR_COMISSOES_BX = 14

Const MENU_CR_CON_CHQ_EMITIDOS = 1
Const MENU_CR_CON_CHQ_PARCELAS = 2

Const MENU_EST_CON_CAD_FORNECEDORES = 1
Const MENU_EST_CON_CAD_PRODUTOS = 2
Const MENU_EST_CON_CAD_ALMOXARIFADO = 3
Const MENU_EST_CON_CAD_KIT = 4
Const MENU_EST_CON_CAD_TIPOPROD = 5
Const MENU_EST_CON_CAD_CLASSEUM = 6
Const MENU_EST_CON_CAD_CATPROD = 7
Const MENU_EST_CON_CAD_CATPRODITEM = 8
Const MENU_EST_CON_CAD_PRODCATPRODITEM = 9

Const MENU_EST_CON_MOV_RECMAT = 1
Const MENU_EST_CON_MOV_RECMATFORN = 2
Const MENU_EST_CON_MOV_RECMATCLI = 3
Const MENU_EST_CON_MOV_RESERVAS = 4
Const MENU_EST_CON_MOV_MOVEST = 5
Const MENU_EST_CON_MOV_MOVESTINT = 6
Const MENU_EST_CON_MOV_MOVESTTRANSF = 7
Const MENU_EST_CON_MOV_CONSUMO = 8

Const MENU_EST_CON_INV_INVENTARIO = 1
Const MENU_EST_CON_INV_INVENTARIOLOTE = 2
Const MENU_EST_CON_INV_INVENTARIOLOTEPEND = 3

Const MENU_EST_CON_PRO_OP = 1
Const MENU_EST_CON_PRO_OP_BAIXA = 2
Const MENU_EST_CON_PRO_EMPENHO = 3
Const MENU_EST_CON_PRO_ITENSOP = 4
Const MENU_EST_CON_PRO_REQPROD = 5
Const MENU_EST_CON_PRO_PRODUCAO = 6
Const MENU_EST_CON_PRO_REQPRODOP = 7
Const MENU_EST_CON_PRO_PRODUCAOOP = 8

Const MENU_EST_CON_EST_ESTPROD = 1
Const MENU_EST_CON_EST_ESTPRODFILIAL = 2
Const MENU_EST_CON_EST_ESTPRODTERC = 3
Const MENU_EST_CON_EST_SALDODISP = 4
Const MENU_EST_CON_EST_ESTPROD_MV = 5
Const MENU_EST_CON_PROD_EM_FALTA = 6

Const MENU_EST_CON_NF_NFISCENTTODAS = 1
Const MENU_EST_CON_NF_NFISCENT = 2
Const MENU_EST_CON_NF_NFISCENTFAT = 3
Const MENU_EST_CON_NF_NFISCENTREM = 4
Const MENU_EST_CON_NF_NFISCENTDEV = 5
Const MENU_EST_CON_NF_ITENSNFISCENT = 6

Const MENU_EST_CON_NFSAI_NFISCTODAS = 1
Const MENU_EST_CON_NFSAI_NFISC = 2
Const MENU_EST_CON_NFSAI_NFISCFAT = 3
Const MENU_EST_CON_NFSAI_NFISCPED = 4
Const MENU_EST_CON_NFSAI_NFISCFATPED = 5
Const MENU_EST_CON_NFSAI_NFISCREM = 6
Const MENU_EST_CON_NFSAI_NFISCDEV = 7
Const MENU_EST_CON_NFSAI_ITENSNFISC = 8

Const MENU_EST_CON_RASTRO_LOTES = 1
Const MENU_EST_CON_RASTRO_SALDOS = 2
Const MENU_EST_CON_RASTRO_MOVIMENTOS = 3
Const MENU_EST_CON_RASTRO_SALDOS_PRECO = 10
Const MENU_EST_CON_RASTRO_MOVIMENTOS_PRECO = 11
Const MENU_EST_CON_RASTRO_ITENS_NF = 12
Const MENU_EST_CON_RASTRO_PHAR_LM_SEP = 13
Const MENU_EST_CON_RASTRO_SALDOS_CM = 14

Const MENU_EST_CON_TRV_MOVEST = 1
Const MENU_EST_CON_TRV_MOVEST_TRASNFER = 2

Const MENU_FAT_CON_LOG = 10
Const MENU_FAT_CON_CONTROLENF = 12

Const MENU_FAT_CON_RPS = 14
Const MENU_FAT_CON_NFE = 15

Const MENU_FAT_CON_RPS_TODOS = 1
Const MENU_FAT_CON_RPS_NAOENV = 2
Const MENU_FAT_CON_RPS_NAONFE = 3
Const MENU_FAT_CON_RPS_ARQ = 4
Const MENU_FAT_CON_RPS_HIST = 5

Const MENU_FAT_CON_NFE_TODAS = 1
Const MENU_FAT_CON_NFE_HIST = 2

Const MENU_FAT_CON_CAD_CLIENTES = 1
Const MENU_FAT_CON_CAD_CLIENTESCONSULTA = 2
Const MENU_FAT_CON_CAD_VENDEDORES = 3
Const MENU_FAT_CON_CAD_TRANSPORTADORAS = 4
Const MENU_FAT_CON_CAD_CATPROD = 5
Const MENU_FAT_CON_CAD_CATPRODITEM = 6

Const MENU_FAT_CON_CAD_EMISSORES = 8
Const MENU_FAT_CON_CAD_ACORDOS = 9

Const MENU_FAT_CON_CAD_ITENS_CONTRATO_FAT = 10

Const MENU_FAT_CON_VEN_PEDVEND_ATIVOS = 1
Const MENU_FAT_CON_VEN_PEDVEND_BAIXADOS = 2
Const MENU_FAT_CON_VEN_TABPRECO = 3
Const MENU_FAT_CON_VEN_TABPRECOIT = 4
Const MENU_FAT_CON_VEN_PREVVEND = 5
Const MENU_FAT_CON_VEN_ORCAMENTO = 6
Const MENU_FAT_CON_VEN_TABPRECOITAT = 7
Const MENU_FAT_CON_VEN_ROTA = 10
Const MENU_FAT_CON_VEN_MAPA = 11
Const MENU_FAT_CON_VEN_PVACOMP = 12
Const MENU_FAT_CON_VEN_PVITENS = 13
Const MENU_FAT_CON_VEN_OVITENS = 14
Const MENU_FAT_CON_VEN_PROTHEUS_HIST_VENDAS = 15

Const MENU_FAT_CON_EST_PRODUTOS = 1
Const MENU_FAT_CON_EST_ESTPRODFILIAL = 2
Const MENU_FAT_CON_EST_ESTPROD = 3
Const MENU_FAT_CON_EST_ESTPRODTERC = 4
        
Const MENU_FAT_CON_NF_NFISCTODAS = 1
Const MENU_FAT_CON_NF_NFISC = 2
Const MENU_FAT_CON_NF_NFISCFAT = 3
Const MENU_FAT_CON_NF_NFISCPED = 4
Const MENU_FAT_CON_NF_NFISCFATPED = 5
'Janaina
Const MENU_FAT_CON_NF_NFISCREMPED = 6
'Janaina
Const MENU_FAT_CON_NF_NFISCREM = 7
Const MENU_FAT_CON_NF_NFISCDEV = 8
Const MENU_FAT_CON_NF_SERIESNF = 9
Const MENU_FAT_CON_NF_ITENSNFISC = 10
Const MENU_FAT_CON_NF_CONHECTRANSP = 11
Const MENU_FAT_CON_NF_DIRECT = 12
Const MENU_FAT_CON_NF_AFAT = 13
Const MENU_FAT_CON_NF_BI = 14
        
Const MENU_FAT_CON_NFENT_NFISCENTTODAS = 1
Const MENU_FAT_CON_NFENT_NFISCENT = 2
Const MENU_FAT_CON_NFENT_NFISCENTFAT = 3
Const MENU_FAT_CON_NFENT_NFISCENTREM = 4
Const MENU_FAT_CON_NFENT_NFISCENTDEV = 5
Const MENU_FAT_CON_NFENT_ITENSNFISCENT = 6

Const MENU_FAT_CON_GRAF_AREA = 1
Const MENU_FAT_CON_GRAF_CLIENTE = 2
Const MENU_FAT_CON_GRAF_MENSAL_DOLAR = 3
Const MENU_FAT_CON_GRAF_MENSAL = 4

Const MENU_FAT_CON_NFEFED_LOTE = 1
Const MENU_FAT_CON_NFEFED_LOTELOG = 2
Const MENU_FAT_CON_NFEFED_RETENVI = 3
Const MENU_FAT_CON_NFEFED_STATUSNF = 4
Const MENU_FAT_CON_NFEFED_RETCONSLOTE = 5
Const MENU_FAT_CON_NFEFED_RETCANC = 6
Const MENU_FAT_CON_NFEFED_RETINUTFAIXA = 7

Const MENU_FAT_CON_NFSE_LOTE = 1
Const MENU_FAT_CON_NFSE_LOTELOG = 2
Const MENU_FAT_CON_NFSE_RETENVI = 3
Const MENU_FAT_CON_NFSE_AUTORIZADAS = 4
Const MENU_FAT_CON_NFSE_RETCONSLOTE = 5
Const MENU_FAT_CON_NFSE_RETCANC = 6
Const MENU_FAT_CON_NFSE_SITLOTE = 7

Const MENU_FAT_CON_BENEF_SALDO = 1
Const MENU_FAT_CON_BENEF_DEV = 2

Const MENU_COM_CON_PEDIDOSCOTACAO = 4
Const MENU_COM_CON_CONCORRENCIA = 5

Const MENU_COM_CON_CAD_PRODUTOS = 1
Const MENU_COM_CON_CAD_FORNECEDORES = 2
Const MENU_COM_CON_CAD_PRODFORN = 3
Const MENU_COM_CON_CAD_REQUISITANTES = 4
Const MENU_COM_CON_CAD_COMPRADORES = 5
Const MENU_COM_CON_CAD_ALCADA = 6
       
Const MENU_CTB_REL_CAD_PLANOCONTAS = 1
Const MENU_CTB_REL_CAD_CCL = 2
Const MENU_CTB_REL_CAD_HISTPADRAO = 3
Const MENU_CTB_REL_CAD_LOTESCONTAB = 4
Const MENU_CTB_REL_CAD_LOTESPEND = 5
Const MENU_CTB_REL_CAD_LANDATA = 6
Const MENU_CTB_REL_CAD_LANCCL = 7
Const MENU_CTB_REL_CAD_LANLOTE = 8
Const MENU_CTB_REL_CAD_LANPEND = 9

Const MENU_CTB_REL_BALANVERIF = 3
Const MENU_CTB_REL_RAZAO = 4
Const MENU_CTB_REL_RAZAOAUX = 5
Const MENU_CTB_REL_RAZAOAGLUT = 6
Const MENU_CTB_REL_DIARIO = 7
Const MENU_CTB_REL_DIARIOAUX = 8
Const MENU_CTB_REL_DIARIOAGLUT = 9
Const MENU_CTB_REL_DEMOSTRATIVOS = 10
Const MENU_CTB_REL_BALANPATRI = 11
Const MENU_CTB_REL_OUTROS = 13
Const MENU_CTB_REL_GERREL = 15

Const MENU_CTB_REL_PLANILHAS = 17 'Inserido por Wagner

Const MENU_TES_REL_EXTRATOTES = 1
Const MENU_TES_REL_EXTRATOBANC = 2
Const MENU_TES_REL_POSAPLIC = 3
Const MENU_TES_REL_BORDEROPRE = 4
Const MENU_TES_REL_PREVFLUXOCAIXA = 5
Const MENU_TES_REL_MOVFINRED = 7
Const MENU_TES_REL_MOVFINDET = 8
Const MENU_TES_REL_ERROSCONAUTO = 9
Const MENU_TES_REL_CONCPEND = 10
Const MENU_TES_REL_OUTROS = 12
Const MENU_TES_REL_GERREL = 14

Const MENU_TES_REL_PLANILHAS = 16 'Inserido por Wagner

Const MENU_CP_REL_CAD_FORN = 1
Const MENU_CP_REL_CAD_TIPOFORN = 2
Const MENU_CP_REL_CAD_PORTADORES = 3
Const MENU_CP_REL_CAD_CONDPAGTO = 4

Const MENU_CP_REL_TITPAGAR = 3
Const MENU_CP_REL_POSFORN = 4
Const MENU_CP_REL_CHEQUES = 6
Const MENU_CP_REL_BAIXAS = 7
Const MENU_CP_REL_PAGCANC = 8
Const MENU_CP_REL_OUTROS = 10
Const MENU_CP_REL_GERREL = 12

Const MENU_CP_REL_PLANILHAS = 14 'Inserido por Wagner

Const MENU_CR_REL_CAD_CLI = 1
Const MENU_CR_REL_CAD_TIPOCLI = 2
Const MENU_CR_REL_CAD_CATEGCLI = 3
Const MENU_CR_REL_CAD_COBRADORES = 5
Const MENU_CR_REL_CAD_TIPOSCARTCOB = 6
Const MENU_CR_REL_CAD_PADROESCOB = 7
Const MENU_CR_REL_CAD_TIPOSINSTRCOB = 8
Const MENU_CR_REL_CAD_VENDEDORES = 10
Const MENU_CR_REL_CAD_TIPOSVEND = 11
Const MENU_CR_REL_CAD_REGIOESVEND = 12

Const MENU_CR_REL_TITREC = 3
Const MENU_CR_REL_TITATRASO = 4
Const MENU_CR_REL_POSGERCOB = 5
Const MENU_CR_REL_POSCLI = 6
Const MENU_CR_REL_BAIXAS = 8
Const MENU_CR_REL_COMISSOESVEND = 10
Const MENU_CR_REL_COMISSOESPAG = 11
Const MENU_CR_REL_MAIORESDEV = 13
Const MENU_CR_REL_TITTELEF = 14
Const MENU_CR_REL_TITMALA = 15
Const MENU_CR_REL_OUTROS = 17
Const MENU_CR_REL_GERREL = 19

Const MENU_CR_REL_PLANILHAS = 21 'Inserido por Wagner

Const MENU_EST_REL_MOV_MOVINT = 1
Const MENU_EST_REL_MOV_REQCONSUMO = 2
Const MENU_EST_REL_MOV_KARDEX = 3
Const MENU_EST_REL_MOV_KARDEXDIA = 4
Const MENU_EST_REL_MOV_RESUMOKARDEX = 5
Const MENU_EST_REL_MOV_BOLETIMENT = 6

Const MENU_EST_REL_INV_ETIQINV = 1
Const MENU_EST_REL_INV_INVENTARIO = 2
Const MENU_EST_REL_INV_DEMAPINV = 3
Const MENU_EST_REL_INV_REGINVMOD7 = 4

Const MENU_EST_REL_PRO_OP = 1
Const MENU_EST_REL_PRO_EMPENHOS = 2
Const MENU_EST_REL_PRO_LISTAFALTAS = 3
Const MENU_EST_REL_PRO_MOVESTOP = 4
Const MENU_EST_REL_PRO_RESPRODOP = 5

Const MENU_EST_REL_CAD_FORN = 1
Const MENU_EST_REL_CAD_PRODUTOS = 2
Const MENU_EST_REL_CAD_ALMOXARIFADO = 3
Const MENU_EST_REL_CAD_KITS = 4
Const MENU_EST_REL_CAD_UTILPROD = 5

Const MENU_EST_REL_PRODVEND = 7
Const MENU_EST_REL_RESENTSAIVALOR = 8
Const MENU_EST_REL_CONSUMOVENDAMES = 9
Const MENU_EST_REL_ANALEST = 11
Const MENU_EST_REL_ANALMOVEST = 12
Const MENU_EST_REL_PONTOPEDIDO = 13
Const MENU_EST_REL_SALDOEST = 14
Const MENU_EST_REL_OUTROS = 16
Const MENU_EST_REL_GERREL = 18

Const MENU_EST_REL_PLANILHAS = 20 'Inserido por Wagner

Const MENU_FAT_REL_COMISSOES = 7
Const MENU_FAT_REL_LISTAPRECOS = 8
Const MENU_FAT_REL_OUTROS = 10
Const MENU_FAT_REL_GERREL = 12

Const MENU_FAT_REL_PLANILHAS = 14 'Inserido por Wagner

Const MENU_FAT_REL_CAD_CLI = 1
Const MENU_FAT_REL_CAD_PRODUTOS = 2
Const MENU_FAT_REL_CAD_VENDEDORES = 3


Const MENU_FAT_REL_DOC_PRENOTA = 1
Const MENU_FAT_REL_DOC_NFISC = 2
Const MENU_FAT_REL_DOC_NFISCDEV = 3
Const MENU_FAT_REL_DOC_NFISCTRANSP = 4


Const MENU_FAT_REL_VEN_FATCLI = 1
Const MENU_FAT_REL_VEN_FATCLIPROD = 2
Const MENU_FAT_REL_VEN_FATVEND = 3
Const MENU_FAT_REL_VEN_FATPRAZOPAG = 4
Const MENU_FAT_REL_VEN_FATREALPREV = 5
Const MENU_FAT_REL_VEN_RESVEND = 6
Const MENU_FAT_REL_VEN_DISPESTVENDA = 7

Const MENU_FAT_REL_PED_PEDIDSOAPTOSFAT = 1
Const MENU_FAT_REL_PED_PEDIDOSNENTREGUE = 2
Const MENU_FAT_REL_PED_PEDIDOPROD = 3
Const MENU_FAT_REL_PED_PEDIDOSVENDCLI = 4
Const MENU_FAT_REL_PED_PEDIDOSVENDPROD = 5
Const MENU_FAT_REL_PED_PEDIDOSPRODUCAO = 6
Const MENU_FAT_REL_PED_PEDIDOSCLI = 7

Const MENU_FAT_REL_TRV_ATEND = 1
Const MENU_FAT_REL_TRV_EST = 2
Const MENU_FAT_REL_TRV_DESV = 3
Const MENU_FAT_REL_TRV_VOU = 4
Const MENU_FAT_REL_TRV_ACOM_INAD = 5
Const MENU_FAT_REL_TRV_POS_INAD = 6
Const MENU_FAT_REL_TRV_IMP = 7

Const MENU_CTB_ROT_ATUALIZALOTE = 1
Const MENU_CTB_ROT_APURAPERIODO = 2
Const MENU_CTB_ROT_APURAEXERC = 3
Const MENU_CTB_ROT_FECHAEXERC = 4
Const MENU_CTB_ROT_REPROCESSAMENTO = 5
Const MENU_CTB_ROT_REABREEXERC = 6
Const MENU_CTB_ROT_RATEIOOFF = 7
Const MENU_CTB_ROT_IMPORTACAO = 8
Const MENU_CTB_ROT_GERACAODRE = 9
Const MENU_CTB_ROT_DESAPURAEXERC = 10
Const MENU_CTB_ROT_IMPORTACAORATEIO = 11
Const MENU_CTB_ROT_TRVRATEIO = 12
Const MENU_CTB_ROT_SPED = 13
Const MENU_CTB_ROT_IMPORTLCTOS = 14

Const MENU_CTB_ROT_SPED_DIARIO = 1
Const MENU_CTB_ROT_SPED_FCONT = 2

Const MENU_TES_ROT_LIMPAARQ = 1
Const MENU_TES_ROT_RECEXTRATOCONCILIA = 2
Const MENU_TES_ROT_CONCILIAEXTRATO = 3

'Const MENU_CP_ROT_CONFIGURACAO = 1
'Const MENU_CP_ROT_GERACAOARQICMS = 3
'Const MENU_CP_ROT_LIMPAARQ = 4

Const MENU_TES_CP_REMESSA_PAGTO = 1
Const MENU_TES_CP_ENVIO_EMAIL_COBR_FAT = 2
Const MENU_TES_CP_ENVIO_EMAIL_AVISO_PAGTO = 3
Const MENU_CP_ROT_IMPORTACAOCF_REIZA = 4
Const MENU_CP_ROT_IMPORTACAO_RETPAGTO = 5

Const MENU_CR_ROT_ATUALIZAPAGCOMISSAO = 1
Const MENU_CR_ROT_LIMPAARQ = 2
Const MENU_CR_ROT_EMISSAO_BOLETOS = 4
Const MENU_CR_ROT_EMISSAO_DUPLICATAS = 5
Const MENU_CR_ROT_REAJUSTE_TITULOS = 6
'Const MENU_CR_ROT_ENVIO_EMAIL_COBR = 7
'Const MENU_CR_ROT_ENVIO_EMAIL_AGRADECIMENTO = 8
Const MENU_CR_ROT_IMPORT_TITREC_AF = 9
Const MENU_CR_ROT_EXPORT_ASSOC_AF = 10
Const MENU_CR_ROT_IMPORTFAT_REIZA = 11
Const MENU_CR_ROT_IMPORT_EXT_REDES = 13
Const MENU_CR_ROT_BAIXA_CARTAO = 14

Const MENU_CR_ROT_EMAIL_AVISOCOBR = 1
Const MENU_CR_ROT_EMAIL_COBR = 2
Const MENU_CR_ROT_EMAIL_AGRADECIMENTO = 3

Const MENU_EST_ROT_ATUALIZALOTE = 1
Const MENU_EST_ROT_ATUALIZARASTROLOTE = 2
Const MENU_EST_ROT_CUSTOMEDIOPRODUCAO = 3
Const MENU_EST_ROT_FECHAMES = 4
Const MENU_EST_ROT_CLASSIFICACAOABC = 5
Const MENU_EST_ROT_EMINFISC = 7
Const MENU_EST_ROT_EMINFISCFAT = 8
Const MENU_EST_ROT_EMISSREC = 9
Const MENU_EST_ROT_REPROCESSAMENTO = 11
Const MENU_EST_ROT_IMPORTACAOINV = 12
Const MENU_EST_ROT_IMPORTARNFRAIZ = 13
Const MENU_EST_ROT_IMPORT_XML = 14

Const MENU_FAT_ROT_REAJUSTEPRECO = 1
Const MENU_FAT_ROT_ATUALIZARASTRO = 2
Const MENU_FAT_ROT_ROMANEIODESPACHO = 4
Const MENU_FAT_ROT_EMINFISC = 5
Const MENU_FAT_ROT_EMINFISCFAT = 6
Const MENU_FAT_ROT_EMIFATURAS = 7
Const MENU_FAT_ROT_EMIDUPLICATAS = 8

Const MENU_FAT_ROT_FP_MARGCONTR = 1
Const MENU_FAT_ROT_FP_RATEIOCUSTODIR = 3
Const MENU_FAT_ROT_FP_RATEIOCUSTOFIXO = 4
Const MENU_FAT_ROT_FP_CALCPRECOS = 5
Const MENU_FAT_ROT_FP_AJUSTEPRECO = 6


Const MENU_FAT_ROT_CONTRATOFATLOTE = 9
Const MENU_FAT_ROT_GERARQPV = 12
Const MENU_FAT_ROT_EXPORTAR_NF = 14
Const MENU_FAT_ROT_IMPORTAR_NF = 15
Const MENU_FAT_ROT_EXPORTAR_DADOS = 17
'Const MENU_FAT_ROT_IMPORTAR_DADOS = 18

Const MENU_FAT_ROT_IMPORTAR_DADOS = 1
Const MENU_FAT_ROT_IMPORTAR_DADOS_MANUAL = 2
Const MENU_FAT_ROT_IMPORTAR_DADOS_XML = 3

Const MENU_FAT_ROT_GERARQ_RPS_LOTE = 20
Const MENU_FAT_ROT_IMPORTAR_NFE = 21
Const MENU_FAT_ROT_EXPORTAR_NF_HARMONIA = 23
Const MENU_FAT_ROT_NF_PAULISTA = 24
Const MENU_FAT_ROT_IMPORTAR_PV_SET = 25
Const MENU_FAT_ROT_REGFATHTML = 27
Const MENU_FAT_ROT_TRP_VEND_REMONTA = 29
Const MENU_FAT_ROT_ARQCOMISSOES = 33
Const MENU_FAT_ROT_EXPORTLOREAL = 55
Const MENU_FAT_ROT_EXPORTGENVEND = 56

Const MENU_FAT_ROT_NFEFED_GERALOTEENVIO = 1
Const MENU_FAT_ROT_NFEFED_CONSULTALOTE = 2
Const MENU_FAT_ROT_NFEFED_EMAIL = 3
Const MENU_FAT_ROT_NFEFED_INUTFAIXA = 4
Const MENU_FAT_ROT_NFEFED_CONSULTANFE = 5
Const MENU_FAT_ROT_NFEFED_NFESCAN = 6
Const MENU_FAT_ROT_NFEFED_EXPORTXMLNFE = 7
Const MENU_FAT_ROT_NFEFED_CARTACORRECAO = 8

Const MENU_FAT_ROT_NFSE_GERALOTEENVIO = 1
Const MENU_FAT_ROT_NFSE_CONSULTALOTE = 2

Const MENU_CTB_CONFIG_EXERCICIO = 1
Const MENU_CTB_CONFIG_EXERCICIOFILIAL = 2
Const MENU_CTB_CONFIG_CONFIGURACOES = 3
Const MENU_CTB_CONFIG_SEGMENTOS = 4
Const MENU_CTB_CONFIG_CAMPOSGLOBAIS = 5

Const MENU_TES_CONFIG_CONFIGURACOES = 1

Const MENU_CP_CONFIG_CONFIGURACAO = 1
Const MENU_CP_CONFIG_SEGMENTOS = 2 'Inserido por Wagner

Const MENU_CR_CONFIG_CONFIGURACAO = 1
Const MENU_CR_CONFIG_COBRANCAELETRONICA = 2
Const MENU_CR_CONFIG_SEGMENTOS = 3 'Inserido por Wagner

Const MENU_EST_CONFIG_SEGMENTOS = 1
Const MENU_EST_CONFIG_CONFIGURACAO = 2

Const MENU_FAT_CONFIG_CONFIGURACAO = 1
Const MENU_FAT_CONFIG_AUTORIZACREDITO = 2
Const MENU_FAT_CONFIG_SEGMENTOS = 3
Const MENU_FAT_CONFIG_COMISSOESREGRAS = 4
Const MENU_FAT_CONFIG_OUTRASCONFIG = 7
Const MENU_FAT_CONFIG_EMAILCONFIG = 8
Const MENU_FAT_CONFIG_REGRASMSG = 9

Const MENU_FAT_CONFIG_FP_MNEMONICOS = 1
Const MENU_FAT_CONFIG_FP_PLANILHAS = 2
Const MENU_FAT_CONFIG_FP_MARGCONTR = 3

Const MENU_COM_CONFIG_CONFIGURACAO = 1

Const MENU_COM_ROT_PARAMPONTOPEDIDO = 1

'------------------- CONSTANTE PARA O MENU DO FIS ---------------------

'Cadastros para o Menu do FIS
Const MENU_FIS_CAD_NATUREZAOP = 1
Const MENU_FIS_CAD_TIPOTRIB = 2
Const MENU_FIS_CAD_EXCECOESICMS = 3
Const MENU_FIS_CAD_EXCECOESIPI = 4
Const MENU_FIS_CAD_TRIBUTACAOFORN = 5
Const MENU_FIS_CAD_TRIBUTACAOCLI = 6
Const MENU_FIS_CAD_TIPOREGAPURACAOICMS = 7
Const MENU_FIS_CAD_TIPOREGAPURACAOIPI = 8

'Movimentos para o Menu do FIS
Const MENU_FIS_MOV_REGENTRADA = 1
Const MENU_FIS_MOV_REGSAIDA = 2
Const MENU_FIS_MOV_ICMS_APURACAO = 1
Const MENU_FIS_MOV_ICMS_REGEMITENTES = 3
Const MENU_FIS_MOV_ICMS_REGCADPRODUTOS = 4
Const MENU_FIS_MOV_ICMS_LANCAPURACAO = 5
Const MENU_FIS_MOV_ICMS_GNRICMS = 6
Const MENU_FIS_MOV_ICMS_REGINVENTARIO = 7
Const MENU_FIS_MOV_ICMS_GUIASICMS = 8
Const MENU_FIS_MOV_ICMS_GUIASICMSST = 9
Const MENU_FIS_MOV_IPI_APURACAO = 1
Const MENU_FIS_MOV_IPI_LANCAPURACAO = 2

Const MENU_FIS_MOV_PIS_APURACAO = 1

Const MENU_FIS_MOV_COFINS_APURACAO = 1

'Rotinas para o Menu do FIS
Const MENU_FIS_ROT_FECHAMENTOLIVRO = 1
Const MENU_FIS_ROT_GERACAOARQICMS = 2
Const MENU_FIS_ROT_GERACAOARQIN86 = 5
Const MENU_FIS_ROT_REABERTURALIVRO = 6
Const MENU_FIS_ROT_LIVREGESATUALIZA = 7
Const MENU_FIS_ROT_SPEDFISCAL = 8
Const MENU_FIS_ROT_SPEDFISCALPIS = 9
Const MENU_FIS_ROT_DFC = 10
Const MENU_FIS_ROT_SPEDECF = 11

'Configuração de FIS
Const MENU_FIS_CONFIG_CONFIGURACAO = 1
Const MENU_FIS_CONFIG_TRIBUTO = 2
Const MENU_FIS_CONFIG_ICMS = 3

'Relatórios FIS
Const MENU_FIS_REL_TERMO = 1
Const MENU_FIS_REL_APURACAO_PIS = 7
Const MENU_FIS_REL_APURACAO_COFINS = 8
Const MENU_FIS_REL_IR = 9
Const MENU_FIS_REL_INSS = 10
Const MENU_FIS_REL_RESUMO_DECLAN = 12
Const MENU_FIS_REL_PISCOFINSCSLL = 13
Const MENU_FIS_REL_OUTROS = 15
Const MENU_FIS_REL_GERREL = 17
Const MENU_FIS_REL_LANC_NATOPERACAO = 1
Const MENU_FIS_REL_LANC_ESTADO = 2
Const MENU_FIS_REL_LANC_TIPOICMS = 3
Const MENU_FIS_REL_LANC_CLIENTE = 4
Const MENU_FIS_REL_LANC_FORNECEDOR = 5
Const MENU_FIS_REL_LANC_TODOS = 6
Const MENU_FIS_REL_LIVREG_ENTRADA = 1
Const MENU_FIS_REL_SAIDA = 2
Const MENU_FIS_REL_RCPE = 3
Const MENU_FIS_REL_REGINVENTARIO = 4
Const MENU_FIS_REL_APURACAO_ICMS = 5
Const MENU_FIS_REL_RESUMOICMS = 6
Const MENU_FIS_REL_APURACAO_IPI = 7
Const MENU_FIS_REL_RESUMOIPI = 8
Const MENU_FIS_REL_EMITENTES = 9
Const MENU_FIS_REL_MERCARDORIAS = 10
Const MENU_FIS_REL_OPERINTEREST = 11
Const MENU_FIS_REL_GNRICMS = 12
Const MENU_FIS_REL_APURACAO_ISS = 1

Const MENU_FIS_REL_PLANILHAS = 19 'Inserido por Wagner

'Consultas de FIS
Const MENU_FIS_CON_EXCECOES_ICMS = 1
Const MENU_FIS_CON_EXCECOES_IPI = 2
Const MENU_FIS_CON_NATUREZA_OPERACAO = 3
Const MENU_FIS_CON_PRODUTOS = 4
Const MENU_FIS_CON_TIPO_TRIBUTACAO = 5
Const MENU_FIS_CON_TIPO_REG_APUR_ICMS = 6
Const MENU_FIS_CON_TIPO_REG_APUR_IPI = 7

Const MENU_FIS_CON_LIVRO_ABERTOS = 9
Const MENU_FIS_CON_LIVROS_FECHADOS = 10

Const MENU_FIS_CON_APUR_ICMS = 12
Const MENU_FIS_CON_APUR_IPI = 13
Const MENU_FIS_CON_LANC_APUR_ICMS = 14
Const MENU_FIS_CON_LANC_APUR_IPI = 15
Const MENU_FIS_CON_REG_ENTRADA = 16
Const MENU_FIS_CON_REG_SAIDA = 17
Const MENU_FIS_CON_ITEM_NF = 19
Const MENU_FIS_CON_ITEM_NF_TRIB = 20

'Relatórios Compras
Const MENU_COM_REL_CONCABERT = 7
Const MENU_COM_REL_ANCOTREC = 8
Const MENU_COM_REL_PREVENTPROD = 10
Const MENU_COM_REL_AGPREVENT = 11
Const MENU_COM_REL_OUTROS = 13
Const MENU_COM_REL_GERREL = 15

Const MENU_COM_REL_PLANILHAS = 17 'Inserido por Wagner

Const MENU_COM_REL_CAD_REQ = 1
Const MENU_COM_REL_CAD_COMP = 2
Const MENU_COM_REL_CAD_FORN = 3

Const MENU_COM_REL_PC_ABERTO = 1
Const MENU_COM_REL_PC_BAIXADOS = 2
Const MENU_COM_REL_PC_ATRASADOS = 3
Const MENU_COM_REL_PC_BLOQUEADOS = 4
Const MENU_COM_REL_PC_EMISSAO = 5
Const MENU_COM_REL_PC_EMISSCONC = 6
Const MENU_COM_REL_PC_NF = 7

Const MENU_COM_REL_PCOT_EMISSAO = 1
Const MENU_COM_REL_PCOT_EMISSGER = 2

Const MENU_COM_REL_RC_ABERTO = 1
Const MENU_COM_REL_RC_BAIXADOS = 2
Const MENU_COM_REL_RC_ATRASADOS = 3
Const MENU_COM_REL_RC_PV = 4
Const MENU_COM_REL_RC_OP = 5

Const MENU_COM_REL_PROD_COT = 1
Const MENU_COM_REL_PROD_FORN = 2
Const MENU_COM_REL_PROD_COMPRAS = 3
Const MENU_COM_REL_PROD_REQ = 4
Const MENU_COM_REL_PROD_PC = 5

'*** MENU CADASTROS LOJA *** Luiz Nogueira 31/03/04
Const MENU_LJ_CAD_CLIENTE_LOJA = 1
Const MENU_LJ_CAD_PROD = 2
Const MENU_LJ_CAD_OPERADOR = 3
Const MENU_LJ_CAD_CAIXA = 4
Const MENU_LJ_CAD_VENDEDOR = 5
Const MENU_LJ_CAD_ECF = 6

Const MENU_LJ_CAD_TA_TABPRECO = 1
Const MENU_LJ_CAD_TA_PRODDESC = 2
Const MENU_LJ_CAD_TA_ADM_MEIO_PAG = 3
Const MENU_LJ_CAD_TA_REDE = 4
Const MENU_LJ_CAD_TA_TECLADO = 5
Const MENU_LJ_CAD_TA_IMPRESSORA = 6
'****************************

Const MENU_LJ_MOV_BORD_BOLETO = 1
Const MENU_LJ_MOV_BORD_CHEQUE = 2
Const MENU_LJ_MOV_BORD_VALE_TICKET = 3
Const MENU_LJ_MOV_BORDEROOUTROS = 4
Const MENU_LJ_MOV_DEPOSITO_BANCARIO = 6
Const MENU_LJ_MOV_DEPOSITO_CAIXA = 7
Const MENU_LJ_MOV_SAQUE_CAIXA = 8
Const MENU_LJ_MOV_CHEQUENESP = 10
Const MENU_LJ_MOV_TRANSFERENCIA = 11
Const MENU_LJ_MOV_RECEB_CARNE = 12


Const MENU_LJ_ROT_GERACAOARQCC = 1
Const MENU_LJ_ROT_GERACAOARQBACK = 2
Const MENU_LJ_ROT_CARGABALANCA = 3

Const MENU_LJ_CONFIG_CONFIGURACAO = 1
Const MENU_LJ_CONFIG_TECLADO = 2
Const MENU_LJ_CONFIG_VENDEDOR_LOJA = 3


Const MENU_LJ_CON_ADMMEIOPAGTO = 1
Const MENU_LJ_CON_MOVCAIXA = 3
Const MENU_LJ_CON_MOVCAIXA_SESSOES = 4
Const MENU_LJ_CON_MOVCAIXACF = 5
Const MENU_LJ_CON_CUPOMFISCAL = 7
Const MENU_LJ_CON_ITEMCUPOMFISCAL = 8


'*** MENU CONSULTAS LOJA *** Luiz Nogueira 07/04/04
Const MENU_LJ_CON_CAD_CLIENTES = 1
Const MENU_LJ_CON_CAD_PRODUTOS = 2
Const MENU_LJ_CON_CAD_OPERADOR = 3
Const MENU_LJ_CON_CAD_CAIXA = 4
Const MENU_LJ_CON_CAD_VENDEDOR = 5
Const MENU_LJ_CON_CAD_ECF = 6
Const MENU_LJ_CON_CAD_PRECOS = 8
Const MENU_LJ_CON_CAD_MEIOSPAGTO = 9
Const MENU_LJ_CON_CAD_REDES = 10
'Const MENU_LJ_CON_CAD_CUPOMFISCAL = 11
'Const MENU_LJ_CON_CAD_ITEMCUPOMFISCAL = 12
'*******************************************************

'*** MENU RELATÓRIOS LOJA *** Luiz Nogueira 31/03/04
Const MENU_LJ_REL_OUTROS = 6
Const MENU_LJ_REL_GERREL = 8

Const MENU_LJ_REL_PLANILHAS = 10 'Inserido por Wagner

Const MENU_LJ_REL_CAD_CLI = 1
Const MENU_LJ_REL_CAD_PRODUTOS = 2
Const MENU_LJ_REL_CAD_OPERADORES = 3
Const MENU_LJ_REL_CAD_CAIXAS = 4
Const MENU_LJ_REL_CAD_ECFS = 5

Const MENU_LJ_REL_CAIXA_MOVCAIXAS = 1
Const MENU_LJ_REL_CAIXA_PAINELCAIXAS = 2
Const MENU_LJ_REL_CAIXA_ORCAMENTOSLOJA = 3
Const MENU_LJ_REL_CAIXA_CUPONSFISCAIS = 4

Const MENU_LJ_REL_VEN_EVOLVENDAS = 1
Const MENU_LJ_REL_VEN_MAPAVENDAS = 2
Const MENU_LJ_REL_VEN_FLASHVENDAS = 3
Const MENU_LJ_REL_VEN_VENDASXMEIOPAGTO = 4
Const MENU_LJ_REL_VEN_RANKINGPRODUTOS = 5
Const MENU_LJ_REL_VEN_PRODDEVTROCA = 6

Const MENU_LJ_REL_BORD_BORDEROBOLETO = 1
Const MENU_LJ_REL_BORD_BORDEROTICKET = 2
Const MENU_LJ_REL_BORD_BORDEROOUTROS = 3
'*******************************************************

Const MENU_PCP_MOV_MANUAL = 1
Const MENU_PCP_MOV_AUTOMATICO = 3
Const MENU_PCP_MOV_EMPENHO = 4
Const MENU_PCP_MOV_ORDPRODBLOQUEIO = 5
Const MENU_PCP_MOV_REQPROD = 6
Const MENU_PCP_MOV_TRANSF = 7
Const MENU_PCP_MOV_PRODENT = 9
Const MENU_PCP_MOV_APONTAMENTOPRODUCAO = 11
Const MENU_PCP_MOV_PMP = 12
Const MENU_PCP_MOV_ORDEMCORTE = 14
Const MENU_PCP_MOV_REQPROD_DPACK = 15
Const MENU_PCP_MOV_OCARTX = 18
Const MENU_PCP_MOV_OCMANUALARTX = 19

Const MENU_PCP_CON_CAD_PRODUTOS = 1
Const MENU_PCP_CON_CAD_ALMOXARIFADO = 2
Const MENU_PCP_CON_CAD_KIT = 3
Const MENU_PCP_CON_CAD_TIPOPROD = 4
Const MENU_PCP_CON_CAD_CLASSEUM = 5
Const MENU_PCP_CON_CAD_COMPETENCIA = 6
Const MENU_PCP_CON_CAD_MAQUINA = 7
Const MENU_PCP_CON_CAD_CT = 8
Const MENU_PCP_CON_CAD_TIPOMAODEOBRA = 9

Const MENU_PCP_CON_PRO_OP = 1
Const MENU_PCP_CON_PRO_OP_BAIXA = 2
Const MENU_PCP_CON_PRO_EMPENHO = 3
Const MENU_PCP_CON_PRO_ITENSOP = 4
Const MENU_PCP_CON_PRO_REQPROD = 5
Const MENU_PCP_CON_PRO_PRODUCAO = 6
Const MENU_PCP_CON_PRO_REQPRODOP = 7
Const MENU_PCP_CON_PRO_PRODUCAOOP = 8
Const MENU_PCP_CON_PRO_ROTEIRO = 9
Const MENU_PCP_CON_PRO_TAXA = 10
Const MENU_PCP_CON_PRO_PMP = 11

Const MENU_PCP_CON_EST_ESTPROD = 1
Const MENU_PCP_CON_EST_ESTPRODFILIAL = 2
Const MENU_PCP_CON_EST_ESTPRODTERC = 3

Const MENU_PCP_CON_RASTRO_LOTES = 1
Const MENU_PCP_CON_RASTRO_SALDOS = 2
Const MENU_PCP_CON_RASTRO_MOVIMENTOS = 3

Const MENU_PCP_CON_CUR_CERTIFICADO = 1
Const MENU_PCP_CON_CUR_CURSO = 2
Const MENU_PCP_CON_CUR_CURSOMO = 3
Const MENU_PCP_CON_CUR_CERTIFICADOMO = 4

Const MENU_PCP_REL_OP = 3
Const MENU_PCP_REL_EMPENHOS = 4
Const MENU_PCP_REL_LISTAFALTAS = 5
Const MENU_PCP_REL_DISTREATOR = 7
Const MENU_PCP_REL_PRODXOP = 8
Const MENU_PCP_REL_MOVESTOP = 9
Const MENU_PCP_REL_OPXREQPROD = 10
Const MENU_PCP_REL_ANALRENDOP = 12
Const MENU_PCP_REL_PVENDAXPCONSUMO = 13
Const MENU_PCP_REL_FORPADRAOCUSTO = 14
Const MENU_PCP_REL_OUTROS = 16
Const MENU_PCP_REL_GERREL = 18

Const MENU_PCP_REL_PLANILHAS = 20 'Inserido por Wagner


Const MENU_PCP_REL_CAD_PRODUTOS = 1
Const MENU_PCP_REL_CAD_ALMOXARIFADO = 2
Const MENU_PCP_REL_CAD_KITS = 3
Const MENU_PCP_REL_CAD_UTILPROD = 4

Const MENU_PCP_CAD_PRODUTOS = 1
Const MENU_PCP_CAD_ALMOXARIFADOS = 2
Const MENU_PCP_CAD_KIT = 3
Const MENU_PCP_CAD_LOTERASTRO = 4
Const MENU_PCP_CAD_PRODALM = 5
Const MENU_PCP_CAD_EMBALAGEM = 6
Const MENU_PCP_CAD_MAQUINA = 7

Const MENU_PCP_CAD_TA_ESTOQUE = 1
Const MENU_PCP_CAD_TA_TIPOPROD = 2
Const MENU_PCP_CAD_TA_UNIDADEMED = 3
Const MENU_PCP_CAD_TA_ESTOQUEINI = 4
Const MENU_PCP_CAD_TA_CATEGORIAPROD = 5
Const MENU_PCP_CAD_TA_PRODUTOEMBALAGEM = 6
Const MENU_PCP_CAD_TA_TESTESQUALIDADE = 7
Const MENU_PCP_CAD_TA_COMPETENCIAS = 8
Const MENU_PCP_CAD_TA_CENTROSDETRABALHOS = 9
Const MENU_PCP_CAD_TA_TAXADEPRODUCAO = 10
Const MENU_PCP_CAD_TA_ROTEIROSDEFABRICACAO = 11
Const MENU_PCP_CAD_TA_TIPOSDEMAODEOBRA = 13
Const MENU_PCP_CAD_TA_USUPRODARTX = 18
Const MENU_PCP_CAD_TA_CERTIFICADOS = 19
Const MENU_PCP_CAD_TA_CURSOS = 20

Const MENU_PCP_ROT_CUSTOMEDIOPRODUCAO = 1
Const MENU_PCP_ROT_MRP = 2

Const MENU_PCP_CONFIG_SEGMENTOS = 1
Const MENU_PCP_CONFIG_CONFIGURACAO = 2

'Incluído por Luiz Nogueira em 13/01/04
Const MENU_CRM_CAD_CLIENTES = 1
Const MENU_CRM_CAD_ATENDENTES = 2
Const MENU_CRM_CAD_CAMPOSGENERICOS = 3
Const MENU_CRM_CAD_VENDEDORES = 4
Const MENU_CRM_CADTA_CATCLIENTES = 1
Const MENU_CRM_CADTA_TIPOSCLIENTES = 2
Const MENU_CRM_CADTA_CLIENTECONTATOS = 3
Const MENU_CRM_CADTA_TIPOSVENDEDORES = 4
Const MENU_CRM_CADTA_REGIOESVENDA = 5
Const MENU_CRM_CADTA_CONTATOS = 6
Const MENU_CRM_CADTA_CLIENTEFCONTATOS = 7
Const MENU_CRM_CADTA_CONTATOCLIPOREMAIL = 8

Const MENU_CRM_MOV_RELACCLI = 1
Const MENU_CRM_MOV_RELACCLICONS = 2
Const MENU_CRM_MOV_RELACCON = 3
Const MENU_CRM_MOV_RELACCONCONS = 4
Const MENU_CRM_MOV_CONTATOCLI = 5

Const MENU_CRM_CONCAD_CLIENTES = 1
Const MENU_CRM_CONCAD_ATENDENTES = 2
Const MENU_CRM_CONCAD_VENDEDORES = 3
Const MENU_CRM_CONCAD_CLICONTATOS = 4
Const MENU_CRM_CONCAD_CLICONTATOSTEL = 5

Const MENU_CRM_CON_CLIENTECONS = 2
Const MENU_CRM_CONRELAC_PENDENTES = 1
Const MENU_CRM_CONRELAC_ENCERRADOS = 2
Const MENU_CRM_CONRELAC_TODOS = 3
Const MENU_CRM_CONRELAC_CALLCENTER = 4
Const MENU_CRM_CONRELAC_SOLSRV = 5

Const MENU_CRM_REL_FOLLOWUP = 3
Const MENU_CRM_REL_RELACESTAT = 4
Const MENU_CRM_REL_CLIENTESSEMRELAC = 5
Const MENU_CRM_REL_OUTROS = 7
Const MENU_CRM_REL_GERREL = 9
Const MENU_CRM_RELCAD_CLI = 1
Const MENU_CRM_RELCAD_ATENDENTES = 2
Const MENU_CRM_RELCAD_VENDEDORES = 3
Const MENU_CRM_RELCAD_CLIENTECONTATOS = 4

Const MENU_CRM_REL_PLANILHAS = 11 'Inserido por Wagner

Const MENU_CRM_ROT_ENVIOEMAILCLI = 1

Const MENU_SRV_CON_CAD_PRODUTOS = 1
Const MENU_SRV_CON_CAD_GARANTIA = 2
Const MENU_SRV_CON_CAD_TIPOGARANTIA = 3
Const MENU_SRV_CON_CAD_CONTRATOSRV = 4
Const MENU_SRV_CON_CAD_ITENSCONTRATOSRV = 5
Const MENU_SRV_CON_CAD_MAODEOBRA = 6
Const MENU_SRV_CON_CAD_MAQUINAS = 7
Const MENU_SRV_CON_CAD_COMPETENCIA = 8
Const MENU_SRV_CON_CAD_CT = 9

Const MENU_SRV_CON_SOLIC_ABERTA = 1
Const MENU_SRV_CON_SOLIC_BAIXADA = 2
Const MENU_SRV_CON_SOLIC_TODAS = 3
Const MENU_SRV_CON_CRM = 4
Const MENU_SRV_CON_BI_SS = 5

Const MENU_SRV_CON_OS_ABERTA = 1
Const MENU_SRV_CON_OS_BAIXADA = 2
Const MENU_SRV_CON_OS_TODAS = 3

Const MENU_SRV_CON_PED_ABERTOS = 1
Const MENU_SRV_CON_PED_BAIXADOS = 2
Const MENU_SRV_CON_PED_TODOS = 3
Const MENU_SRV_CON_BI_PS = 4

Const MENU_SRV_CON_ORCAMENTOS = 3
Const MENU_SRV_CON_PEDIDOS = 4
Const MENU_SRV_CON_ITENSPEDIDOS = 5
Const MENU_SRV_CON_OS = 6
Const MENU_SRV_CON_ITEMOS = 7
Const MENU_SRV_CON_NFSRV = 8
Const MENU_SRV_CON_ITEMNFSRV = 8

Const MENU_SRV_MOV_SOLICSRV = 1

Const MENU_SRV_MOV_ORCAMENTOSRV = 10
Const MENU_SRV_MOV_LIBORCSRV = 11

Const MENU_SRV_MOV_PEDIDOSRV = 20
Const MENU_SRV_MOV_LIBPEDSRV = 21
Const MENU_SRV_MOV_BAIXAPEDSRV = 22

Const MENU_SRV_MOV_OS = 30
Const MENU_SRV_MOV_OSAPONT = 31
Const MENU_SRV_MOV_MOVEST = 32

Const MENU_SRV_MOV_NFSRV = 40
Const MENU_SRV_MOV_NFFATSRV = 41
Const MENU_SRV_MOV_NFPEDIDOSRV = 42
Const MENU_SRV_MOV_NFFATPEDIDOSRV = 43
Const MENU_SRV_MOV_NFFATGARANTIASRV = 46

'Const MENU_SRV_MOV_ACOMPANHAMENTOSRV = 7

Const MENU_SRV_CAD_CLIENTE = 1
Const MENU_SRV_CAD_PRODUTO = 2

Const MENU_SRV_CAD_TIPOGARANTIA = 10
Const MENU_SRV_CAD_GARANTIA = 11

Const MENU_SRV_CAD_CONTRATO = 20
Const MENU_SRV_CAD_CONTRATOSRV = 21

Const MENU_SRV_CAD_TIPOMAODEOBRA = 30
Const MENU_SRV_CAD_MAODEOBRA = 31

Const MENU_SRV_CAD_MAQUINA = 40
Const MENU_SRV_CAD_COMPETENCIA = 41
Const MENU_SRV_CAD_CENTROTRABALHO = 42
Const MENU_SRV_CAD_ROTEIRO = 43

Const MENU_SRV_REL_SOLIC = 1
Const MENU_SRV_REL_OS = 3

Const MENU_SRV_REL_PS_APTOSFAT = 1
Const MENU_SRV_REL_PS_NAOENTR = 2
Const MENU_SRV_REL_PS_PORSRV = 3
Const MENU_SRV_REL_PS_PORCLI = 4

Const MENU_SRV_CONFIG_CONFIGURACAO = 1

Const MENU_ARQ_FILIAL = 1
Const MENU_ARQ_DATA = 2
Const MENU_ARQ_FERIADO = 3
Const MENU_ARQ_COTACAOMOEDA = 4
Const MENU_ARQ_CADASTROMOEDA = 5
Const MENU_ARQ_IMPRESSORA = 7
Const MENU_ARQ_EDICAO = 9
Const MENU_ARQ_ADMIN = 11
Const MENU_ARQ_ECF = 13
Const MENU_ARQ_WORKFLOW = 15
Const MENU_ARQ_OUTLOOK = 16
Const MENU_ARQ_SAIR = 18

Const MENU_ARQ_OC_IMPRESSORA = 1
Const MENU_ARQ_OC_BACKUP = 2
Const MENU_ARQ_OC_LOGO = 3
Const MENU_ARQ_OC_TELA = 4
Const MENU_ARQ_OC_ARQUIVAMENTO = 5

Const MENU_AJUDA_INDICE = 1
Const MENU_AJUDA_SOBRE = 2
Const MENU_AJUDA_SUPORTE = 4
Const MENU_AJUDA_ATUALIZACAO = 5

Private Sub BotaoAvisos_Click()
    Call Chama_Tela("AvisosInternos")
    BotaoAvisos.BackColor = &H8000000C
    BotaoAvisos.ToolTipText = "Existem 0 avisos novos"
End Sub

Private Sub BotaoConsulta_Click()

    Call Seta_Click(BOTAO_CONSULTA)

End Sub

Private Sub BotaoData_Click()

Dim objData As New AdmGenerico
Dim lErro As Long

On Error GoTo Erro_BotaoData_Click

    objData.vVariavel = CDate(BotaoData.Caption)

    'Chama a tela Calendário
    Call Chama_Tela("Calendario", objData)

    If gdtDataAtual <> objData.vVariavel Then
        
    If SGE_DATA_SIMULADO <> DATA_NULA And CDate(objData.vVariavel) > SGE_DATA_SIMULADO Then
        MsgBox ("para este teste informe uma data até " & CStr(SGE_DATA_SIMULADO))
    Else
        gdtDataAtual = objData.vVariavel
        BotaoData.Caption = Format(objData.vVariavel, "dd/mm/yyyy")
    
        'reinicializa o periodo e exercicio atual se o sistema usar o modulo de contabilidade
        lErro = CTB_Inicializa_Periodo_ExercicioAtual(gdtDataAtual)
        If lErro <> SUCESSO Then Error 44656
    End If

    End If
    
    Exit Sub
    
Erro_BotaoData_Click:

    Select Case Err

        Case 44656

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165225)

    End Select

    Exit Sub


End Sub

Private Sub BotaoFacebook_Click()

Dim sDummy As String
Dim sBrowserExec As String
Dim lRetVal As Long
Dim sURL As String

On Error GoTo Erro_BotaoFacebook_Click

    sBrowserExec = Space(255)
    
    'Aproveita o htm com o Javascript que abre o chat no tamanho certo para verificar o navegador padrão do usuário
    lRetVal = FindExecutable(App.Path & "\suporteonline.htm", sDummy, sBrowserExec)
    sBrowserExec = Trim(sBrowserExec)
    
    'Se não achou o navegador dá erro
    If lRetVal <= 32 Or IsEmpty(sBrowserExec) Then gError 209502
    
    sURL = "http://www.facebook.com.br/corporator.com.br"

    'Abre o htm que redireciona para o chat
    lRetVal = ShellExecute(Me.hWnd, "open", sBrowserExec, sURL, sDummy, SW_NORMAL)
      
    Exit Sub
      
Erro_BotaoFacebook_Click:

    Select Case gErr
    
        Case 209502
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_NAVEGADOR_PADRAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209503)
        
    End Select

    Exit Sub

End Sub

Private Sub BotaoGeraCodigoLight_Click()

    FormGeraCodigoLight.Show
    
End Sub

Private Sub BotaoNumLock_Click()
   NumLock = Not NumLock
   Call TestKeys
End Sub

Private Sub BotaoTestaInt_Click()
'    Dim lNumIntNF As Long
    
'Dim dtDataInicial As Date, dtDataFinal As Date, sProdutoInicial As String, sProdutoFinal As String
'???

 '   Call GV_Exporta("c:\lixo\gv.txt", DATA_NULA, DATA_NULA)

    'Call CreateObject("Rotinasfis.classfisgrava").LivRegESLinha_RefazerNFs(1, CDate("30/03/2011"), CDate("30/03/2011"), False)
    
'dtDataInicial = CDate("01/01/2009")
'dtDataFinal = CDate("31/12/2011")
'sProdutoInicial = "003000001999"
'sProdutoFinal = "003000001999"
'
'    Call CreateObject("RotinasEST.ClassESTGrava").Integridade_movsEst_EstoqueProduto
'    Call CreateObject("rotinasfat.classfatgrava").Correcao_FATEST_Integridades(True, True, dtDataInicial, dtDataFinal, sProdutoInicial, sProdutoFinal)
''    Call CreateObject("rotinasfattrv.classrotimpcoinfo").Coinfo_ImportarDados

 '   Call CF("NFe_PAF_ImportarXml", "\\asp42\f$\Dados\Demo\XML\33140473841488000153551010000000101185650100-nfe.xml", giFilialEmpresa, lNumIntNF)
  MsgBox (CStr(SQL_Comandos_Abertos))
  
End Sub

Private Sub ComboModulo_Click()

Dim lErro As Long
Dim iModuloAnterior As Integer

On Error GoTo Erro_ComboModulo_Click

    If ComboModulo.ListIndex = -1 Then Exit Sub

    iModuloAnterior = ComboModulo.ItemData(ComboModulo.ListIndex)

    'Verifica se houve mudança de Módulo
    If giModuloAnterior <> iModuloAnterior Then

        lErro = Mostra_Menu(ComboModulo.ListIndex)
        If lErro <> SUCESSO Then Error 43637

        giModuloAnterior = ComboModulo.ItemData(ComboModulo.ListIndex)
    End If
    
    If Len(Trim(ComboModulo.Text)) > 0 Then gsModulo = ComboModulo.Text

    Exit Sub

Erro_ComboModulo_Click:

    Select Case Err

        Case 43637

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165226)

    End Select

    Exit Sub

End Sub

Private Sub MDIForm_Load()

Dim lErro As Long
Dim sTexto As String
Dim lTeste As Long
Dim vTeste As Variant
Dim sBuffer As String
Dim iIndice As Integer
Dim dtData As Date
Dim colMenuItens As New Collection
Dim iOrdemStrCmp As Integer
Dim colCampos As Collection
Dim objUsuarios As New ClassUsuarios
Dim iBKPAtivo As Integer
Dim objPrinc As Object
Dim lngLen As Long, lngX As Long
Dim strCompName As String
Dim iNumAvisosNovos As Integer, iForcaAberturaTela As Integer

On Error GoTo Erro_MDIForm_Load
    
    'Carrega referencia à tela principal necessária
    'para preencher a combo Indice (índices de tela)
    'Inicializa outras properties de AdmSeta
    Set objAdmSeta.gobj_ST_TelaPrincipal = Me
    Set objAdmSeta.gobj_ST_TelaAtiva = Nothing
    objAdmSeta.gs_ST_TelaTabela = ""
    objAdmSeta.gs_ST_TelaIndice = ""
    objAdmSeta.gs_ST_TelaSetaClick = ""
    objAdmSeta.gl_ST_ComandoSeta = 0
    objAdmSeta.gi_ST_SetaIgnoraClick = 1
    
    'Guarda ordem na comparacao das setas (posicionamento de cursor)
    Call Conexao_ObterTipoOrdenacaoInt(GL_lConexao, iOrdemStrCmp)
    'Condensamos as ordens maior que 1 em 1
    If iOrdemStrCmp > 1 Then iOrdemStrCmp = 1
    
    objAdmSeta.gi_ST_Ordem_StrCmp = iOrdemStrCmp
    
    ComboModulo.Clear
    
    'Carrega os Módulos
    lErro = Carrega_ComboModulo()
    If lErro <> SUCESSO Then Error 43629

'    Call Jones_Gera_TabConfere_MenuItens

    'Habilita item de menu "Geração de Senha" apenas para Forprint
    lErro = Habilita_Geracao_Senha()
    If lErro <> SUCESSO Then Error 25960

    If gbVPN Then
    
        'esconde opcoes de menu de telas inexistentes
        lErro = Habilita_Itens_Menu_VPN
        If lErro <> SUCESSO Then Error 43638
    
    Else
    
        'esconde os itens de menus à que o usuario nao tem acesso
        lErro = Habilita_Itens_Menu(colMenuItens)
        If lErro <> SUCESSO Then Error 43638
    
        lErro = Habilita_Separadores(colMenuItens)
        If lErro <> SUCESSO Then Error 44635

    
    End If
    
    lErro = Mostra_Menu(NENHUM_MODULO)
    If lErro <> SUCESSO Then Error 43635
    
    lErro = Monta_Botoes_Filiais
    If lErro <> SUCESSO Then Error 43635
   
    If SGE_DATA_SIMULADO = DATA_NULA Then
        gdtDataHoje = Date
    Else
        gdtDataHoje = SGE_DATA_SIMULADO
    End If
    gdtDataAtual = gdtDataHoje

'    Set gobjCheckboxChecked = LoadPicture("checkboxchecked.bmp")
'    Set gobjCheckboxUnchecked = LoadPicture("checkboxunchecked.bmp")
'    Set gobjOptionButtonChecked = LoadPicture("optionbuttonchecked.bmp")
'    Set gobjOptionButtonUnChecked = LoadPicture("optionbuttonunchecked.bmp")
'    Set gobjButton = LoadPicture("botao.bmp")

    'retirado para otimizar o tempo de carga do sistema
    'le todos os registros da tabela Campos e coloca-os na coleção gcolCampos
'    lErro = CF("Campos_Le_Todos",colCampos)
'    If lErro <> SUCESSO Then Error 55988
'
'    Set gcolCampos = colCampos

    Set GL_objMDIForm = Me
    
    dtData = gdtDataAtual
    
    'inicializa o periodo e exercicio atual se o sistema usar o modulo de contabilidade
    lErro = CTB_Inicializa_Periodo_ExercicioAtual(dtData)
    If lErro <> SUCESSO Then Error 44657

    gdtDataAtual = dtData

    BotaoData.Caption = Format(gdtDataAtual, "dd/mm/yyyy")

    Me.Caption = TITULO_TELA_PRINCIPAL & " - " & gsNomeEmpresa & " - " & gsNomeFilialEmpresa
 
    lErro = Keep_Rotinas_Alive
    If lErro <> SUCESSO Then Error 59351
    
    'Inicializa a conta de cheque Pre da filial, se ela ainda não existir e se não estiver na empresa toda
    If giFilialEmpresa <> EMPRESA_TODA And giFilialEmpresa <> Abs(giFilialAuxiliar) Then
        
        lErro = CF("CPR_Inicializa_CCI_ChequePre")
        If lErro <> SUCESSO Then Error 20847
    
    End If
    
    Call CF("sCombo_Seleciona2", ComboModulo, gsModulo)
    
    objUsuarios.sCodUsuario = gsUsuario

    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO Then gError 20847

    If objUsuarios.iWorkFlowAtivo = WORKFLOW_ATIVO Then
    
        Call Chama_Tela("AvisoWFW", 1)
    
        Timer1.Interval = 60000
        
    End If

    iCountBkp = 0
    giExeBkp = DESMARCADO
    lErro = CF("Backup_Verifica_Habilitado", iBKPAtivo)
    If lErro <> SUCESSO Then gError 20847
    
    If iBKPAtivo = MARCADO Then
        Timer2.Interval = 60000
    End If
    
    Call Impressoras_DesabilitaSuporteBiDirecional
    
    lErro = CF("Executa_Rotinas_Inicializacao")
    If lErro <> SUCESSO Then Error 20847
    
    If giTelaTamanhoVariavel = 1 Then
    
        mnuArqSubOC(MENU_ARQ_OC_TELA).Checked = True
        
    Else
        
        mnuArqSubOC(MENU_ARQ_OC_TELA).Checked = False
        
    End If
    
    Set objPrinc = Me
    lErro = CF("MenuItens_Trata_NomeExibicao", objPrinc)
    If lErro <> SUCESSO Then Error 43635
    
    lngLen = 255
    strCompName = String$(lngLen, 0)
    lngX = GetComputerName(strCompName, lngLen)
    If lngX <> 0 Then
        If UCase(left(strCompName, 3)) = "ASP" Then BotaoFacebook.Visible = False
    End If
    
    lErro = CF("Avisos_Obtem_Status", iNumAvisosNovos, iForcaAberturaTela)
    If lErro <> SUCESSO Then Error 43635
    
    If iNumAvisosNovos > 0 Then
        BotaoAvisos.BackColor = vbYellow
        If iNumAvisosNovos > 1 Then
            BotaoAvisos.ToolTipText = Replace(BotaoAvisos.ToolTipText, "0", CStr(iNumAvisosNovos))
        Else
            BotaoAvisos.ToolTipText = "Existe 1 aviso novo"
        End If
    End If
    If iForcaAberturaTela = MARCADO Then Call BotaoAvisos_Click
    
    lErro = CF("EstoqueMes_Trata_Abertura")
    If lErro <> SUCESSO Then Error 43635
    
    gbPreLoadGravar = False
        
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_MDIForm_Load:

    lErro_Chama_Tela = Err
    
    Select Case Err

        Case 43629, 43635, 43638, 44635, 44657, 44984, 55988, 59351, 20847

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165227)

    End Select

    Exit Sub

End Sub

Private Function Mostra_Menu(iIndice As Integer) As Long

Dim lErro As Long
Dim iIndice1 As Integer

On Error GoTo Erro_Mostra_Menu

    'Desabilita todos os Módulos
    For iIndice1 = 1 To NUM_MODULO
        mnuCadastros.Item(iIndice1).Visible = False
        mnuConsultas.Item(iIndice1).Visible = False
        mnuMovimentos.Item(iIndice1).Visible = False
        mnuRotinas.Item(iIndice1).Visible = False
        mnuRelatorios.Item(iIndice1).Visible = False
        mnuConfiguracoes.Item(iIndice1).Visible = False
    Next

    If iIndice <> NENHUM_MODULO Then
        If mnuCadastros.Item(ComboModulo.ItemData(ComboModulo.ListIndex)).Enabled = True Then mnuCadastros.Item(ComboModulo.ItemData(ComboModulo.ListIndex)).Visible = True
        If mnuConsultas.Item(ComboModulo.ItemData(ComboModulo.ListIndex)).Enabled = True Then mnuConsultas.Item(ComboModulo.ItemData(ComboModulo.ListIndex)).Visible = True
        If mnuMovimentos.Item(ComboModulo.ItemData(ComboModulo.ListIndex)).Enabled = True Then mnuMovimentos.Item(ComboModulo.ItemData(ComboModulo.ListIndex)).Visible = True
        If mnuRotinas.Item(ComboModulo.ItemData(ComboModulo.ListIndex)).Enabled = True Then mnuRotinas.Item(ComboModulo.ItemData(ComboModulo.ListIndex)).Visible = True
        If mnuRelatorios.Item(ComboModulo.ItemData(ComboModulo.ListIndex)).Enabled = True Then mnuRelatorios.Item(ComboModulo.ItemData(ComboModulo.ListIndex)).Visible = True
        If mnuConfiguracoes.Item(ComboModulo.ItemData(ComboModulo.ListIndex)).Enabled = True Then mnuConfiguracoes.Item(ComboModulo.ItemData(ComboModulo.ListIndex)).Visible = True
    Else
        giModuloAnterior = 0
    End If

    Mostra_Menu = SUCESSO

    Exit Function

Erro_Mostra_Menu:

    Mostra_Menu = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165228)

    End Select

    Exit Function

End Function

Private Function Habilita_Geracao_Senha() As Long

Dim lErro As Long
Dim objObjeto As Object
Dim objMenuItem As New ClassMenuItens
Dim objDicConfig As New ClassDicConfig
Dim tDicConfig As typeDicConfig
Dim sTextoSenha As String
Dim sCgc As String, sNomeEmpresa As String
Dim colModulosLib As New Collection

On Error GoTo Erro_Habilita_Geracao_Senha
    
    'Lê todos os itens de menu
    lErro = CF("MenuItem_Le_Titulo", "Geração de Senha", objMenuItem)
    If lErro <> SUCESSO Then Error 25965

    lErro = DicConfig_Le(objDicConfig)
    If lErro <> SUCESSO Then Error 25966

    lErro = Senha_Empresa_Decifra(objDicConfig.sSenha, sCgc, sNomeEmpresa, tDicConfig.iLimiteLogs, tDicConfig.iLimiteEmpresas, tDicConfig.iLimiteFiliais, colModulosLib, tDicConfig.dtValidadeAte, sTextoSenha)
    If lErro <> SUCESSO Then Error 25967

    'coloca o ítem de menu visivel
    Set objObjeto = Me.Controls(objMenuItem.sNomeControle)
    
    'Se empresa for FORPRINT coloca ítem de menu visível
    If sCgc = "7384" And sNomeEmpresa = "Fo" Then
        objObjeto(objMenuItem.iIndiceControle).Visible = True
    Else
        objObjeto(objMenuItem.iIndiceControle).Visible = False
    End If
    
    Habilita_Geracao_Senha = SUCESSO

    Exit Function

Erro_Habilita_Geracao_Senha:

    Habilita_Geracao_Senha = Err

    Select Case Err

        Case 25965, 25966, 25967

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165229)

    End Select

    Exit Function

End Function

Private Sub Desabilita_ItemMenu_Pai(objMenuItens As ClassMenuItens, colMenuItens As Collection)
'desabilita o controle pai do item de menu passado como parametro

Dim objMenuItensAux As ClassMenuItens
Dim lErro As Long, objMenuItemPai As ClassMenuItens
Dim objObjeto As Object, bNaoTemIrmaoVisivel As Boolean

On Error GoTo Erro_Desabilita_ItemMenu_Pai

    If objMenuItens.sNomeControlePai <> "" Then
    
        'Pesquisa cada um itens de menu da tela procurando o pai do item passado
        For Each objMenuItemPai In colMenuItens

            'se o item de menu for o controle pai do passado p/esta rotina
            If objMenuItemPai.sNomeControle = objMenuItens.sNomeControlePai And objMenuItemPai.iIndiceControle = objMenuItens.iIndiceControlePai Then Exit For
        
        Next
        
        If Not objMenuItemPai Is Nothing Then
        
            bNaoTemIrmaoVisivel = True
            
            'Pesquisa cada um itens de menu da tela
            For Each objMenuItensAux In colMenuItens
        
                'se sao irmaos (do pai, isto é, tios)
                If (Not (objMenuItensAux Is objMenuItemPai)) And objMenuItensAux.sNomeControlePai = objMenuItemPai.sNomeControlePai And _
                    objMenuItensAux.iIndiceControlePai = objMenuItemPai.iIndiceControlePai Then
                    
                    Set objObjeto = Me.Controls(objMenuItensAux.sNomeControle)
                    
                    If objMenuItensAux.iIndiceControle <> 0 Then
                    
                        'se o irmao está visivel
                        If objObjeto(objMenuItensAux.iIndiceControle).Visible Then
                            
                            bNaoTemIrmaoVisivel = False
                            Exit For
                        
                        End If
                        
                    Else
                    
                        'se o irmao está visivel
                        If objObjeto.Visible Then
                        
                            bNaoTemIrmaoVisivel = False
                            Exit For
                                            
                        End If
                    
                    End If
                    
                End If
                
            Next
        
        Else
        
            bNaoTemIrmaoVisivel = False
        
        End If
        
        'se nao tem irmao visivel vou ter que sumir com o pai deste controle
        If bNaoTemIrmaoVisivel Then
        
            Call Desabilita_ItemMenu_Pai(objMenuItemPai, colMenuItens)
            
        Else
        
            Set objObjeto = Me.Controls(objMenuItens.sNomeControlePai)
            If objMenuItens.iIndiceControlePai <> 0 Then
                objObjeto(objMenuItens.iIndiceControlePai).Visible = False
                objObjeto(objMenuItens.iIndiceControlePai).Enabled = False
            Else
                objObjeto.Visible = False
                objObjeto.Enabled = False
            End If
    
        End If
        
    End If
    
    Exit Sub
     
Erro_Desabilita_ItemMenu_Pai:

    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165230)
     
    End Select
     
    Exit Sub

End Sub

Private Function Habillita_Itens_Menu_PaiVisivel(colItensPaiVisiveis As Collection, ByVal objMenuItens1 As ClassMenuItens) As Boolean
'retorna se o pai de objMenuItens1 está na lista de colItensPaiVisiveis
Dim bAchou As Boolean
Dim objMenuItens As ClassMenuItens

    bAchou = False
    For Each objMenuItens In colItensPaiVisiveis
        If objMenuItens.sNomeControlePai = objMenuItens1.sNomeControlePai And _
            objMenuItens.iIndiceControlePai = objMenuItens1.iIndiceControlePai Then
            bAchou = True
            Exit For
        End If
    Next
    
    Habillita_Itens_Menu_PaiVisivel = bAchou
    
End Function

Private Sub Habillita_Itens_Menu_Aux(colMenuItens As Collection, colItensPaiVisiveis As Collection, ByVal objMenuItens1 As ClassMenuItens)
'inclui o controle pai de objMenuItens1 na lista de pais visiveis, recursivamente
Dim objMenuItemPai As ClassMenuItens
Dim objMenuItemAux As ClassMenuItens

    'se o controle pai ainda nao está na coleção
    If Habillita_Itens_Menu_PaiVisivel(colItensPaiVisiveis, objMenuItens1) = False Then
        
        Set objMenuItemPai = New ClassMenuItens
        objMenuItemPai.sNomeControlePai = objMenuItens1.sNomeControlePai
        objMenuItemPai.iIndiceControlePai = objMenuItens1.iIndiceControlePai
        colItensPaiVisiveis.Add objMenuItemPai
        
        'verifica se existe um avo:
        For Each objMenuItemAux In colMenuItens
            If objMenuItemAux.sNomeControle = objMenuItens1.sNomeControlePai And _
                objMenuItemAux.iIndiceControle = objMenuItens1.iIndiceControlePai Then
                
                If objMenuItemAux.sNomeControlePai <> "" Then
                    Call Habillita_Itens_Menu_Aux(colMenuItens, colItensPaiVisiveis, objMenuItemAux)
                End If
                Exit For
                
            End If
        Next
        
    End If

End Sub

Private Function Habilita_Itens_Menu(colMenuItens As Collection) As Long

Dim lErro As Long
Dim objControl As Control
Dim objObjeto As Object
Dim objUsuItensMenu As New ClassUsuarioItensMenu
Dim colUsuarioItensMenu As New Collection
'Dim iPaiVisivel As Integer
Dim iMenuVisivel As Integer
Dim objMenuItens As ClassMenuItens
Dim objMenuItens1 As ClassMenuItens
Dim objUsuarioItensMenu As ClassUsuarioItensMenu
Dim iIndex As Integer
Dim colItensPaiVisiveis As New Collection, bRemover As Boolean

On Error GoTo Erro_Habilita_Itens_Menu

    'Preenche objUsuarioModulo
    objUsuItensMenu.sCodUsuario = gsUsuario
    objUsuItensMenu.lCodEmpresa = glEmpresa
    objUsuItensMenu.iCodFilial = giFilialEmpresa
    objUsuItensMenu.dtDataValidade = gdtDataAtual

    'Lê todos os itens de menu
    lErro = CF("MenuItens_Le", colMenuItens)
    If lErro <> SUCESSO Then Error 44634

    'Lê os itens do menu visiveis para o usuario
    lErro = CF("UsuarioItensMenu_Le", objUsuItensMenu, colUsuarioItensMenu)
    If lErro <> SUCESSO Then Error 43369

    'retira de colUsuarioItensMenu os itens que nao podem ser exibidos
    'em funcao do local de operacao
    Call Habilita_Itens_Menu2(colUsuarioItensMenu)
    
'    iPaiVisivel = False

    'Pesquisa cada um itens de menu da tela
    For Each objMenuItens In colMenuItens

        If giTipoVersao = VERSAO_LIGHT Then
        
            bRemover = False
            
            If UCase(objMenuItens.sNomeControle) = "MNUARQSUB" Then
            
                Select Case objMenuItens.iIndiceControle
                
                    Case 1, 8, 9, 12, 13, 14, 15
                        bRemover = True
                    
                End Select
                
            End If
            
            If UCase(objMenuItens.sNomeControle) = "MNUARQSUBOC" And objMenuItens.iIndiceControle = 4 Then bRemover = True
        
            If bRemover Then
            
                Set objObjeto = Me.Controls(objMenuItens.sNomeControle)
                objObjeto(objMenuItens.iIndiceControle).Visible = False
                
            End If
        
        End If
        
        If Len(objMenuItens.sNomeTela) > 0 Then

            'se já tinha tratado um menu anteriormente
            If Not (objMenuItens1 Is Nothing) Then

                'se a indicacao de visibilidade do menu for falsa
                If iMenuVisivel = False Then

                    'se o menu anterior é irmão do menu atual
                    If objMenuItens1.sNomeControlePai = objMenuItens.sNomeControlePai And objMenuItens1.sNomeControlePai <> "" _
                        And objMenuItens1.iIndiceControlePai = objMenuItens.iIndiceControlePai Then

                        'coloca o menu anterior como invisivel
                        Set objObjeto = Me.Controls(objMenuItens1.sNomeControle)
                        If objMenuItens1.iIndiceControle <> 0 Then
                            objObjeto(objMenuItens1.iIndiceControle).Visible = False
                        Else
                            objObjeto.Visible = False
                        End If
                        
                    Else

                        'se o menu anterior não é irmão do menu atual ==> os irmãos do menu anterior acabaram

                        'se houver controle pai e este estiver marcado como invisivel
                        If Habillita_Itens_Menu_PaiVisivel(colItensPaiVisiveis, objMenuItens1) = False And objMenuItens1.sNomeControlePai <> "" Then

                            'tornar invisivel o menu pai do objMenuItens1
                            Call Desabilita_ItemMenu_Pai(objMenuItens1, colMenuItens)
                            
                        Else

                            'se o menu pai estiver marcado como visivel, torna invisivel o menu anterior
                            Set objObjeto = Me.Controls(objMenuItens1.sNomeControle)
                            If objMenuItens1.iIndiceControle <> 0 Then
                                objObjeto(objMenuItens1.iIndiceControle).Visible = False
                            Else
                                objObjeto.Visible = False
                            End If

                        End If
                    End If

                End If

                'se o menu anterior não é irmão do menu atual ==> resseta a flag que indica a visibilidade do menu pai
'                If (objMenuItens1.sNomeControlePai <> objMenuItens.sNomeControlePai) Or objMenuItens.sNomeControlePai = "" Or (objMenuItens1.iIndiceControlePai <> objMenuItens.iIndiceControlePai) Then iPaiVisivel = False

            End If

            Set objMenuItens1 = objMenuItens

            iMenuVisivel = False

            'pesquisa a visibilidade do item de menu
            For Each objUsuItensMenu In colUsuarioItensMenu

                'se o item de menu for acessivel
                If objMenuItens1.sNomeControle = objUsuItensMenu.sNomeControle And objMenuItens1.iIndiceControle = objUsuItensMenu.iIndiceControle Then

                    'seta a indicacao de visibilidade do item
                    iMenuVisivel = True

                    'seta a indicacao de visiblidade do pai do item
                    Call Habillita_Itens_Menu_Aux(colMenuItens, colItensPaiVisiveis, objMenuItens1)

                    Exit For

                End If

            Next

        End If

    Next

    If iMenuVisivel = False Then

        If Habillita_Itens_Menu_PaiVisivel(colItensPaiVisiveis, objMenuItens1) = False And objMenuItens1.sNomeControlePai <> "" Then
                            
            'tornar invisivel o menu pai do objMenuItens1
            Call Desabilita_ItemMenu_Pai(objMenuItens1, colMenuItens)
        
        Else
            Set objObjeto = Me.Controls(objMenuItens1.sNomeControle)
            If objMenuItens1.iIndiceControle <> 0 Then
                objObjeto(objMenuItens1.iIndiceControle).Visible = False
            Else
                objObjeto.Visible = False
            End If
        End If
    End If

    Habilita_Itens_Menu = SUCESSO

    Exit Function

Erro_Habilita_Itens_Menu:

    Habilita_Itens_Menu = Err

    Select Case Err

        Case 43369, 44634
'
'        Case 387, 730
'            Resume Next
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165231)

    End Select

    Exit Function

End Function

Private Function Habilita_Separadores(colMenuItens As Collection) As Long

Dim iItensAnteriores As Integer
Dim iItensPosteriores As Integer
Dim objObjeto As Object
Dim objSeparador As Object
Dim objMenuItens As ClassMenuItens
Dim objMenuItens1 As ClassMenuItens, sTitulo As String
Dim lErro As Long, bControleExiste As Boolean

On Error GoTo Erro_Habilita_Separadores

    iItensAnteriores = False
    iItensPosteriores = False

    'Pesquisa cada item de menu da tela
    For Each objMenuItens In colMenuItens

        If objMenuItens.sNomeControle <> "" Then
        
            bControleExiste = True
            
            If objMenuItens.iIndiceControle = 0 Then
                Set objObjeto = Me.Controls(objMenuItens.sNomeControle)
            Else
                Set objObjeto = Me.Controls(objMenuItens.sNomeControle)(objMenuItens.iIndiceControle)
            End If
    
            If bControleExiste Then
            
                sTitulo = objObjeto.Caption
                
                If bControleExiste Then
            
                    If Not (objMenuItens1 Is Nothing) Then
            
                        'se o menu anterior não é irmão do menu atual
                        If objMenuItens1.sNomeControlePai <> objMenuItens.sNomeControlePai Or objMenuItens1.sNomeControlePai = "" _
                            Or objMenuItens1.iIndiceControlePai <> objMenuItens.iIndiceControlePai Then
            
                           If Not (objSeparador Is Nothing) And (iItensAnteriores = False Or iItensPosteriores = False) Then
            
                                objSeparador.Visible = False
            
                           End If
            
                           Set objSeparador = Nothing
                           iItensAnteriores = False
                           iItensPosteriores = False
            
                       End If
            
                    End If
            
                    'se for um separador
                    If sTitulo = SEPARADOR Then
            
                        If Not (objSeparador Is Nothing) And (iItensAnteriores = False Or iItensPosteriores = False) Then
            
                            objSeparador.Visible = False
            
                        End If
            
                        Set objSeparador = objObjeto
                        If iItensPosteriores Then iItensAnteriores = iItensPosteriores
                        iItensPosteriores = False
            
                    Else
            
                        'se o item em questão não for um separador e estiver visivel
                        If objObjeto.Visible = True Then
            
                            'se ainda não tiver passado por um separador ==> o item é anterior ao separador
                            If objSeparador Is Nothing Then
                                iItensAnteriores = True
                            Else
                                'se já tiver passado por um separador ==> o item é posterior ao separador
                                iItensPosteriores = True
                            End If
            
                        End If
            
                    End If
            
                    Set objMenuItens1 = objMenuItens
    
                End If
                
            End If
        
        End If
        
    Next

    If Not (objSeparador Is Nothing) And (iItensAnteriores = False Or iItensPosteriores = False) Then

        objSeparador.Visible = False

    End If

    Habilita_Separadores = SUCESSO

    Exit Function

Erro_Habilita_Separadores:

    Habilita_Separadores = Err

    Select Case Err

        Case 340 'se nao achou o controle(indice)
            bControleExiste = False
            Resume Next
        
        Case 730 'se nao achou o controle
            bControleExiste = False
            Resume Next
            
        Case 387
            Resume Next
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165232)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboModulo() As Long

Dim lErro As Long
Dim collCodigoNome As New AdmCollCodigoNome
Dim objlCodigoNome As AdmlCodigoNome
Dim objUsuarioModulo As New ClassUsuarioModulo

On Error GoTo Erro_Carrega_ComboModulo

    ComboModulo.Clear
    
    'Preenche objUsuarioModulo
    objUsuarioModulo.sCodUsuario = gsUsuario
    objUsuarioModulo.lCodEmpresa = glEmpresa
    objUsuarioModulo.iCodFilial = giFilialEmpresa
    objUsuarioModulo.dtDataValidade = Date
    
    'Lê os Módulos
    lErro = CF("UsuarioModulos_Le", objUsuarioModulo, collCodigoNome)
    If lErro <> SUCESSO Then Error 43630

    For Each objlCodigoNome In collCodigoNome

        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

            If UCase(objlCodigoNome.sNome) = "LOJA" Then
                'Insere na combo de Módulos
                ComboModulo.AddItem objlCodigoNome.sNome
                ComboModulo.ItemData(ComboModulo.NewIndex) = objlCodigoNome.lCodigo
            End If

        Else
            'Insere na combo de Módulos
            ComboModulo.AddItem objlCodigoNome.sNome
            ComboModulo.ItemData(ComboModulo.NewIndex) = objlCodigoNome.lCodigo
        End If

    Next

    'ativa/desativa a opção de menu que acessa associacao de conta com ccl (contabil ou extra-contabil)
    Call MenuCadCTB_Contabil_ExtraContabil
    
    Carrega_ComboModulo = SUCESSO

    Exit Function

Erro_Carrega_ComboModulo:

    Carrega_ComboModulo = Err

    Select Case Err

        Case 43630

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165233)

    End Select

    Exit Function

End Function

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim iRetornoTela As Integer
    Dim iRetornoTelaAux As Integer
    If Timer1.Interval > 0 Then
        iRetornoTelaAux = giRetornoTela
        Call Chama_Tela_Modal("AvisoWFW", 2)
        iRetornoTela = giRetornoTela
        giRetornoTela = iRetornoTelaAux
        If iRetornoTela = vbCancel Then Cancel = True
    End If
    If (Sistema_QueryUnload() = False) Then Cancel = True
End Sub

Private Sub MDIForm_Resize()

    'Altera o Tamanho da linha que divide o menu da Picture1
    LinhaDivisao.X2 = Me.Width + 200

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

Dim lErro As Long

    If gl_ST_ComandoSeta <> 0 Then

        lErro = Comando_Fechar(gl_ST_ComandoSeta)
        gl_ST_ComandoSeta = 0

    End If

    Set objAdmSeta = Nothing
    
    'Edicao Tela
    If mnuArqSub(MENU_ARQ_EDICAO).Checked = True Then Call mnuEdicao_Click

    If (Not (gobjEstInicial Is Nothing)) Then
        
        Unload gobjEstInicial
        Set gobjEstInicial = Nothing
        
    End If
    
    Call Usuario_Altera_SituacaoLogin(gsUsuario, USUARIO_NAO_LOGADO)

    lErro = Sistema_Fechar()
    
End Sub

Private Sub mnuArqAdmSistema_Click()
'Faz a chamada do Dicionário de Dados

    'Faz a chamada com o path atual do corporator.
    Call Shell(App.Path & "\DicPrincipal2.exe", vbNormalFocus)

End Sub

Private Sub mnuArqECF_Click()
'Faz a chamada do ECF

    'Faz a chamada com o path atual do corporator.
    Call Shell(App.Path & "\SGEECF.exe", vbNormalFocus)

End Sub

Private Sub mnuArqConfImpr_Click()
    Call Sist_ImpressoraDlg(1)
End Sub

Private Sub mnuArqCotMoeda_Click()
    
    Call Chama_Tela("CotacaoMoeda")

End Sub

'Private Sub mnuArqImpressoras_Click()
'    ShellExecute Me.hWnd, "Open", App.Path & "\IMP.lnk", vbNullString, vbNullString, vbMaximizedFocus
'End Sub

Private Sub mnuArqOutLook_Click()
Dim sPgm As String, iPos As Integer
    'ShellExecute Me.hWnd, "Open", "msimn", vbNullString, vbNullString, vbNormalFocus
    
    'ShellExecute Me.hWnd, "Open", "mailto:", vbNullString, vbNullString, vbNormalFocus
    
    sPgm = Obter_Pgm_Padrao_Email
    
    If Len(Trim(sPgm)) > 0 Then
        sPgm = Replace(LCase(sPgm), """", "")
        iPos = InStr(sPgm, ".exe")
        If iPos <> 0 Then sPgm = left(sPgm, iPos + 3)
        sPgm = Replace(sPgm, "%programfiles%", Environ$("ProgramFiles"))
        
        ShellExecute Me.hWnd, "Open", sPgm, vbNullString, vbNullString, vbNormalFocus
    End If
End Sub


Private Sub mnuArqWorkFlow_Click()
    Call Chama_Tela("Workflow", Me)
End Sub

Private Sub mnuArqMoedas_Click()
'??? Trecho temporário para teste do suporte
    
    Call Chama_Tela("Moedas")

End Sub

Private Sub mnuAjud_Click(Index As Integer)

    Select Case Index
    
        Case MENU_AJUDA_INDICE
            Call WinHelp(hWnd, App.HelpFile, HELP_CONTENTS, CLng(0))
    
        Case MENU_AJUDA_SUPORTE
            Call mnuAjudaSuporte_Click

        Case MENU_AJUDA_ATUALIZACAO
            Call mnuAjudaAtualizacao_Click

    End Select

End Sub

Private Sub mnuArqSub_Click(Index As Integer)

    Select Case Index
    
        Case MENU_ARQ_FILIAL
            Call mnuArqFilialEmpresa_Click
            
        Case MENU_ARQ_DATA
            Call mnuArqData_Click

        Case MENU_ARQ_FERIADO
            Call mnuArqFeriados_Click
        
        Case MENU_ARQ_COTACAOMOEDA
            Call mnuArqCotMoeda_Click
        
        Case MENU_ARQ_CADASTROMOEDA
            Call mnuArqMoedas_Click

'        Case MENU_ARQ_IMPRESSORA
'            Call mnuArqConfImpr_Click

        Case MENU_ARQ_EDICAO
            Call mnuEdicao_Click
            
        Case MENU_ARQ_ADMIN
            Call mnuArqAdmSistema_Click
        
        Case MENU_ARQ_ECF
            Call mnuArqECF_Click
        
        Case MENU_ARQ_WORKFLOW
            Call mnuArqWorkFlow_Click
        
        Case MENU_ARQ_OUTLOOK
            Call mnuArqOutLook_Click
        
        Case MENU_ARQ_SAIR
            Call mnuArqSair_Click

    End Select

End Sub

Private Sub mnuAjudaSuporte_Click()

Dim sDummy As String
Dim sBrowserExec As String
Dim lRetVal As Long
Dim sURL As String

On Error GoTo Erro_mnuAjudaSuporte_Click

    sBrowserExec = Space(255)
    
    'Aproveita o htm com o Javascript que abre o chat no tamanho certo para verificar o navegador padrão do usuário
    lRetVal = FindExecutable(App.Path & "\suporteonline.htm", sDummy, sBrowserExec)
    sBrowserExec = Trim(sBrowserExec)
    
    'Se não achou o navegador dá erro
    If lRetVal <= 32 Or IsEmpty(sBrowserExec) Then gError 209502
    
'    sURL = App.Path & "\suporteonline.htm?usu=" & DesacentuaTexto(Replace(Replace(gsUsuario, " ", "+"), "&", " ")) & "&emp=" & DesacentuaTexto(Replace(Replace(gsNomeEmpresa, " ", "+"), "&", " "))
    sURL = "file://" & Replace(App.Path, "\", "/") & "/suporteonline.htm?usu=" & DesacentuaTexto(Replace(Replace(gsUsuario, " ", "+"), "&", " ")) & "&emp=" & DesacentuaTexto(Replace(Replace(gsNomeEmpresa, " ", "+"), "&", " "))

    'Abre o htm que redireciona para o chat
    lRetVal = ShellExecute(Me.hWnd, "open", sBrowserExec, sURL, sDummy, SW_NORMAL)
      
    Exit Sub
      
Erro_mnuAjudaSuporte_Click:

    Select Case gErr
    
        Case 209502
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_NAVEGADOR_PADRAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209503)
        
    End Select

    Exit Sub

End Sub

Private Sub mnuAjudaAtualizacao_Click()

Dim lngLen As Long, lngX As Long
Dim strCompName As String

On Error GoTo Erro_mnuAjudaAtualizacao_Click

    lngLen = 255
    strCompName = String$(lngLen, 0)
    lngX = GetComputerName(strCompName, lngLen)
    If lngX <> 0 Then
        If UCase(left(strCompName, 3)) = "ASP" Then
            Call Rotina_Aviso(vbOKOnly, "AVISO_FUNCIONALIDADE_SO_INST_LOCAL")
        Else
            Call WinExec(App.Path & "\AtualizaCorporator", SW_NORMAL)
        End If
    End If
    
    Exit Sub
      
Erro_mnuAjudaAtualizacao_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209504)
        
    End Select

    Exit Sub
    
End Sub

Private Sub mnuArqSubOC_Click(Index As Integer)
Dim lErro As Long

On Error GoTo Erro_mnuArqSubOC_Click

    Select Case Index
        Case MENU_ARQ_OC_IMPRESSORA
            Call mnuArqConfImpr_Click
        Case MENU_ARQ_OC_BACKUP
            Call Chama_Tela("BackupConfig")
        Case MENU_ARQ_OC_LOGO
            Call Chama_Tela("Logo")
        Case MENU_ARQ_OC_TELA
            If mnuArqSubOC(Index).Checked = False Then
            
                lErro = CF("Config_Grava", "AdmConfig", "TELAS_TAMANHO_VARIAVEL", EMPRESA_TODA, 1)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                mnuArqSubOC(Index).Checked = True
                
                giTelaTamanhoVariavel = MARCADO
            
            Else
            
                lErro = CF("Config_Grava", "AdmConfig", "TELAS_TAMANHO_VARIAVEL", EMPRESA_TODA, 0)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                mnuArqSubOC(Index).Checked = False
                
                giTelaTamanhoVariavel = DESMARCADO
                
            End If
        Case MENU_ARQ_OC_ARQUIVAMENTO
            Call Chama_Tela("Arquivamento")
            'If giRetornoTela = vbOK Then Unload Me
    
    End Select
    
    Exit Sub
      
Erro_mnuArqSubOC_Click:

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210541)
        
    End Select

    Exit Sub
    
End Sub

Private Sub mnuCOMConCad_Click(Index As Integer)

    Select Case Index
        
        Case MENU_COM_CON_CAD_PRODUTOS
            Call Chama_Tela("ProdutoLista_Consulta")
        
        Case MENU_COM_CON_CAD_FORNECEDORES
            Call Chama_Tela("FornecedorLista")
        
        Case MENU_COM_CON_CAD_PRODFORN
            Call Chama_Tela("FornFilialProdutoLista")
        
        Case MENU_COM_CON_CAD_REQUISITANTES
            Call Chama_Tela("RequisitanteLista")
        
        Case MENU_COM_CON_CAD_COMPRADORES
            Call Chama_Tela("CompradoresLista")
      
        Case MENU_COM_CON_CAD_ALCADA
            Call Chama_Tela("AlcadaUsuarioLista")
      
    End Select

End Sub

Private Sub mnuCOMConcorrencias_Click(Index As Integer)
    Select Case Index
        Case 1
            Call Chama_Tela("ConcorrenciaTodasLista")
        Case 2
            Call Chama_Tela("ConcorrenciaBaixadaLista")
        Case 3
            Call Chama_Tela("ConcorrenciaLista")
    End Select
End Sub

Private Sub mnuCOMConPedCompra_Click(Index As Integer)
Dim colSelecao As New Collection
    Select Case Index
        Case 1
            Call Chama_Tela("PedComprasTodosLista")
        Case 2
            Call Chama_Tela("PedComprasEnvLista")
        Case 3
            Call Chama_Tela("PedComprasNaoEnvLista")
        Case 4
            Call Chama_Tela("PedidoCompraAbertoLista")
        Case 5
            Call Chama_Tela("ItensPedCompraLista")
        Case 6
            Call Chama_Tela("UltPCPorProdLista", colSelecao, Nothing, Nothing, "Posicao=1")
        Case 7
            Call Chama_Tela("ReqCompraPedCompraLista")

    End Select
End Sub

Private Sub mnuCOMConPedCotacao_Click(Index As Integer)
Dim colSelecao As New Collection
    Select Case Index
        Case 1
            Call Chama_Tela("PedidoCotacaoTodosLista")
        Case 2
            Call Chama_Tela("PedidoCotacaoBaixadoLista")
        Case 3
            Call Chama_Tela("PedidoCotacaoLista")
        Case 5
            Call Chama_Tela("PedCotCompletoLista", colSelecao, Nothing, Nothing, "ItemCotPrecoUnitario IS NULL")
        Case 6
            Call Chama_Tela("PedCotCompletoLista", colSelecao, Nothing, Nothing, "NOT ItemCotPrecoUnitario IS NULL")
    End Select
End Sub

Private Sub mnuCOMConReqCompra_Click(Index As Integer)
    Select Case Index
        Case 1
            Call Chama_Tela("ReqComprasTodasLista")
        Case 2
            Call Chama_Tela("ReqComprasEnvLista")
        Case 3
            Call Chama_Tela("ReqComprasNaoEnvLista")
        Case 4
            Call Chama_Tela("ItensReqCompraLista")
    End Select
End Sub

Private Sub mnuCOMMovCotacao_Click(Index As Integer)
    Select Case Index
        Case 1
            Call Chama_Tela("GeracaoPedCotacao")
        Case 2
            Call Chama_Tela("GeracaoPedCotacaoAvulsa")
        Case 3
            Call Chama_Tela("PedidoCotacao")
        Case 4
            Call Chama_Tela("BaixaPedCotacao")
        Case 6
            Call Chama_Tela("MapaCotacao")
    End Select
End Sub

Private Sub mnuCOMMovGerPC_Click(Index As Integer)
    Select Case Index
        Case 1
            Call Chama_Tela("GeracaoPedCompraGerCot")
        Case 2
            Call Chama_Tela("GeracaoPedCompraReq")
        Case 3
            Call Chama_Tela("GeracaoPedCompraConc")
        Case 4
            Call Chama_Tela("GeracaoPedCompraAvulsa")
        Case 5
            Call Chama_Tela("GeracaoPedCompraOV")
    End Select
    
End Sub

Private Sub mnuCOMMovGerReq_Click(Index As Integer)
    Select Case Index
        Case 1
            Call Chama_Tela("GeracaoReqPtoPedido")
        Case 2
            Call Chama_Tela("GeracaoReqPedVenda")
    End Select
End Sub

Private Sub mnuCOMMovPC_Click(Index As Integer)
    Select Case Index
        Case 1
            Call Chama_Tela("PedidoCompras")
        Case 2
            Call Chama_Tela("PedComprasGerado")
        Case 3
            Call Chama_Tela("LiberaBloqueioPC")
        Case 4
            Call Chama_Tela("BaixaPedCompras")
        Case 5
            Call Chama_Tela("PedCompraAprova")
    End Select
End Sub

Private Sub mnuCOMMovReq_Click(Index As Integer)
    Select Case Index
        Case 1
            Call Chama_Tela("ReqCompras")
        Case 2
            Call Chama_Tela("BaixaReqCompras")
        Case 3
            Call Chama_Tela("ReqCompraEnvio")
        Case 4
            Call Chama_Tela("ReqCompraAprova")
    End Select
End Sub

Private Sub mnuCOMRel_Click(Index As Integer)

Dim lErro As Long
Dim iCancela As Integer
Dim sCodRel As String
Dim objRelSel As New AdmRelSel
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_mnuESTRel_Click

    Select Case Index

        Case MENU_COM_REL_CONCABERT
            objRelatorio.Rel_Menu_Executar ("Concorrências Abertas")

        Case MENU_COM_REL_ANCOTREC
            objRelatorio.Rel_Menu_Executar ("Análise de Cotações Recebidas")

        Case MENU_COM_REL_PREVENTPROD
            objRelatorio.Rel_Menu_Executar ("Previsão de Entrega de Produtos")

        Case MENU_COM_REL_AGPREVENT
            objRelatorio.Rel_Menu_Executar ("Agenda de Previsão de Entrega")

        Case MENU_COM_REL_OUTROS
            lErro = Chama_Tela("RelSelecionar", objRelSel, MODULO_COMPRAS)
            If (lErro <> 0) Then Error 7058

            If (objRelSel.iCancela <> 0) Then Exit Sub

            'Prosseguir executando relatório
            objRelatorio.Rel_Menu_Executar (objRelSel.sCodRel)

        Case MENU_COM_REL_GERREL
            Sistema_EditarRel ("")

        '####################################
        'Inserido por Wagner
        Case MENU_COM_REL_PLANILHAS
            lErro = Chama_Tela("PlanilhasSelecionar", MODULO_COMPRAS)
            If (lErro <> 0) Then Error 7058
        '####################################
        
    End Select

    Exit Sub

Erro_mnuESTRel_Click:

    Select Case Err

        Case 7058

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165234)

    End Select

    Exit Sub

End Sub

Private Sub mnuCOMRelCad_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_COM_REL_CAD_REQ
            objRelatorio.Rel_Menu_Executar ("Requisitantes")

        Case MENU_COM_REL_CAD_COMP
            objRelatorio.Rel_Menu_Executar ("Compradores")

        Case MENU_COM_REL_CAD_FORN
            objRelatorio.Rel_Menu_Executar ("Fornecedores")

    End Select
    
End Sub

Private Sub mnuCOMRelPC_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_COM_REL_PC_ABERTO
            objRelatorio.Rel_Menu_Executar ("Pedidos de Compra em Aberto")

        Case MENU_COM_REL_PC_BAIXADOS
            objRelatorio.Rel_Menu_Executar ("Pedidos de Compra Baixados")

        Case MENU_COM_REL_PC_ATRASADOS
            objRelatorio.Rel_Menu_Executar ("Pedidos de Compra Atrasados")
        
        Case MENU_COM_REL_PC_BLOQUEADOS
            objRelatorio.Rel_Menu_Executar ("Pedidos de Compra Bloqueados")

        Case MENU_COM_REL_PC_EMISSAO
            objRelatorio.Rel_Menu_Executar ("Pedidos de Compra Emitidos")

        Case MENU_COM_REL_PC_EMISSCONC
            objRelatorio.Rel_Menu_Executar ("Pedidos de Compra Emitidos por Concorrência")
        
        Case MENU_COM_REL_PC_NF
            objRelatorio.Rel_Menu_Executar ("Pedidos de Compra x Notas Fiscais")

    End Select
    
End Sub

Private Sub mnuCOMRelPCot_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_COM_REL_PCOT_EMISSAO
            objRelatorio.Rel_Menu_Executar ("Pedidos de Cotação Emitidos")

        Case MENU_COM_REL_PCOT_EMISSGER
            objRelatorio.Rel_Menu_Executar ("Pedidos de Cotação Emitidos por Geração")

    End Select

End Sub

Private Sub mnuCOMRelProd_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_COM_REL_PROD_COT
            objRelatorio.Rel_Menu_Executar ("Produtos x Cotações")

        Case MENU_COM_REL_PROD_FORN
            objRelatorio.Rel_Menu_Executar ("Produtos x Fornecedores - Compras")

        Case MENU_COM_REL_PROD_COMPRAS
            objRelatorio.Rel_Menu_Executar ("Produtos x Compras")
        
        Case MENU_COM_REL_PROD_REQ
            objRelatorio.Rel_Menu_Executar ("Produtos x Requisições")

        Case MENU_COM_REL_PROD_PC
            objRelatorio.Rel_Menu_Executar ("Produtos x Pedidos de Compra")

    End Select
    
End Sub

Private Sub mnuCOMRelRC_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_COM_REL_RC_ABERTO
            objRelatorio.Rel_Menu_Executar ("Requisições de Compra em Aberto")

        Case MENU_COM_REL_RC_BAIXADOS
            objRelatorio.Rel_Menu_Executar ("Requisições de Compra Baixadas")

        Case MENU_COM_REL_RC_ATRASADOS
            objRelatorio.Rel_Menu_Executar ("Requisições de Compra Atrasadas")
        
        Case MENU_COM_REL_RC_PV
            objRelatorio.Rel_Menu_Executar ("Requisições de Compra x Pedidos de Venda")

        Case MENU_COM_REL_RC_OP
            objRelatorio.Rel_Menu_Executar ("Requisições de Compra x Ordens de Produção")

    End Select
    
End Sub

Private Sub mnuComRot_Click(Index As Integer)

    Select Case Index
                
        Case MENU_COM_ROT_PARAMPONTOPEDIDO
            Call Chama_Tela("ParametrosPtoPed")
    
    End Select
    
End Sub

Private Sub mnuCPConCad_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
        
        Case MENU_CP_CON_CAD_FORNECEDORES
            Call Chama_Tela("FornecedorLista")
        
        Case MENU_CP_CON_CAD_TIPOFORN
            Call Chama_Tela("TipoFornecedorLista")
        
        Case MENU_CP_CON_CAD_BANCOS
            Call Chama_Tela("BancoLista")
        
        Case MENU_CP_CON_CAD_CONTASCORRENTESFILIAIS
            Call Chama_Tela("CtaCorrenteTodasLista")
      
        Case MENU_CP_CON_CAD_PORTADORES
            Call Chama_Tela("PortadoresLista")
      
        Case MENU_CP_CON_CAD_CONDPAG
            Call Chama_Tela("CondicaoPagtoCPLista", colSelecao)
                    
    End Select

End Sub

Private Sub mnuCPMovBaixa_Click(Index As Integer)
   
    Select Case Index

        Case MENU_CP_MOV_BAIXA_CANCELARBAIXA_MANUAL
            Call Chama_Tela("BaixaPagCancelar")

        Case MENU_CP_MOV_BAIXA_CANCELARBAIXA_SELECAO
            Call Chama_Tela("BaixaPagtosCancelar")

        Case MENU_CP_MOV_BAIXA_CANCELARBAIXA_ADIANT
            Call Chama_Tela("BaixaAntecipPagCanc")

    End Select

End Sub

Private Sub mnuCPMovBx_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
    
        Case MENU_CP_MOV_BX_BAIXAMANUAL
            Call Chama_Tela("BaixaPag")

        Case MENU_CP_MOV_BX_BAIXAADIANFORN
            Call Chama_Tela("BaixaAntecipCredFornecedor")

        Case MENU_CP_MOV_BX_BAIXAMANUAL_CHEQUETERC
            Call Chama_Tela("BaixaPagChequeTerc")

    End Select
   
End Sub

Private Sub mnuCRConCad_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
    
        Case MENU_CR_CON_CAD_CLIENTES
            Call Chama_Tela("ClientesLista")

        Case MENU_CR_CON_CAD_COBRADORES
            Call Chama_Tela("CobradorLista")

        Case MENU_CR_CON_CAD_VENDEDORES
            Call Chama_Tela("VendedorLista")

        Case MENU_CR_CON_CAD_CONDPAG
            Call Chama_Tela("CondicaoPagtoCRLista", colSelecao)

        Case MENU_CR_CON_CAD_TRANSP
            Call Chama_Tela("TransportadoraLista", colSelecao)

    End Select
    
End Sub

'Incluído por Luiz Nogueira em 13/01/04
Private Sub mnuCRMMov_Click(Index As Integer)

    Select Case Index

        Case MENU_CRM_MOV_RELACCLI
            Call Chama_Tela("RelacionamentoClientes")

        Case MENU_CRM_MOV_RELACCLICONS
            Call Chama_Tela("RelacionamentoClientesCons")

        Case MENU_CRM_MOV_RELACCON
            Call Chama_Tela("RelacionamentoContatos")

        Case MENU_CRM_MOV_RELACCONCONS
            Call Chama_Tela("RelacionamentoContatoCons")

        Case MENU_CRM_MOV_CONTATOCLI
            Call Chama_Tela("ContatoCliente")

    End Select

End Sub

'Incluído por Luiz Nogueira em 13/01/04
Private Sub mnuCRMCon_Click(Index As Integer)

    Select Case Index

        Case MENU_CRM_CON_CLIENTECONS
            Call Chama_Tela("ClienteConsulta")

    End Select

End Sub

'Incluído por Luiz Nogueira em 13/01/04
Private Sub mnuCRMConCad_Click(Index As Integer)
    
    Select Case Index

        Case MENU_CRM_CONCAD_CLIENTES
            Call Chama_Tela("ClientesLista")

        Case MENU_CRM_CONCAD_ATENDENTES
            Call Chama_Tela("Atendentes_Lista")

        Case MENU_CRM_CONCAD_VENDEDORES
            Call Chama_Tela("VendedorLista")

        Case MENU_CRM_CONCAD_CLICONTATOS
            Call Chama_Tela("ClienteContatos_Lista")

        Case MENU_CRM_CONCAD_CLICONTATOSTEL
            Call Chama_Tela("CliContatosLista")

    End Select

End Sub

'Incluído por Luiz Nogueira em 13/01/04
Private Sub mnuCRMConRelac_Click(Index As Integer)

    Select Case Index

        Case MENU_CRM_CONRELAC_PENDENTES
            Call Chama_Tela("RelacionamentoClientes_Pendentes_Lista")

        Case MENU_CRM_CONRELAC_ENCERRADOS
            Call Chama_Tela("RelacionamentoClientes_Encerrados_Lista")

        Case MENU_CRM_CONRELAC_TODOS
            Call Chama_Tela("RelacionamentoClientes_Lista")

        Case MENU_CRM_CONRELAC_CALLCENTER
            Call Chama_Tela("ContatosCallCenterLista")

        Case MENU_CRM_CONRELAC_SOLSRV
            Call Chama_Tela("RelacCliSolSRVLista")


    End Select

End Sub

'Incluído por Luiz Nogueira em 13/01/04
Private Sub mnuCRMCad_Click(Index As Integer)

    Select Case Index

        Case MENU_CRM_CAD_CLIENTES
            Call Chama_Tela("Clientes")

        Case MENU_CRM_CAD_ATENDENTES
            Call Chama_Tela("Atendentes")

        Case MENU_CRM_CAD_CAMPOSGENERICOS
            Call Chama_Tela("CamposGenericos")
        
        Case MENU_CRM_CAD_VENDEDORES
            Call Chama_Tela("Vendedores")
        
    End Select

End Sub

'Incluído por Luiz Nogueira em 13/01/04
Private Sub mnuCRMCadTA_Click(Index As Integer)

    Select Case Index

        Case MENU_CRM_CADTA_CATCLIENTES
            Call Chama_Tela("CategoriaCliente")

        Case MENU_CRM_CADTA_TIPOSCLIENTES
            Call Chama_Tela("TipoCliente")
        
        Case MENU_CRM_CADTA_CLIENTECONTATOS
            Call Chama_Tela("ClienteContatos")
        
        Case MENU_CRM_CADTA_TIPOSVENDEDORES
            Call Chama_Tela("TipoVendedor")
        
        Case MENU_CRM_CADTA_REGIOESVENDA
            Call Chama_Tela("RegiaoVenda")
            
        Case MENU_CRM_CADTA_CONTATOS
            Call Chama_Tela("Contatos")
        
        Case MENU_CRM_CADTA_CLIENTEFCONTATOS
            Call Chama_Tela("ClienteFContatos")
        
        Case MENU_CRM_CADTA_CONTATOCLIPOREMAIL
            Call Chama_Tela("ModelosEmail")
                
    End Select

End Sub

Private Sub mnuCRMovBxM_Click(Index As Integer)
    Select Case Index
        Case 1
            Call Chama_Tela("BaixaRecDig")
        Case 2
            Call Chama_Tela("BaixaRec")
        Case 3
            Call Chama_Tela("BaixaRecCancelar")
        Case 4
            Call Chama_Tela("BaixaAntecipDebCliente")
    End Select
End Sub

'Incluído por Luiz Nogueira em 13/01/04
Private Sub mnuCRMRel_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio
Dim iCancela As Integer
Dim sCodRel As String
Dim objRelSel As New AdmRelSel
Dim lErro As Long

On Error GoTo Erro_mnuCRMRel_Click

    Select Case Index

        Case MENU_CRM_REL_FOLLOWUP
            objRelatorio.Rel_Menu_Executar ("FollowUp")
        
        Case MENU_CRM_REL_RELACESTAT
            objRelatorio.Rel_Menu_Executar ("Relacionamentos x Estatísticas")
        
        Case MENU_CRM_REL_CLIENTESSEMRELAC
            objRelatorio.Rel_Menu_Executar ("Clientes Sem Relacionamentos")
        
        Case MENU_CRM_REL_OUTROS
            lErro = Chama_Tela("RelSelecionar", objRelSel, MODULO_CRM)
            If (lErro <> SUCESSO) Then gError 102987

            If (objRelSel.iCancela <> 0) Then Exit Sub

            'Prosseguir executando relatório
            objRelatorio.Rel_Menu_Executar (objRelSel.sCodRel)

        Case MENU_CRM_REL_GERREL
            Sistema_EditarRel ("")
        
        '####################################
        'Inserido por Wagner
        Case MENU_CRM_REL_PLANILHAS
            lErro = Chama_Tela("PlanilhasSelecionar", MODULO_CRM)
            If (lErro <> 0) Then gError 102987
        '####################################
        
    End Select

Exit Sub

Erro_mnuCRMRel_Click:

    Select Case gErr

        Case 102987

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165235)

    End Select

    Exit Sub

End Sub

'Incluído por Luiz Nogueira em 13/01/04
Private Sub mnuCRMRelCad_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_CRM_RELCAD_CLI
            objRelatorio.Rel_Menu_Executar ("Relação de Clientes")

        Case MENU_CRM_RELCAD_ATENDENTES
            objRelatorio.Rel_Menu_Executar ("Relação de Atendentes")

        Case MENU_CRM_RELCAD_VENDEDORES
            objRelatorio.Rel_Menu_Executar ("Relação de Vendedores")

        Case MENU_CRM_RELCAD_CLIENTECONTATOS
            objRelatorio.Rel_Menu_Executar ("Clientes x Contatos")

    End Select

End Sub

'Inicio Edicao Tela
Private Sub mnuEdicao_Click()

Dim objFlag As New AdmGenerico
Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_mnuEdicao_Click
    
    If mnuArqSub(MENU_ARQ_EDICAO).Checked = True Then
        mnuArqSub(MENU_ARQ_EDICAO).Checked = False
        Unload gobjPropriedades
        
        objUsuarios.sCodUsuario = gsUsuario
    
        lErro = CF("Usuarios_Le", objUsuarios)
        If lErro <> SUCESSO Then gError 20847
    
        If objUsuarios.iWorkFlowAtivo = WORKFLOW_ATIVO Then
        
            Timer1.Interval = 60000
        
        End If
        
    Else
        
        Timer1.Interval = 0
        'Alterado por Wagner
        'Chama a tela para permitir ou não a entrada no modo de edição
        'carrega a tela p/identificacao do usuario
        '####################
        
        Load EdicaoLogin

        lErro = EdicaoLogin.Trata_Parametros(objFlag)
        If lErro <> SUCESSO Then gError 129290

        EdicaoLogin.Show vbModal

        If objFlag.vVariavel = False Then gError 129291
        '##########################
        
        Call Chama_Tela("Propriedades")
        Call Chama_Tela("CamposInvisiveis")
    
        mnuArqSub(MENU_ARQ_EDICAO).Checked = True
        
    End If
    
    Exit Sub
    
Erro_mnuEdicao_Click:

    Select Case gErr
        
        Case 129290, 129291, 20847
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165236)

    End Select
    
    Exit Sub

End Sub
'Fim Edicao Tela

Private Function Seta_Click(eSeta As enumSeta) As Long

Dim objTelaIndice As New AdmTelaIndice
Dim lErro As Long
Dim colCampoValor As New AdmColCampoValor 'campos-valores da tela
Dim objCampoValor As AdmCampoValor
Dim colCampoIndiceValor As New AdmColCampoIndiceValor  'campos do índice
Dim colSelecao As New AdmColFiltro 'filtros para o SELECT do sistema de SETAS
Dim sTabela As String
Dim iIndice As Integer
Dim eComparacao As enumSetaComparacao

On Error GoTo Erro_Seta_Click

    If gi_ST_SetaIgnoraClick = 0 And (Not gobj_ST_TelaAtiva Is Nothing) And Indice.ListIndex > -1 Then

        objTelaIndice.sNomeTela = gobj_ST_TelaAtiva.Name
        objTelaIndice.sNomeExterno = Indice.Text
        objTelaIndice.iIndice = Indice.ItemData(Indice.ListIndex)

        'Le os campos correspondentes ao índice
        lErro = TelaIndiceCampos_Le(objTelaIndice, colCampoIndiceValor)
        If lErro Then Error 6608

        'Extrai da tela nome da tabela associada e campos-valores atuais
        Call gobj_ST_TelaAtiva.Tela_Extrai(sTabela, colCampoValor, colSelecao)
        If Len(sTabela) = 0 Then Error 6610
        
        'Se o número de campos vindos da Tela for menor que o número do índice, erro
        If colCampoValor.Count < colCampoIndiceValor.Count Then Error 6611
        
        'Se tela que fez último click é a tela ativa e o índice não mudou e o comando de seta está aberto
        If gs_ST_TelaSetaClick = gobj_ST_TelaAtiva.Name And objAdmSeta.gs_ST_TelaIndice = Indice.Text And gl_ST_ComandoSeta <> 0 And eSeta <> BOTAO_CONSULTA Then

            'Compara valores dos campos índice anteriores com os novos que vem da tela
            lErro = objAdmSeta.Compara_Campos_Indice(colCampoIndiceValor, colCampoValor, eComparacao)
            If lErro <> SUCESSO Then Error 25662

            If eComparacao = SETA_COMP_IGUAL Then
            
                'Lê seguindo a seta, aproveitando comando aberto existente
                lErro = objAdmSeta.Seta_Le(gl_ST_ComandoSeta, eSeta)
                If lErro Then Error 6616
                
            ElseIf eComparacao = SETA_COMP_DIFERENTE Then
            
                'Lê seguindo a seta, abrindo comando
                lErro = objAdmSeta.Seta_Le_AbreComando(eSeta, sTabela, colCampoIndiceValor, colCampoValor, colSelecao, objTelaIndice)
                If lErro <> SUCESSO Then Error 25661
                
            End If

        Else
            
            'Lê seguindo a seta, abrindo comando
            lErro = objAdmSeta.Seta_Le_AbreComando(eSeta, sTabela, colCampoIndiceValor, colCampoValor, colSelecao, objTelaIndice)
            If lErro <> SUCESSO Then Error 25660

        End If

        'Passa os novos valores de campo para a Tela
        Call gobj_ST_TelaAtiva.Tela_Preenche(objAdmSeta.gcol_ST_CampoValor)

    End If

    Seta_Click = SUCESSO

    Exit Function

Erro_Seta_Click:

    Seta_Click = Err

    Select Case Err

        Case 6608, 6616, 25660, 25661, 25662  'Erro tratado na rotina chamada ou não faz nada (ausencia de seguintes)

        Case 6610
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_TABELA_VAZIO", Err)

        Case 6611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MENOS_CAMPOS_TELA_QUE_CAMPOS_INDICE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165237)

    End Select

    Exit Function

End Function


Private Sub mnuArqData_Click()
    
    Call BotaoData_Click

End Sub

Private Sub mnuArqFeriados_Click()

    Call Chama_Tela("Feriados")
    
    'Call CreateObject("RotinasCPR.ClassRotBaixaCartao").AdmExtFin_ImportarExtratos(1, 1)
'    Call CreateObject("RotinasCPR.ClassRotBaixaCartao").AdmExtFin_ValidarExtratos(1, 1, DATA_NULA, DATA_NULA)

    'Call CreateObject("RotinasCPR.ClassRotBaixaCartao").AdmExtFin_AtualizarExtratos(1, 1)

End Sub

Private Sub mnuArqFilialEmpresa_Click()

Dim objForm As Form, objMenu As Object, objAux As Object
Dim lErro As Long, iFilEmpAnterior As Integer
Dim colMenuItens As New Collection

On Error GoTo Erro_mnuArqFilialEmpresa_Click

    For Each objForm In Forms
        If Not (objForm Is Me) And Not (objForm Is gobjEstInicial) Then gError 44222
    Next

    iFilEmpAnterior = giFilialEmpresa
    
    EmpresaFilial.Show vbModal
    
    'se trocou de filial p/EMPRESA_TODA ou vice-versa
    If ((iFilEmpAnterior = EMPRESA_TODA Or iFilEmpAnterior = Abs(giFilialAuxiliar)) And (giFilialEmpresa <> EMPRESA_TODA And giFilialEmpresa <> Abs(giFilialAuxiliar))) Or ((iFilEmpAnterior <> EMPRESA_TODA And iFilEmpAnterior <> Abs(giFilialAuxiliar)) And (giFilialEmpresa = EMPRESA_TODA Or giFilialEmpresa = Abs(giFilialAuxiliar))) Then
    
        'vou ter que acertar o menu
        
        Me.Visible = False
        FormAguarde.Show
        DoEvents
        
        'torna todos os itens de menu visiveis novamente
        For Each objMenu In Me.Controls()
        
            If TypeName(objMenu) = "Menu" Then
                If objMenu.Index = 0 Then
                
                    objMenu.Enabled = True
                    objMenu.Visible = True
                    
                Else
                
                    Set objAux = Me.Controls(objMenu.Name)
                    objAux(objMenu.Index).Enabled = True
                    objAux(objMenu.Index).Visible = True
                    
                End If
            End If
            
        Next
        
        'esconde os itens de menus à que o usuario nao tem acesso
        lErro = Habilita_Itens_Menu(colMenuItens)
        If lErro <> SUCESSO Then gError 81500
    
        lErro = Habilita_Separadores(colMenuItens)
        If lErro <> SUCESSO Then gError 81501
        
        lErro = Mostra_Menu(NENHUM_MODULO)
        If lErro <> SUCESSO Then gError 81502

        Unload FormAguarde
        Me.Visible = True
        
    End If
    
    Call Monta_Botoes_Filiais
    
    Exit Sub
    
Erro_mnuArqFilialEmpresa_Click:

    Select Case gErr

        Case 343
            Resume Next
        
        Case 81500, 81501, 81502
        
        Case 44222
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FECHAR_JANELAS_FILHAS", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165238)

    End Select

    Unload FormAguarde
    Me.Visible = True
    
    Exit Sub

End Sub

Private Sub mnuArqModulo_Click()

    Call Chama_Tela("Modulo", Me)

End Sub


Private Sub mnuCOMCad_Click(Index As Integer)

    Select Case Index

        Case MENU_COM_CAD_PRODUTOS
            Call Chama_Tela("Produto")
        
        Case MENU_COM_CAD_FORNECEDORES
            Call Chama_Tela("Fornecedores")
        
        Case MENU_COM_CAD_FORNECEDORPRODUTOFF
            Call Chama_Tela("FornFilialProduto")
    
        Case MENU_COM_CAD_COMPRADORES
            Call Chama_Tela("Comprador")

        Case MENU_COM_CAD_REQUISITANTES
            Call Chama_Tela("Requisitante")
            
        Case MENU_COM_CAD_ALCADAS
            Call Chama_Tela("Alcada")
            
    End Select

End Sub

Private Sub mnuCOMCadTA_Click(Index As Integer)

    Select Case Index
    
        Case MENU_COM_CAD_TA_TIPOSDEPRODUTO
            Call Chama_Tela("TipoProduto")
    
        Case MENU_COM_CAD_TA_MODELOSDEREQUISICOES
            Call Chama_Tela("RequisicaoModelo")
    
        Case MENU_COM_CAD_TA_TIPOSDEBLOQUEIO
            Call Chama_Tela("TipoDeBloqueioPC")
        
        Case MENU_COM_CAD_TA_ESTOQUE
            Call Chama_Tela("Estoque")
        
        Case MENU_COM_CAD_TA_CONDICAOPAGAMENTO
            Call Chama_Tela("CondicoesPagto")
            
        Case MENU_COM_CAD_TA_TRANSPORTADORAS
            Call Chama_Tela("Transportadora")
            
        Case MENU_COM_CAD_TA_CATEGFORNECEDOR
            Call Chama_Tela("CategoriaFornec")
        
        Case MENU_COM_CAD_TA_NOTASPC
            Call Chama_Tela("NotasPC")
            
        Case MENU_COM_CAD_TA_PRODUTOFORNECEDOR
            Call Chama_Tela("ProdutoFornecedor")
        
    End Select

End Sub

Private Sub mnuCOMConfig_Click(Index As Integer)

    Select Case Index
    
        Case MENU_COM_CONFIG_CONFIGURACAO
            Call Chama_Tela("ConfiguraCOM")
            
    End Select

End Sub

Private Sub mnuCPCad_Click(Index As Integer)

Dim objLote As ClassLote

    Select Case Index

        Case MENU_CP_CAD_FORNECEDORES
            Call Chama_Tela("Fornecedores")

        Case MENU_CP_CAD_PORTADORES
            Call Chama_Tela("Portadores")

        Case MENU_CP_CAD_LOTES
            Call Chama_Tela("LoteTela", objLote, MODULO_CONTASAPAGAR)
        
        Case MENU_CP_CAD_BANCOS
            Call Chama_Tela("Bancos")

        Case MENU_CP_CAD_CONTACORRENTE
            Call Chama_Tela("CtaCorrenteInt")

    End Select

End Sub

Private Sub mnuCPCadTA_Click(Index As Integer)

    Select Case Index

        Case MENU_CP_CAD_TA_TIPOSFORN
            Call Chama_Tela("TipoFornecedor")

        Case MENU_CP_CAD_TA_CONDPAG
            Call Chama_Tela("CondicoesPagto")

        Case MENU_CP_CAD_TA_FERIADOS
            Call Chama_Tela("Feriados")

        Case MENU_CP_CAD_TA_MENSAGENS
            Call Chama_Tela("Mensagens")

        Case MENU_CP_CAD_TA_PAISES
            Call Chama_Tela("Paises")

        Case MENU_CP_CAD_TA_HISTEXTRATO
            Call Chama_Tela("HistMovCta")
            
        Case MENU_CP_CAD_TA_CATFORNECEDOR
            Call Chama_Tela("CategoriaFornec")

        Case MENU_CP_CAD_TA_PAGTO_PERIODICO
            Call Chama_Tela("PagtosPeriodicos")

        Case MENU_CP_CAD_TA_COBREMAILPADRAO
            Call Chama_Tela("ModelosEmail")
    
    End Select

End Sub

Private Sub mnuCPCon_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index

        Case MENU_CP_CON_FORNECEDOR
            Call Chama_Tela("FornecedorConsulta")
        
'        Case MENU_CP_CON_TITPAG
'            Call Chama_Tela("TitulosPagarLista", colSelecao)
'
'        Case MENU_CP_CON_TITPAG_TODOS
'            Call Chama_Tela("TitPagTodosLista")
'
'        Case MENU_CP_CON_NFISC
'            Call Chama_Tela("NFPagLista_Consulta", colSelecao)
'
'        Case MENU_CP_CON_NFISCAL_TODOS
'            Call Chama_Tela("NFPagTodosLista")

        Case MENU_CP_CON_CREDITOPRE
            Call Chama_Tela("CredPagarLista_Consulta", colSelecao)

        Case MENU_CP_CON_PAGANTECIP
            Call Chama_Tela("AntecipPagLista_Consulta")

        Case MENU_CP_CON_PAGAMENTOS
            Call Chama_Tela("PagtoLista_Consulta", colSelecao)

'        Case MENU_CP_CON_TITPAG_TODOS_TF
'            Call Chama_Tela("TitPagTodosTFLista")
'
'        Case MENU_CP_CON_NFISCAL_TODOS_TF
'            Call Chama_Tela("NFPagEmpTodaLista", colSelecao)
'
'        Case MENU_CP_CON_COMISSAO_CARTAO
'            Call Chama_Tela("NFPagEmpTodaLista", colSelecao, Nothing, Nothing, "NumIntTitPag = 0 AND NumIntDoc IN (SELECT NumIntDocOrigem FROM TRVTitulos WHERE TipoDocOrigem =5 AND TipoDoc LIKE '%CMCC%')")
'
'        Case MENU_CP_CON_BAIXASPAG
'            Call Chama_Tela("BaixasPagLista")
'
'        Case MENU_CP_CON_NOTAFATPAGTO
'            Call Chama_Tela("NotaFaturaPagtoLista")
'
'        Case MENU_CP_CON_TITPAG_NAKA
'            Call Chama_Tela("NakaTit_ParcelasPagLista")
            
    End Select

End Sub

Private Sub mnuCPConNF_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
                    
        Case MENU_CP_CON_NF_ABERTO
            Call Chama_Tela("NFPagLista_Consulta", colSelecao)
        
        Case MENU_CP_CON_NF_TODOS
            Call Chama_Tela("NFPagTodosLista")

        Case MENU_CP_CON_NF_TODOS_ET
            Call Chama_Tela("NFPagEmpTodaLista", colSelecao)

        Case MENU_CP_CON_NF_CMCC_SEM_FAT
            Call Chama_Tela("NFPagEmpTodaLista", colSelecao, Nothing, Nothing, "NumIntTitPag = 0 AND NumIntDoc IN (SELECT NumIntDocOrigem FROM TRVTitulos WHERE TipoDocOrigem =5 AND TipoDoc LIKE '%CMCC%')")

        Case MENU_CP_CON_NF_FAT_BX
            Call Chama_Tela("NotaFaturaPagtoLista")
    
    End Select

End Sub

Private Sub mnuCRConCHQ_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
                    
        Case MENU_CR_CON_CHQ_EMITIDOS
            Call Chama_Tela("ChequesCRLista")
        
        Case MENU_CR_CON_CHQ_PARCELAS
            Call Chama_Tela("ChequesParcCRLista")

    End Select

End Sub

Private Sub mnuCPConCHQ_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
                    
        Case MENU_CP_CON_CHQ_EMITIDOS
            Call Chama_Tela("ChequesCPLista")
        
        Case MENU_CP_CON_CHQ_COMP
            Call Chama_Tela("ChequesCPCompLista")
        
        Case MENU_CP_CON_CHQ_PARCELAS
            Call Chama_Tela("ChequesPagParcLista")

    End Select

End Sub

Private Sub mnuCPConTP_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index

        Case MENU_CP_CON_TP_ABERTO
            Call Chama_Tela("TitulosPagarLista", colSelecao)
        
        Case MENU_CP_CON_TP_TODOS
            Call Chama_Tela("TitPagTodosLista")

        Case MENU_CP_CON_TP_TODOS_ET
            Call Chama_Tela("TitPagTodosTFLista")

        Case MENU_CP_CON_TP_ATRASADOS
        
            colSelecao.Add gdtDataAtual
            colSelecao.Add 0
            
            Call Chama_Tela("TitPagTodosTFLista", colSelecao, Nothing, Nothing, "DataVencimentoReal < ? AND SaldoParc > ?")
            
        Case MENU_CP_CON_TP_BAIXAS
            Call Chama_Tela("BaixasPagLista")

        Case MENU_CP_CON_TP_NAKA
            Call Chama_Tela("NakaTit_ParcelasPagLista")
            
    End Select

End Sub

Private Sub mnuCPMov_Click(Index As Integer)

Dim objBorderoPagEmissao As New ClassBorderoPagEmissao
Dim objChequesPag As New ClassChequesPag
Dim objChequesPagAvulso As New ClassChequesPagAvulso

    Select Case Index

        Case MENU_CP_MOV_NFISCFATURA
            Call Chama_Tela("NFFATPAG")

        Case MENU_CP_MOV_NFISCAIS
            Call Chama_Tela("NFPag")

        Case MENU_CP_MOV_FATURAS
            Call Chama_Tela("FaturasPag")

        Case MENU_CP_MOV_OUTROSTITULOS
            Call Chama_Tela("OutrosPag")

        Case MENU_CP_MOV_CONFIRMACOBRANCA
            Call Chama_Tela("DetPag")

        Case MENU_CP_MOV_BORDEROPAG
            Call Chama_Tela("BorderoPag1", objBorderoPagEmissao)

        Case MENU_CP_MOV_CHEQUEAUTO
            Call Chama_Tela("ChequesPag", objChequesPag)

        Case MENU_CP_MOV_CHEQUEMANUAIS
            Call Chama_Tela("ChequePagAvulso1", objChequesPagAvulso)

        Case MENU_CP_MOV_CANCELARPAG
            Call Chama_Tela("PagtoCancelar")

        Case MENU_CP_MOV_DEVCREDITOS
            Call Chama_Tela("CreditosPagar")

        Case MENU_CP_MOV_ADIANTAMENTOFORN
            Call Chama_Tela("AntecipPag")

        Case MENU_CP_MOV_COMPENSAR_CHEQUE_PRE
            Call Chama_Tela("ChequePrePag")

        Case MENU_CP_MOV_LIBERA_PAGTO
            Call Chama_Tela("LiberaPagto")

    End Select

End Sub

Private Sub mnuCPRel_Click(Index As Integer)

Dim lErro As Long
Dim iCancela As Integer
Dim sCodRel As String
Dim objRelSel As New AdmRelSel
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_mnuCPRel_Click

    Select Case Index

        Case MENU_CP_REL_TITPAGAR
            objRelatorio.Rel_Menu_Executar ("Títulos a Pagar")
        
        Case MENU_CP_REL_POSFORN
            objRelatorio.Rel_Menu_Executar ("Posição dos Fornecedores")
        
        Case MENU_CP_REL_CHEQUES
            objRelatorio.Rel_Menu_Executar ("Relação de Cheques Emitidos")
        
        Case MENU_CP_REL_BAIXAS
            objRelatorio.Rel_Menu_Executar ("Relação de Baixas no Contas a Pagar")
        
        Case MENU_CP_REL_PAGCANC
            objRelatorio.Rel_Menu_Executar ("Pagamentos Cancelados")

        Case MENU_CP_REL_OUTROS
            lErro = Chama_Tela("RelSelecionar", objRelSel, MODULO_CONTASAPAGAR)
            If (lErro <> SUCESSO) Then Error 7714

            If (objRelSel.iCancela <> SUCESSO) Then Exit Sub

            'Prosseguir executando relatório
            objRelatorio.Rel_Menu_Executar (objRelSel.sCodRel)

        Case MENU_CP_REL_GERREL
            Sistema_EditarRel ("")

        '####################################
        'Inserido por Wagner
        Case MENU_CP_REL_PLANILHAS
            lErro = Chama_Tela("PlanilhasSelecionar", MODULO_CONTASAPAGAR)
            If (lErro <> 0) Then Error 7714
        '####################################
        
    End Select

    Exit Sub

Erro_mnuCPRel_Click:

    Select Case Err

        Case 7714

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165239)

    End Select

    Exit Sub

End Sub

Private Sub mnuCPRelCad_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_CP_REL_CAD_FORN
            objRelatorio.Rel_Menu_Executar ("Relação de Fornecedores")

        Case MENU_CP_REL_CAD_TIPOFORN
            objRelatorio.Rel_Menu_Executar ("Relação de Tipos de Fornecedor")

        Case MENU_CP_REL_CAD_PORTADORES
            objRelatorio.Rel_Menu_Executar ("Relação de Portadores")

        Case MENU_CP_REL_CAD_CONDPAGTO
            objRelatorio.Rel_Menu_Executar ("Relação de Condições de Pagamento")

    End Select
    
End Sub

Private Sub mnuCPRot_Click(Index As Integer)

    Select Case Index
    
        Case MENU_TES_CP_REMESSA_PAGTO
            Call Chama_Tela("BorderoPagSeleciona")
    
        Case MENU_TES_CP_ENVIO_EMAIL_COBR_FAT
            Call Chama_Tela("CobrancaFaturaPorEmail")
    
        Case MENU_TES_CP_ENVIO_EMAIL_AVISO_PAGTO
            Call Chama_Tela("AvisoPagtoCPPorEmail")
    
        Case MENU_CP_ROT_IMPORTACAOCF_REIZA
            Call Chama_Tela("ImportaCartaFrete")
    
        Case MENU_CP_ROT_IMPORTACAO_RETPAGTO
            Call Chama_Tela("ImportarRetPagto")
    
    End Select
    
End Sub

Private Sub mnuCRCad_Click(Index As Integer)

Dim objLote As ClassLote

    Select Case Index

        Case MENU_CR_CAD_CLIENTES
            Call Chama_Tela("Clientes")

        Case MENU_CR_CAD_COBRADORES
            Call Chama_Tela("Cobradores")

        Case MENU_CR_CAD_VENDEDORES
            Call Chama_Tela("Vendedores")

        Case MENU_CR_CAD_TRANSPORTADORAS
            Call Chama_Tela("Transportadora")

        Case MENU_CR_CAD_LOTES
            Call Chama_Tela("LoteTela", objLote, MODULO_CONTASARECEBER)

    End Select

End Sub

Private Sub mnuCPConfig_Click(Index As Integer)
'Alterado por Wagner

Dim objForm As Form
Dim lErro As Long

On Error GoTo Erro_mnuCPConfig_Click

    Select Case Index

        Case MENU_CP_CONFIG_CONFIGURACAO
            Call Chama_Tela("ConfiguraCP")
            
        Case MENU_CP_CONFIG_SEGMENTOS
            
            For Each objForm In Forms
                If Not (objForm Is Me) And Not (objForm Is gobjEstInicial) Then gError 136668
            Next

            Call Chama_Tela_Modal("SegmentosMAT")
            
    End Select

    Exit Sub
    
Erro_mnuCPConfig_Click:

    Select Case Err

        Case 136668
            Call Rotina_Erro(vbOKOnly, "ERRO_FECHAR_JANELAS_FILHAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165240)

    End Select

    Exit Sub

End Sub

Private Sub mnuCRCadTA_Click(Index As Integer)

    Select Case Index

        Case MENU_CR_CAD_TA_CATEGORIACLI
            Call Chama_Tela("CategoriaCliente")

        Case MENU_CR_CAD_TA_TIPOSCLI
            Call Chama_Tela("TipoCliente")

        Case MENU_CR_CAD_TA_TIPOSVEND
            Call Chama_Tela("TipoVendedor")

        Case MENU_CR_CAD_TA_REGIOESVEND
            Call Chama_Tela("RegiaoVenda")

        Case MENU_CR_CAD_TA_CONDPAG
            Call Chama_Tela("CondicoesPagto")

        Case MENU_CR_CAD_TA_CARTEIRACOBRANCA
            Call Chama_Tela("CarteirasCobranca")

        Case MENU_CR_CAD_TA_PADROESCOBRANCA
            Call Chama_Tela("PadroesCobranca")

        Case MENU_CR_CAD_TA_PAISES
            Call Chama_Tela("Paises")

        Case MENU_CR_CAD_TA_MENSAGENS
            Call Chama_Tela("Mensagens")

        Case MENU_CR_CAD_TA_HISTEXTRATO
            Call Chama_Tela("HistMovCta")

        Case MENU_CR_CAD_TA_COMISSAVULSA
            Call Chama_Tela("ComissaoAvulsa")

        Case MENU_CR_CAD_TA_RECEB_PERIODICO
            Call Chama_Tela("RecebPeriodicos")

        Case MENU_CR_CAD_TA_ARQRET_DETALHE
            Call Chama_Tela("TiposDetRetCobr")

        Case MENU_CR_CAD_TA_ARQRET_DIF
            Call Chama_Tela("TiposDifParcRec")
            
        Case MENU_CR_CAD_TA_COBREMAILPADRAO
            Call Chama_Tela("ModelosEmail")

    End Select

End Sub

Private Sub mnuCRCon_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index

''        Case MENU_CR_CON_CHEQUEPRE
''            Call Chama_Tela("ChequePreLista_Consulta", colSelecao)
''
        Case MENU_CR_CON_DEBREC
            Call Chama_Tela("DebitosRecebLista_Consulta", colSelecao)
''
''        Case MENU_CR_CON_MENSAGENS
''            Call Chama_Tela("MensagemLista")
''
''        Case MENU_CR_CON_PADRAOCOBRANCA
''            Call Chama_Tela("PadraoCobrancaLista")
''
''        Case MENU_CR_CON_PLANOCONTAS
''            Call Chama_Tela("PlanoContaCRLista", colSelecao)
''
        Case MENU_CR_CON_RECANTECIP
            Call Chama_Tela("AntecipRecebLista_Consulta")

''        Case MENU_CR_CON_TIPOCLI
''            Call Chama_Tela("TipoClienteLista")
''
''        Case MENU_CR_CON_TIPOVEND
''            Call Chama_Tela("TipoVendedorLista")
''
        Case MENU_CR_CON_CLICONS
            Call Chama_Tela("ClienteConsulta")

'        Case MENU_CR_CON_TITULOSREC
'            Call Chama_Tela("TitulosReceberLista", colSelecao)

'        Case MENU_CR_CON_TITULOSREC_TODOS
'            Call Chama_Tela("TitRecTodosLista")

'        Case MENU_CR_CON_TITULOSREC_BAIXADOS
'            Call Chama_Tela("TitRecTodosBaixadosLista")

        Case MENU_CR_CON_BORDCOBRANCA
            Call Chama_Tela("BordCobrancaLista")

'        Case MENU_CR_CON_VOUCHER
'            If glEmpresa = 1 Then
'                colSelecao.Add "TVA"
'                If giFilialEmpresa <> EMPRESA_TODA And giFilialEmpresa <> Abs(giFilialAuxiliar) Then
'                    colSelecao.Add giFilialEmpresa
'                    Call Chama_Tela("CoInfoItemFaturaImportLista", colSelecao, Nothing, Nothing, "CodEST = ? AND FilialEmpresa = ?")
'                Else
'                    Call Chama_Tela("CoInfoItemFaturaImportLista", colSelecao, Nothing, Nothing, "CodEST = ?")
'                End If
'            Else
'                colSelecao.Add "TVI"
'                If giFilialEmpresa <> EMPRESA_TODA And giFilialEmpresa <> Abs(giFilialAuxiliar) Then
'                    colSelecao.Add giFilialEmpresa
'                    Call Chama_Tela("CoInfoItemFaturaImportLista", colSelecao, Nothing, Nothing, "CodEST = ? AND FilialEmpresa = ?")
'                Else
'                    Call Chama_Tela("CoInfoItemFaturaImportLista", colSelecao, Nothing, Nothing, "CodEST = ? ")
'                End If
'            End If
'            Call Chama_Tela("TRVVouFatCCLista")
'
'        Case MENU_CR_CON_TITULOSREC_TODOS_TF
'            Call Chama_Tela("TitRecTodosTFLista")
'
'        Case MENU_CR_CON_TITULOSRECEBER_TF
'            Call Chama_Tela("TituloReceberTFLista")
'
'        Case MENU_CR_CON_PARCELASRECEBER_TF
'            Call Chama_Tela("TituloReceberTF2Lista")
'
'        Case MENU_CR_CON_BAIXASREC
'            Call Chama_Tela("BaixasRecLista")
            
        Case MENU_CR_CON_PRODCR
            Call Chama_Tela("ProdCRLista")
    
    
    End Select

End Sub

Private Sub mnuCRConTR_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index

        Case MENU_CR_CON_TR_ABERTO
            Call Chama_Tela("TitulosReceberLista", colSelecao)

        Case MENU_CR_CON_TR_TODOS
            Call Chama_Tela("TitRecTodosLista")

        Case MENU_CR_CON_TR_BAIXADOS
            Call Chama_Tela("TitRecTodosBaixadosLista")

        Case MENU_CR_CON_TR_CARTAO_VOUCHER
            Call Chama_Tela("TRVVouFatCCLista")
            
        Case MENU_CR_CON_TR_TODOS_ET
            Call Chama_Tela("TitRecTodosTFLista")
                       
        Case MENU_CR_CON_TR_BAIXADOS_ET
            Call Chama_Tela("TitRecTodosTFLista", colSelecao, Nothing, Nothing, "Status = 2")
                       
        Case MENU_CR_CON_TR_ABERTO_ET
            Call Chama_Tela("TitRecTodosTFLista", colSelecao, Nothing, Nothing, "Status <> 2")
            
        Case MENU_CR_CON_TR_BAIXAS
            Call Chama_Tela("BaixasRecLista")
            
        Case MENU_CR_CON_TR_ATRASADOS
        
            colSelecao.Add gdtDataAtual
            colSelecao.Add 0
            
            Call Chama_Tela("TitRecTodosTFLista", colSelecao, Nothing, Nothing, "DataVencimentoReal < ? AND SaldoParc > ?")
            
        Case MENU_CR_CON_TR_BOLETO_CANC
            Call Chama_Tela("BoletosCancParcLista")
            
        Case MENU_CR_CON_TR_COMISSOES_BX
            Call Chama_Tela("TitRecNFPVComissLista")
            
    End Select

End Sub

Private Sub mnuCRConfig_Click(Index As Integer)
'Alterado por Wagner

Dim objForm As Form
Dim lErro As Long

On Error GoTo Erro_mnuCRConfig_Click

    Select Case Index

        Case MENU_CR_CONFIG_CONFIGURACAO
            Call Chama_Tela("ConfiguraCR")
            
        Case MENU_CR_CONFIG_COBRANCAELETRONICA
            Call Chama_Tela("BancosInfo")
            
        Case MENU_CR_CONFIG_SEGMENTOS
            
            For Each objForm In Forms
                If Not (objForm Is Me) And Not (objForm Is gobjEstInicial) Then gError 136667
            Next

            Call Chama_Tela_Modal("SegmentosMAT")
            
    End Select

    Exit Sub
    
Erro_mnuCRConfig_Click:

    Select Case Err

        Case 136667
            Call Rotina_Erro(vbOKOnly, "ERRO_FECHAR_JANELAS_FILHAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165241)

    End Select

    Exit Sub
    
End Sub

Private Sub mnuCRMov_Click(Index As Integer)

    Select Case Index
    
        Case MENU_CR_MOV_BORDERODESCCHQEXCLUI
            Call Chama_Tela("BorderoDescChqExclui")
    
        Case MENU_CR_MOV_BORDERODESCCHQ
            Call Chama_Tela("BorderoDescChq1")

        Case MENU_CR_MOV_TITULOSREC
            Call Chama_Tela("TituloReceber")

        Case MENU_CR_MOV_CHEQUEPRE
            Call Chama_Tela("ChequePre")

        Case MENU_CR_MOV_BORDEROCHEQUEPRE
            Call Chama_Tela("BorderoChequesPre")

        Case MENU_CR_MOV_BORDEROCOBRANCA
            Call Chama_Tela("BorderoCobranca")

        Case MENU_CR_MOV_INSTCOBELETRONICA
            Call Chama_Tela("AlteracoesCobranca")

        Case MENU_CR_MOV_CANCELARBORDEROCOB
            Call Chama_Tela("BorderoCobrancaCancelar")

        Case MENU_CR_MOV_TRANFERENCIAMANUAL
            Call Chama_Tela("TransfCartCobr")

        Case MENU_CR_MOV_DEVOLUCOESENTRADA
            Call Chama_Tela("DebitosReceb")

        Case MENU_CR_MOV_ADIANTAMENTOCLI
            Call Chama_Tela("AntecipReceb")

        Case MENU_CR_MOV_DEVOLUCAOCHEQUE
            Call Chama_Tela("DevolucaoCheque")
        
        Case MENU_CR_MOV_BORDEROCHQPREEXCLUI
            Call Chama_Tela("BorderoChqPreExclui")

        Case MENU_CR_MOV_COBRANCA
            Call Chama_Tela("Cobranca")
        
        Case MENU_CR_MOV_HISTORICOCLIENTE
            Call Chama_Tela("HistoricoCliente")
        
'        Case MENU_CR_MOV_PARCRECDIF
'            Call Chama_Tela("ParcelasRecDif")

    End Select

End Sub

Private Sub mnuCRMRot_Click(Index As Integer)

    Select Case Index

        Case MENU_CRM_ROT_ENVIOEMAILCLI
            Call Chama_Tela("ContatoClientePorEmail")
            
    End Select
End Sub

Private Sub mnuCRRel_Click(Index As Integer)

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim iCancela As Integer
Dim sCodRel As String
Dim objRelSel As New AdmRelSel

On Error GoTo Erro_mnuCRRel_Click

    Select Case Index

        Case MENU_CR_REL_TITREC
            objRelatorio.Rel_Menu_Executar ("Títulos a Receber")
        
        Case MENU_CR_REL_TITATRASO
            objRelatorio.Rel_Menu_Executar ("Títulos em Atraso")
        
        Case MENU_CR_REL_POSGERCOB
            objRelatorio.Rel_Menu_Executar ("Posição Geral da Cobrança")

        Case MENU_CR_REL_POSCLI
            objRelatorio.Rel_Menu_Executar ("Posição do Cliente")

        Case MENU_CR_REL_BAIXAS
            objRelatorio.Rel_Menu_Executar ("Relação de Baixas no Contas a Receber")
        
        Case MENU_CR_REL_COMISSOESVEND
            objRelatorio.Rel_Menu_Executar ("Relatório de Comissões")

        Case MENU_CR_REL_COMISSOESPAG
            objRelatorio.Rel_Menu_Executar ("Resumo de Comissões a Pagar")

        Case MENU_CR_REL_MAIORESDEV
            objRelatorio.Rel_Menu_Executar ("Maiores Devedores")

        Case MENU_CR_REL_TITTELEF
            objRelatorio.Rel_Menu_Executar ("Títulos para cobrança por telefone")
        
        Case MENU_CR_REL_TITMALA
            objRelatorio.Rel_Menu_Executar ("Títulos para cobrança via mala direta")

        Case MENU_CR_REL_OUTROS
            lErro = Chama_Tela("RelSelecionar", objRelSel, MODULO_CONTASARECEBER)
            If (lErro <> SUCESSO) Then Error 7715

            If (objRelSel.iCancela <> 0) Then Exit Sub

            'Prosseguir executando relatório
            objRelatorio.Rel_Menu_Executar (objRelSel.sCodRel)

        Case MENU_CR_REL_GERREL
            Sistema_EditarRel ("")

        '####################################
        'Inserido por Wagner
        Case MENU_CR_REL_PLANILHAS
            lErro = Chama_Tela("PlanilhasSelecionar", MODULO_CONTASARECEBER)
            If (lErro <> 0) Then Error 7715
        '####################################
        
    End Select

    Exit Sub

Erro_mnuCRRel_Click:

    Select Case Err

        Case 7715

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165242)

    End Select

    Exit Sub

End Sub

Private Sub mnuCTBRotSped_Click(Index As Integer)

    Select Case Index
    
        Case MENU_CTB_ROT_SPED_DIARIO
            Call Chama_Tela("SpedDiario")
        
        Case MENU_CTB_ROT_SPED_FCONT
            Call Chama_Tela("SpedFCont")
        
    End Select

End Sub

Private Sub mnuESTCadLoteRastroLoc_Click()
    Call Chama_Tela("RastreamentoLoteLoc")
End Sub

Private Sub mnuESTMovOP_Click(Index As Integer)
    Select Case Index
        Case 1
            Call Chama_Tela("OrdemProducao")
        Case 2
            Call Chama_Tela("GeracaoOP")
        Case 3
            Call Chama_Tela("Empenho")
    End Select
End Sub

Private Sub mnuFATCadFP_Click(Index As Integer)

    Select Case Index

        Case MENU_FAT_CAD_FP_CUSTOEMBMP
            Call Chama_Tela("CustoEmbMP")
        
        Case MENU_FAT_CAD_FP_CUSTODIRPROD
            Call Chama_Tela("CustoDiretoProd")
        
        Case MENU_FAT_CAD_FP_CUSTOFIXOPROD
            Call Chama_Tela("CustoFixoProd")
        
        Case MENU_FAT_CAD_FP_DVVCLIENTE
            Call Chama_Tela("DVVCliente")
        
        Case MENU_FAT_CAD_FP_TIPOFRETE
            Call Chama_Tela("TipoFreteFP")
        
    End Select
    
End Sub

Private Sub MnuFATCadTransp_Click(Index As Integer)

    Select Case Index
    
        Case MENU_FAT_CAD_TRANSP_DESPACHANTE
            Call Chama_Tela("Despachante")
            
        Case MENU_FAT_CAD_TRANSP_NAVIO
            Call Chama_Tela("ProgNavio")
            
        Case MENU_FAT_CAD_TRANSP_TABELAPRECO
            Call Chama_Tela("TabPreco")
            
        Case MENU_FAT_CAD_TRANSP_ITENSSERVICO
            Call Chama_Tela("ItemServico")
            
        Case MENU_FAT_CAD_TRANSP_TIPOCARGA
            Call Chama_Tela("TipoEmbalagem")
    
        Case MENU_FAT_CAD_TRANSP_ORIGEMDESTINO
            Call Chama_Tela("OrigemDestino")
    
        Case MENU_FAT_CAD_TRANSP_TIPOCONTAINER
            Call Chama_Tela("TipoContainer")
        
        Case MENU_FAT_CAD_TRANSP_SERVITEMSERV
            Call Chama_Tela("ServItemServ")
        
        Case MENU_FAT_CAD_TRANSP_DOCUMENTO
            Call Chama_Tela("Documento")
            
    End Select
        
End Sub

Private Sub mnuFATCon_Click(Index As Integer)
    
    Select Case Index
        
        Case MENU_FAT_CON_LOG
            Call Chama_Tela("LogWFWLista")
        
        Case MENU_FAT_CON_CONTROLENF
            Call Chama_Tela("NFiscalControleNF1Lista")
        
    End Select
    
End Sub

Private Sub mnuFATConfigFP_Click(Index As Integer)

    Select Case Index
    
        Case MENU_FAT_CONFIG_FP_MNEMONICOS
            Call Chama_Tela("MnemonicoFPPlanilha")
        
        Case MENU_FAT_CONFIG_FP_PLANILHAS
            Call Chama_Tela("Planilhas")
        
        Case MENU_FAT_CONFIG_FP_MARGCONTR
            Call Chama_Tela("PlanMargContrConfig")
            
    End Select

End Sub

Private Sub mnuFATMov2_Click(Index As Integer)
    Select Case Index
        Case 5
            Call Chama_Tela("OrcamentoVenda")
        Case 10
            Call Chama_Tela("OVAcompanhamento")
        Case 15
            Call Chama_Tela("PedidoVenda")
        Case 20
            Call Chama_Tela("LiberaBloqueio")
        Case 25
            Call Chama_Tela("BaixaPedido")
    End Select
End Sub

Private Sub mnuFATMovTransp_Click(Index As Integer)

    Select Case Index

        Case MENU_FAT_MOV_TRANSP_PEDIDOCOTACAO
            Call Chama_Tela("PropostaCotacao")
    
        Case MENU_FAT_MOV_TRANSP_PROPOSTACOTACAO
            Call Chama_Tela("Cotacao")
    
        Case MENU_FAT_MOV_TRANSP_SOLICITACAOSERVICO
            Call Chama_Tela("SolicitacaoServico")
    
        Case MENU_FAT_MOV_TRANSP_COMPROVANTESERVICO
            Call Chama_Tela("CompServico")
    
        Case MENU_FAT_MOV_TRANSP_CONHECIMENTOFRETE
            Call Chama_Tela("ConhecimentoFrete")
    
        Case MENU_FAT_MOV_TRANSP_CONHECIMENTOFRETEFAT
            Call Chama_Tela("ConhecimentoFreteFatura")

    End Select

End Sub

Private Sub mnuFATRelTRV_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_FAT_REL_TRV_ATEND
            objRelatorio.Rel_Menu_Executar ("Resumo no Período por Atendente")
            
        Case MENU_FAT_REL_TRV_EST
            objRelatorio.Rel_Menu_Executar ("Estatísticas no Período")
            
        Case MENU_FAT_REL_TRV_DESV
            objRelatorio.Rel_Menu_Executar ("Desvios de Venda por Cliente")
            
        Case MENU_FAT_REL_TRV_VOU
            objRelatorio.Rel_Menu_Executar ("Vouchers Emitidos")
            
        Case MENU_FAT_REL_TRV_ACOM_INAD
            objRelatorio.Rel_Menu_Executar ("Acompanhamento de inadimplência")
            
        Case MENU_FAT_REL_TRV_POS_INAD
            objRelatorio.Rel_Menu_Executar ("Posição de inadimplência")
            
        Case MENU_FAT_REL_TRV_IMP
            objRelatorio.Rel_Menu_Executar ("Log de Importação\Atualização dos dados da Coinfo")

    End Select
    
End Sub

Private Sub mnuFATRotFP_Click(Index As Integer)

    Select Case Index
        
        Case MENU_FAT_ROT_FP_MARGCONTR
            Call Chama_Tela("MargContr")
        
        Case MENU_FAT_ROT_FP_RATEIOCUSTODIR
            Call Chama_Tela("CustoDirFabrCalcula")
            
        Case MENU_FAT_ROT_FP_RATEIOCUSTOFIXO
            Call Chama_Tela("CustoFixo")
        
        Case MENU_FAT_ROT_FP_CALCPRECOS
            Call Chama_Tela("CalcVend")
        
        Case MENU_FAT_ROT_FP_AJUSTEPRECO
            Call Chama_Tela("CustoTabelaPreco")
        
    End Select
    
End Sub

Private Sub mnuFISMovPIS_Click(Index As Integer)
    Select Case Index
        
        Case MENU_FIS_MOV_PIS_APURACAO
            Call Chama_Tela("ApuracaoPisCofins", Nothing, 1)
    End Select
End Sub

Private Sub mnuFISMovCOFINS_Click(Index As Integer)
    Select Case Index
        
        Case MENU_FIS_MOV_COFINS_APURACAO
            Call Chama_Tela("ApuracaoPisCofins", Nothing, 2)
            
    End Select
End Sub


Private Sub mnuLJConCad_Click(Index As Integer)

    Select Case Index
    
        Case MENU_LJ_CON_CAD_CLIENTES
            If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                Call Chama_Tela("ClientesLista")
            Else
                Call Chama_Tela("ClientesLista")
            End If
            
        Case MENU_LJ_CON_CAD_PRODUTOS
            If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                Call Chama_Tela("ProdutoLista_Consulta")
            Else
                Call Chama_Tela("ProdutoLista_Consulta")
            End If
            
        Case MENU_LJ_CON_CAD_OPERADOR
            If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then
                Call Chama_Tela("OperadorLista")
            Else
                Call Chama_Tela("OperadorLista")
            End If
            
        Case MENU_LJ_CON_CAD_CAIXA
            If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then
                Call Chama_Tela("CaixaLista")
            Else
                Call Chama_Tela("CaixaLista")
            End If
            
        Case MENU_LJ_CON_CAD_VENDEDOR
            If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                Call Chama_Tela("VendedorLista")
            Else
                Call Chama_Tela("VendedorLista")
            End If

        Case MENU_LJ_CON_CAD_ECF
            If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then
                Call Chama_Tela("ECFLista")
            Else
                Call Chama_Tela("ECFLista")
            End If

        Case MENU_LJ_CON_CAD_PRECOS
            If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                Call Chama_Tela("TabelaPrecoItemLista")
            Else
                Call Chama_Tela("TabelaPrecoItemLista")
            End If
            
        Case MENU_LJ_CON_CAD_MEIOSPAGTO
            If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then
                Call Chama_Tela("AdmMeioPagtoLista")
            Else
                Call Chama_Tela("AdmMeioPagtoLista")
            End If
            
        Case MENU_LJ_CON_CAD_REDES
            If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then
                Call Chama_Tela("RedesLista")
            Else
                Call Chama_Tela("RedesLista")
            End If
            
'        Case MENU_LJ_CON_CAD_CUPOMFISCAL
'            If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then
'                Call Chama_Tela("CupomFiscalLista")
'            Else
'                Call Chama_Tela("CupomFiscalLista")
'            End If
'
'        Case MENU_LJ_CON_CAD_ITEMCUPOMFISCAL
'            If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then
'                Call Chama_Tela("ItensCupomFiscalLista")
'            Else
'                Call Chama_Tela("ItensCupomFiscalLista")
'            End If
            
            
    End Select
    
End Sub

'Luiz Nogueira  31/03/04
Private Sub mnuLJRel_Click(Index As Integer)

Dim objRelSel As New AdmRelSel
Dim objRelatorio As New AdmRelatorio
Dim lErro As Long

    Select Case Index

        Case MENU_LJ_REL_OUTROS
            lErro = Chama_Tela("RelSelecionar", objRelSel, MODULO_LOJA)
            If (lErro <> 0) Then Error 7058

            If (objRelSel.iCancela <> 0) Then Exit Sub

            'Prosseguir executando relatório
            objRelatorio.Rel_Menu_Executar (objRelSel.sCodRel)


        Case MENU_LJ_REL_GERREL
            Sistema_EditarRel ("")

        '####################################
        'Inserido por Wagner
        Case MENU_LJ_REL_PLANILHAS
            lErro = Chama_Tela("PlanilhasSelecionar", MODULO_LOJA)
            If (lErro <> 0) Then Error 7058
        '####################################
        
    End Select

End Sub

'Luiz Nogueira  31/03/04
Private Sub mnuLJRelCad_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index
        
        Case MENU_LJ_REL_CAD_CLI
            objRelatorio.Rel_Menu_Executar ("Relação de Clientes")
        
        Case MENU_LJ_REL_CAD_PRODUTOS
            objRelatorio.Rel_Menu_Executar ("Relação de Produtos")
        
        Case MENU_LJ_REL_CAD_OPERADORES
            objRelatorio.Rel_Menu_Executar ("Relação de Operadores")
        
        Case MENU_LJ_REL_CAD_CAIXAS
            objRelatorio.Rel_Menu_Executar ("Relação de Caixas")

        Case MENU_LJ_REL_CAD_ECFS
            objRelatorio.Rel_Menu_Executar ("Relação de Equipamentos Emissores de Cupom Fiscal")
    
    End Select
    
End Sub

'Luiz Nogueira  31/03/04
Private Sub mnuLJRelCx_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_LJ_REL_CAIXA_MOVCAIXAS
            objRelatorio.Rel_Menu_Executar ("Relatorio de Movimentos de Caixa")

        Case MENU_LJ_REL_CAIXA_PAINELCAIXAS
            objRelatorio.Rel_Menu_Executar ("Painel de Caixas")
        
        Case MENU_LJ_REL_CAIXA_ORCAMENTOSLOJA
            objRelatorio.Rel_Menu_Executar ("Relação de Orçamentos emitidos em ECF")

        Case MENU_LJ_REL_CAIXA_CUPONSFISCAIS
            objRelatorio.Rel_Menu_Executar ("Relação de Cupons Fiscais")

    End Select
    
End Sub

'Luiz Nogueira  31/03/04
Private Sub mnuLJRelVen_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_LJ_REL_VEN_EVOLVENDAS
            objRelatorio.Rel_Menu_Executar ("Relatorio de vendas")
        
        Case MENU_LJ_REL_VEN_MAPAVENDAS
            objRelatorio.Rel_Menu_Executar ("Lista as vendas de Produtos")
        
        Case MENU_LJ_REL_VEN_FLASHVENDAS
            objRelatorio.Rel_Menu_Executar ("Relatorio de Flash de Vendas")

        Case MENU_LJ_REL_VEN_VENDASXMEIOPAGTO
            objRelatorio.Rel_Menu_Executar ("Vendas x Meios de Pagamento")

        Case MENU_LJ_REL_VEN_RANKINGPRODUTOS
            objRelatorio.Rel_Menu_Executar ("Relatorio de Ranking de Produto")

        Case MENU_LJ_REL_VEN_PRODDEVTROCA
            objRelatorio.Rel_Menu_Executar ("Produtos Devolvidos em Trocas")

    End Select
    
End Sub

'Luiz Nogueira  31/03/04
Private Sub mnuLJRelBord_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_LJ_REL_BORD_BORDEROBOLETO
            objRelatorio.Rel_Menu_Executar ("Borderô Boleto")
        
        Case MENU_LJ_REL_BORD_BORDEROTICKET
            objRelatorio.Rel_Menu_Executar ("Bordero Ticket")
        
        Case MENU_LJ_REL_BORD_BORDEROOUTROS
            objRelatorio.Rel_Menu_Executar ("Relatorio de Flash de Vendas")

    End Select
    
End Sub

Private Sub mnuCRRelCad_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_CR_REL_CAD_CLI
            objRelatorio.Rel_Menu_Executar ("Relação de Clientes")

        Case MENU_CR_REL_CAD_TIPOCLI
            objRelatorio.Rel_Menu_Executar ("Relação de Tipos de Cliente")

        Case MENU_CR_REL_CAD_CATEGCLI
            objRelatorio.Rel_Menu_Executar ("Relação de Categorias de Cliente")

        Case MENU_CR_REL_CAD_COBRADORES
            objRelatorio.Rel_Menu_Executar ("Relação de Cobradores")

        Case MENU_CR_REL_CAD_TIPOSCARTCOB
            objRelatorio.Rel_Menu_Executar ("Relação de Carteiras de Cobrança")

        Case MENU_CR_REL_CAD_PADROESCOB
            objRelatorio.Rel_Menu_Executar ("Relação de Padrões de Cobrança")

        Case MENU_CR_REL_CAD_TIPOSINSTRCOB
            objRelatorio.Rel_Menu_Executar ("Relação dos Tipos de Instruções de Cobrança")

        Case MENU_CR_REL_CAD_VENDEDORES
            objRelatorio.Rel_Menu_Executar ("Relação de Vendedores")

        Case MENU_CR_REL_CAD_TIPOSVEND
            objRelatorio.Rel_Menu_Executar ("Relação de Tipos de Vendedor")

        Case MENU_CR_REL_CAD_REGIOESVEND
            objRelatorio.Rel_Menu_Executar ("Relação de Regiões de Venda")

    End Select
    
End Sub

Private Sub mnuCRRotEmail_Click(Index As Integer)

    Select Case Index
            
        Case MENU_CR_ROT_EMAIL_AVISOCOBR
            Call Chama_Tela("AvisoCobrPorEmail")
            
        Case MENU_CR_ROT_EMAIL_COBR
            Call Chama_Tela("CobrancaPorEmail")
        
        Case MENU_CR_ROT_EMAIL_AGRADECIMENTO
            Call Chama_Tela("AgradecimentoPorEmail")
            
    End Select

End Sub

Private Sub mnuCRRot_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_CR_ROT_ATUALIZAPAGCOMISSAO
            Call Chama_Tela("ComissoesPag")

        Case MENU_CR_ROT_LIMPAARQ
            '???

        Case MENU_CR_ROT_EMISSAO_BOLETOS
            Call Chama_Tela("EmissaoBoletos")
        
        Case MENU_CR_ROT_EMISSAO_DUPLICATAS
            objRelatorio.Rel_Menu_Executar ("Emissão de Duplicatas")
            
        Case MENU_CR_ROT_REAJUSTE_TITULOS
            Call Chama_Tela("ReajusteTitRec")
            
'        Case MENU_CR_ROT_ENVIO_EMAIL_COBR
'            Call Chama_Tela("CobrancaPorEmail")
'
'        Case MENU_CR_ROT_ENVIO_EMAIL_AGRADECIMENTO
'            Call Chama_Tela("AgradecimentoPorEmail")
        
        Case MENU_CR_ROT_IMPORT_TITREC_AF
            Call Chama_Tela("ImportarCRAF")
            
        Case MENU_CR_ROT_EXPORT_ASSOC_AF
            Call Chama_Tela("ExportarClientesAF")
    
        Case MENU_CR_ROT_IMPORTFAT_REIZA
            Call Chama_Tela("ImportaFatura")
    
        Case MENU_CR_ROT_IMPORT_EXT_REDES
            'Call Chama_Tela("RotImpExtRedes")
            objRelatorio.Rel_Menu_Executar ("Importação de Extratos de Redes de Cartão")
    
        Case MENU_CR_ROT_BAIXA_CARTAO
            Call Chama_Tela("BaixaCartao")
    
    End Select

End Sub

Private Sub mnuCRRotRetornoTitulos_Click()
    Call Chama_Tela("ProcessaArqRetCobr")
End Sub

Private Sub mnuCRRotTituloCobranca_Click()
    Call Chama_Tela("GeracaoArqRemCobr")
End Sub

Private Sub mnuCTBCad_Click(Index As Integer)

Dim objLote As ClassLote

    Select Case Index

        Case MENU_CTB_CAD_PLANOCONTA
            Call Chama_Tela("PlanoConta")

        Case MENU_CTB_CAD_CATEGORIA
            Call Chama_Tela("ContaCategoria")

        Case MENU_CTB_CAD_CCL
            Call Chama_Tela("CclTela")

        Case MENU_CTB_CAD_ASSOCCCL
            Call Chama_Tela("ContaCcl")

        Case MENU_CTB_CAD_ASSOCCCLCTB
            Call Chama_Tela("ContaCcl2")

        Case MENU_CTB_CAD_HISTPADRAO
            Call Chama_Tela("HistoricoPadrao")

        Case MENU_CTB_CAD_LOTES
            Call Chama_Tela("LoteTela", objLote)

        Case MENU_CTB_CAD_DOCAUTO
            Call Chama_Tela("DocAuto")

        Case MENU_CTB_CAD_RATEIOOFF
            Call Chama_Tela("RateioOff")

        Case MENU_CTB_CAD_RATEIOON
            Call Chama_Tela("RateioOn")

        Case MENU_CTB_CAD_ORCAMENTO
            Call Chama_Tela("Orcamento")

        Case MENU_CTB_CAD_SALDOINI
            Call Chama_Tela("SaldoInicial")

        Case MENU_CTB_CAD_PADRAOCONTAB
            Call Chama_Tela("PadraoContab")
            
        Case MENU_CTB_CAD_PLANOCONTAREF
            Call Chama_Tela("PlanoContaRef")

    End Select

End Sub

Private Sub mnuCTBCon_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index

        Case MENU_CTB_CON_PLANOCONTA
            Call Chama_Tela("PlanoContaLista")

        Case MENU_CTB_CON_CCL
            Call Chama_Tela("CclLista", colSelecao)

        Case MENU_CTB_CON_CONTACCL
            Call Chama_Tela("ContaCclLista")

        Case MENU_CTB_CON_HISTPADRAO
            Call Chama_Tela("HistPadraoLista")

        Case MENU_CTB_CON_LOTES
            Call Chama_Tela("LoteLista", colSelecao)

        Case MENU_CTB_CON_LOTEPEND
        
            colSelecao.Add giFilialEmpresa
            colSelecao.Add 0
            
            Call Chama_Tela("LotePendenteCTBLista", colSelecao)

        Case MENU_CTB_CON_LAN
            Call Chama_Tela("LancamentoLista", colSelecao)

        Case MENU_CTB_CON_LANPEND

            Call Chama_Tela("LanPendenteLista", colSelecao)

        Case MENU_CTB_CON_DOCAUTO
            Call Chama_Tela("DocAutoLista")

        Case MENU_CTB_CON_ORCAMENTO
            Call Chama_Tela("OrcamentoLista", colSelecao)

        Case MENU_CTB_CON_RATEIOON
            Call Chama_Tela("RateioOnLista")

        Case MENU_CTB_CON_RATEIOOFF
            Call Chama_Tela("RateioOffLista")

        Case MENU_CTB_CON_MAPPLANCTAREF
            Call Chama_Tela("PlnCtaRefConfigLista")

        Case MENU_CTB_CON_ASSOCCTAREF
            Call Chama_Tela("AssociacaoContasRefLista")

    End Select

End Sub

Private Sub mnuCTBConfig_Click(Index As Integer)
Dim objForm As Form
Dim lErro As Long

On Error GoTo Erro_mnuCTBConfig_Click

    Select Case Index

        Case MENU_CTB_CONFIG_EXERCICIO
            Call Chama_Tela("ExercicioTela")

        Case MENU_CTB_CONFIG_EXERCICIOFILIAL
            Call Chama_Tela("ExercicioFilial")

        Case MENU_CTB_CONFIG_CONFIGURACOES
            Call Chama_Tela("Configuracao")

        Case MENU_CTB_CONFIG_SEGMENTOS
            
            For Each objForm In Forms
                If Not (objForm Is Me) And Not (objForm Is gobjEstInicial) Then Error 59360
            Next

            Call Chama_Tela_Modal("SegmentoTela")

        Case MENU_CTB_CONFIG_CAMPOSGLOBAIS
            Call Chama_Tela("MnemonicoGlobal")

    End Select

    Exit Sub
    
Erro_mnuCTBConfig_Click:

    Select Case Err

        Case 59360
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FECHAR_JANELAS_FILHAS", Err, Error)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165243)

    End Select

    Exit Sub

End Sub

Private Sub mnuCTBMov_Click(Index As Integer)

    Select Case Index

        Case MENU_CTB_MOV_LANCALOTE
            Call Chama_Tela("Lancamentos")

        Case MENU_CTB_MOV_LAN
            Call Chama_Tela("LancamentosAt")

        Case MENU_CTB_MOV_ESTORNOLOTE
            Call Chama_Tela("LoteEstorno")

        Case MENU_CTB_MOV_ESTORNODOC
            Call Chama_Tela("LancamentoEstorno")

    End Select

End Sub

Private Sub mnuCTBRel_Click(Index As Integer)

Dim lErro As Long
Dim iCancela As Integer
Dim sCodRel As String
Dim objRelSel As New AdmRelSel
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_mnuCTBRel_Click

    Select Case Index

        Case MENU_CTB_REL_BALANVERIF
            objRelatorio.Rel_Menu_Executar ("Balancete de Verificação")

        Case MENU_CTB_REL_RAZAO
            objRelatorio.Rel_Menu_Executar ("Razão")
        
        Case MENU_CTB_REL_RAZAOAUX
            objRelatorio.Rel_Menu_Executar ("Razão Auxiliar")

        Case MENU_CTB_REL_RAZAOAGLUT
            objRelatorio.Rel_Menu_Executar ("Razão Aglutinado")

        Case MENU_CTB_REL_DIARIO
            objRelatorio.Rel_Menu_Executar ("Diário")

        Case MENU_CTB_REL_DIARIOAUX
            objRelatorio.Rel_Menu_Executar ("Diário Auxiliar")
        
        Case MENU_CTB_REL_DIARIOAGLUT
            objRelatorio.Rel_Menu_Executar ("Diário Aglutinado")

        Case MENU_CTB_REL_BALANPATRI
            objRelatorio.Rel_Menu_Executar ("Balanço Patrimonial")

        Case MENU_CTB_REL_OUTROS

            lErro = Chama_Tela("RelSelecionar", objRelSel, MODULO_CONTABILIDADE)
            If (lErro <> SUCESSO) Then Error 7058

            If (objRelSel.iCancela <> SUCESSO) Then Exit Sub

            'prosseguir executando relatorio
            objRelatorio.Rel_Menu_Executar (objRelSel.sCodRel)

        Case MENU_CTB_REL_GERREL
            Sistema_EditarRel ("")

        '####################################
        'Inserido por Wagner
        Case MENU_CTB_REL_PLANILHAS
            lErro = Chama_Tela("PlanilhasSelecionar", MODULO_CONTABILIDADE)
            If (lErro <> 0) Then Error 7058
        '####################################
        
    End Select

    Exit Sub

Erro_mnuCTBRel_Click:

    Select Case Err

        Case 7058

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165244)

    End Select

    Exit Sub

End Sub

Private Sub mnuCTBRelCad_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_CTB_REL_CAD_PLANOCONTAS
            objRelatorio.Rel_Menu_Executar ("Plano de Contas")

        Case MENU_CTB_REL_CAD_CCL
            objRelatorio.Rel_Menu_Executar ("Centros de Custo")

        Case MENU_CTB_REL_CAD_HISTPADRAO
            objRelatorio.Rel_Menu_Executar ("Históricos Padrão")

        Case MENU_CTB_REL_CAD_LOTESCONTAB
            objRelatorio.Rel_Menu_Executar ("Lotes")

        Case MENU_CTB_REL_CAD_LOTESPEND
            objRelatorio.Rel_Menu_Executar ("Lotes Pendentes")

        Case MENU_CTB_REL_CAD_LANDATA
            objRelatorio.Rel_Menu_Executar ("Lançamentos por Data")

        Case MENU_CTB_REL_CAD_LANCCL
            objRelatorio.Rel_Menu_Executar ("Lançamentos por Centro de Custo")

        Case MENU_CTB_REL_CAD_LANLOTE
            objRelatorio.Rel_Menu_Executar ("Lançamentos por Lote")

        Case MENU_CTB_REL_CAD_LANPEND
            objRelatorio.Rel_Menu_Executar ("Lançamentos Pendentes")

    End Select

End Sub

Private Sub mnuCTBRelDemMutPatrLiq_Click()

Dim objRelatorio As New AdmRelatorio

    objRelatorio.Rel_Menu_Executar ("Demonstrativo de Mutação do Patrimônio Líquido")
    
End Sub

Private Sub mnuCTBRelDemOAR_Click()

Dim objRelatorio As New AdmRelatorio

    objRelatorio.Rel_Menu_Executar ("Demonstração das Origens e Aplicações de Recursos")

End Sub

Private Sub mnuCTBRelDRE_Click()

Dim objRelatorio As New AdmRelatorio

    objRelatorio.Rel_Menu_Executar ("Demonstrativo de Resultado do Exercício")

End Sub

Private Sub mnuCTBRelDRP_Click()

Dim objRelatorio As New AdmRelatorio

    objRelatorio.Rel_Menu_Executar ("Demonstrativo de Resultado do Período")

End Sub

Private Sub mnuCTBRot_Click(Index As Integer)

    Select Case Index

        Case MENU_CTB_ROT_ATUALIZALOTE
            Call Chama_Tela("LoteAtualiza")

        Case MENU_CTB_ROT_APURAPERIODO
            Call Chama_Tela("ApuraPeriodo")

        Case MENU_CTB_ROT_APURAEXERC
            Call Chama_Tela("ApuraExercicio")

        Case MENU_CTB_ROT_FECHAEXERC
            Call Chama_Tela("FechamentoExercicio")

        Case MENU_CTB_ROT_REPROCESSAMENTO
            Call Chama_Tela("Reprocessamento")

        Case MENU_CTB_ROT_REABREEXERC
            Call Chama_Tela("ReaberturaExercicio")

        Case MENU_CTB_ROT_RATEIOOFF
            Call Chama_Tela("RateioOffBatch")

        Case MENU_CTB_ROT_IMPORTACAO
            Call Chama_Tela("ImportacaoCtb")

        Case MENU_CTB_ROT_GERACAODRE
            Call Chama_Tela("GeracaoDRE")

        Case MENU_CTB_ROT_DESAPURAEXERC
            Call Chama_Tela("DesapuraExercicio")

        Case MENU_CTB_ROT_IMPORTACAORATEIO
            Call Chama_Tela("ImportacaoRateio")

        Case MENU_CTB_ROT_TRVRATEIO
            Call Chama_Tela("TRVRateio")
            
        Case MENU_CTB_ROT_IMPORTLCTOS
            Call Chama_Tela("ImportLctos")

    End Select

End Sub

Private Sub mnuESTCad_Click(Index As Integer)

    Select Case Index
    
        Case MENU_EST_CAD_FORN
            Call Chama_Tela("Fornecedores")
            
        Case MENU_EST_CAD_ALMOXARIFADOS
            Call Chama_Tela("Almoxarifado")
                    
        Case MENU_EST_CAD_PRODUTOS
            Call Chama_Tela("Produto")
        
        Case MENU_EST_CAD_KIT
            Call Chama_Tela("Kit")
            
        Case MENU_EST_CAD_CCL
            Call Chama_Tela("CclTela")

        Case MENU_EST_CAD_TRANSPORTADORAS
            Call Chama_Tela("Transportadora")

        Case MENU_EST_CAD_PRODALM
            Call Chama_Tela("EstoqueProduto")
                    
        Case MENU_EST_CAD_EMBALAGEM
            Call Chama_Tela("Embalagem")
        
        Case MENU_EST_CAD_PRODUTO_DESCONTO
            Call Chama_Tela("ProdutoDesconto")
            
                    
    End Select

End Sub

Private Sub mnuESTCadLoteContabil_Click()

Dim objLote As ClassLote

    Call Chama_Tela("LoteTela", objLote, MODULO_ESTOQUE)

End Sub

Private Sub mnuESTCadLoteInventario_Click()
        Call Chama_Tela("LoteEst")
End Sub

Private Sub mnuESTCadLoteRastro_Click()
    Call Chama_Tela("RastreamentoLote")
End Sub

Private Sub mnuESTCadTA_Click(Index As Integer)

    Select Case Index

        Case MENU_EST_CAD_TA_CONDPAG
            Call Chama_Tela("CondicoesPagto")

        Case MENU_EST_CAD_TA_TIPOSFORN
            Call Chama_Tela("TipoFornecedor")

        Case MENU_EST_CAD_TA_CUSTOS
            Call Chama_Tela("Custos")

        Case MENU_EST_CAD_TA_CUSTOPROD
            Call Chama_Tela("CustoProducao")

        Case MENU_EST_CAD_TA_ESTOQUE
            Call Chama_Tela("Estoque")

        Case MENU_EST_CAD_TA_TIPOPROD
            Call Chama_Tela("TipoProduto")

        Case MENU_EST_CAD_TA_UNIDADEMED
            Call Chama_Tela("ClasseUM")

        Case MENU_EST_CAD_TA_ESTOQUEINI
            Call Chama_Tela("EstoqueInicial")

        Case MENU_EST_CAD_TA_CATEGORIAPROD
            Call Chama_Tela("CategoriaProduto")
        
        Case MENU_EST_CAD_TA_ESTADOS
            Call Chama_Tela("Estados")

        Case MENU_EST_CAD_TA_SERIESNFISC
            Call Chama_Tela("SerieNFiscal")

        Case MENU_EST_CAD_TA_NATUREZAOP
            Call Chama_Tela("NaturezaOperacao")

        Case MENU_EST_CAD_TA_EXCECOESICMS
            Call Chama_Tela("ExcecoesICMS")

        Case MENU_EST_CAD_TA_EXCECOESIPI
            Call Chama_Tela("ExcecoesIPI")

        Case MENU_EST_CAD_TA_TIPOTRIB
            Call Chama_Tela("TipoDeTributacao")

        Case MENU_EST_CAD_TA_TRIBUTACAOFORN
            Call Chama_Tela("PadraoTribEntrada")

        Case MENU_EST_CAD_TA_TRIBUTACAOCLI
            Call Chama_Tela("PadraoTribSaida")

        Case MENU_EST_CAD_TA_CATEGFORNECEDOR
            Call Chama_Tela("CategoriaFornec")

        Case MENU_EST_CAD_TA_PRODUTOEMBALAGEM
            Call Chama_Tela("ProdutoEmbalagem")

        '##################################
        'Inserido por Wagner
        Case MENU_EST_CAD_TA_CONTRATOPAG
            Call Chama_Tela("ContratoPagar")
        '##################################

        Case MENU_EST_CAD_TA_INVTERC
            Call Chama_Tela("InventarioTerc")
    
        Case MENU_EST_CAD_TA_CORVAR
            Call Chama_Tela("CorVariacao")
    
        Case MENU_EST_CAD_TA_PINTURA
            Call Chama_Tela("Pintura")
    
        Case MENU_EST_CAD_TA_COLECAO
            Call Chama_Tela("Colecao")
    
        Case MENU_EST_CAD_TA_DECL_IMPORTACAO
            Call Chama_Tela("DIInfo")
        
        Case MENU_EST_CAD_TA_PROD_GRADE
            Call Chama_Tela("ProdutoGrade")
        
    End Select

End Sub

Private Sub mnuESTConCad_Click(Index As Integer)
        
Dim colSelecao As New Collection

    Select Case Index
        
        Case MENU_EST_CON_CAD_FORNECEDORES
            Call Chama_Tela("FornecedorLista")

        Case MENU_EST_CON_CAD_PRODUTOS
            Call Chama_Tela("ProdutoLista_Consulta", colSelecao)

        Case MENU_EST_CON_CAD_ALMOXARIFADO
            Call Chama_Tela("AlmoxarifadoLista_Consulta", colSelecao)
        
        Case MENU_EST_CON_CAD_KIT
            Call Chama_Tela("KitLista")

        Case MENU_EST_CON_CAD_TIPOPROD
            Call Chama_Tela("TipoProdutoLista")

        Case MENU_EST_CON_CAD_CLASSEUM
            Call Chama_Tela("ClasseUMLista", colSelecao)

        Case MENU_EST_CON_CAD_CATPROD
            Call Chama_Tela("CategoriaProdutoLista", colSelecao)

        Case MENU_EST_CON_CAD_CATPRODITEM
            Call Chama_Tela("CategoriaProdutoItemLista", colSelecao)

        Case MENU_EST_CON_CAD_PRODCATPRODITEM
            Call Chama_Tela("ProdutoFilialCategoriaLista", colSelecao)

    End Select

End Sub

Private Sub mnuESTConEst_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
                
        Case MENU_EST_CON_EST_ESTPROD
            Call Chama_Tela("EstProdLista_Consulta")

        Case MENU_EST_CON_EST_ESTPRODFILIAL
            Call Chama_Tela("EstProdFilialLista_Cons", colSelecao)

        Case MENU_EST_CON_EST_ESTPRODTERC
            Call Chama_Tela("EstProdTercLista_Consulta", colSelecao)

        Case MENU_EST_CON_EST_SALDODISP
            Call Chama_Tela("EstProdQtdeDispLista", colSelecao)

        Case MENU_EST_CON_EST_ESTPROD_MV
            Call Chama_Tela("EstProd_AlmoxPharLista", colSelecao)

        Case MENU_EST_CON_PROD_EM_FALTA
            Call Chama_Tela("EstProdFilMenorEstSegLista", colSelecao)

    End Select

End Sub

Private Sub mnuESTConfig_Click(Index As Integer)

Dim objForm As Form
Dim lErro As Long

On Error GoTo Erro_mnuESTConfig_Click

    Select Case Index

        Case MENU_EST_CONFIG_SEGMENTOS
            
            For Each objForm In Forms
                If Not (objForm Is Me) And Not (objForm Is gobjEstInicial) Then Error 59361
            Next

            Call Chama_Tela_Modal("SegmentosMAT")
        
        Case MENU_EST_CONFIG_CONFIGURACAO
            Call Chama_Tela("ConfiguraEST")

    End Select

    Exit Sub
    
Erro_mnuESTConfig_Click:

    Select Case Err

        Case 59361
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FECHAR_JANELAS_FILHAS", Err, Error)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165245)

    End Select

    Exit Sub

End Sub

Private Sub mnuESTConInv_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
        
        Case MENU_EST_CON_INV_INVENTARIO
            Call Chama_Tela("InventarioLista", colSelecao)

        Case MENU_EST_CON_INV_INVENTARIOLOTE
            Call Chama_Tela("InventarioLoteLista", colSelecao)

        Case MENU_EST_CON_INV_INVENTARIOLOTEPEND
            Call Chama_Tela("InvLotePendenteLista", colSelecao)

    End Select

End Sub

Private Sub mnuESTConMov_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
            
        Case MENU_EST_CON_MOV_RECMAT
            Call Chama_Tela("RecebMaterialLista", colSelecao)

        Case MENU_EST_CON_MOV_RECMATFORN
            Call Chama_Tela("RecebMaterialFLista", colSelecao)

        Case MENU_EST_CON_MOV_RECMATCLI
            Call Chama_Tela("RecebMaterialCLista", colSelecao)

        Case MENU_EST_CON_MOV_RESERVAS
            Call Chama_Tela("ReservaLista", colSelecao)
        
        Case MENU_EST_CON_MOV_MOVEST
            'Call Chama_Tela("MovEstoqueLista_Consulta", colSelecao)
            Call Chama_Tela("MovimentosEstoque2Lista", colSelecao)

        Case MENU_EST_CON_MOV_MOVESTINT
            Call Chama_Tela("MovEstoqueInternoLista", colSelecao)

        Case MENU_EST_CON_MOV_MOVESTTRANSF
            Call Chama_Tela("MovEstoqueTransferenciaLista", colSelecao)
        
        Case MENU_EST_CON_MOV_CONSUMO
            colSelecao.Add MOV_EST_CONSUMO
            Call Chama_Tela("MovEstoqueLista", colSelecao)

    End Select

End Sub

Private Sub mnuESTConNF_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index

        Case MENU_EST_CON_NF_NFISCENTTODAS
            Call Chama_Tela("NFiscalEntradaTodasLista", colSelecao)

        Case MENU_EST_CON_NF_NFISCENT
            Call Chama_Tela("NFiscalEntradaLista", colSelecao)

        Case MENU_EST_CON_NF_NFISCENTFAT
            Call Chama_Tela("NFiscalFatEntradaLista", colSelecao)

        Case MENU_EST_CON_NF_NFISCENTREM
            Call Chama_Tela("NFiscalEntRemLista", colSelecao)

        Case MENU_EST_CON_NF_NFISCENTDEV
            Call Chama_Tela("NFiscalEntDevLista", colSelecao)

        Case MENU_EST_CON_NF_ITENSNFISCENT
            Call Chama_Tela("ItensNFiscalTodasEnt_Lista", colSelecao)

    End Select

End Sub

Private Sub mnuESTConNFSai_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
                
        Case MENU_EST_CON_NFSAI_NFISCTODAS
            Call Chama_Tela("NFiscalSaidaTodasLista", colSelecao)
                
        Case MENU_EST_CON_NFSAI_NFISC
            Call Chama_Tela("NFiscalLista", colSelecao)

        Case MENU_EST_CON_NFSAI_NFISCFAT
            Call Chama_Tela("NFiscalFaturaLista", colSelecao)
        
        Case MENU_EST_CON_NFSAI_NFISCPED
            Call Chama_Tela("NFiscalPedidoLista", colSelecao)
        
        Case MENU_EST_CON_NFSAI_NFISCFATPED
            Call Chama_Tela("NFiscalFaturaPedidoLista", colSelecao)

        Case MENU_EST_CON_NFSAI_NFISCREM
            Call Chama_Tela("NFiscalRemLista", colSelecao)

        Case MENU_EST_CON_NFSAI_NFISCDEV
            Call Chama_Tela("NFiscalDevLista", colSelecao)
        
        Case MENU_EST_CON_NFSAI_ITENSNFISC
            Call Chama_Tela("ItensNFiscalTodosSaida_Lista", colSelecao)
        
    End Select


End Sub

Private Sub mnuESTConPro_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index

        Case MENU_EST_CON_PRO_OP
            Call Chama_Tela("OrdemProducaoLista", colSelecao)
            
        Case MENU_EST_CON_PRO_OP_BAIXA
            Call Chama_Tela("OrdemProdBaixadasLista", colSelecao)
            
        Case MENU_EST_CON_PRO_EMPENHO
            Call Chama_Tela("EmpenhoLista", colSelecao)

        Case MENU_EST_CON_PRO_ITENSOP
            Call Chama_Tela("ItemOrdemProducaoLista", colSelecao)

        Case MENU_EST_CON_PRO_REQPROD
            colSelecao.Add MOV_EST_REQ_PRODUCAO
            colSelecao.Add MOV_EST_REQ_PRODUCAO_BENEF3
            colSelecao.Add MOV_EST_REQ_PRODUCAO_OUTROS
            
            Call Chama_Tela("MovEstoqueLista1", colSelecao)
            
        Case MENU_EST_CON_PRO_PRODUCAO
            colSelecao.Add MOV_EST_PRODUCAO
            colSelecao.Add MOV_EST_PRODUCAO_BENEF3
            colSelecao.Add MOV_EST_PRODUCAO_OUTROS
            
            Call Chama_Tela("MovEstoqueLista1", colSelecao)
            
        Case MENU_EST_CON_PRO_REQPRODOP
            colSelecao.Add MOV_EST_REQ_PRODUCAO
            colSelecao.Add MOV_EST_REQ_PRODUCAO_BENEF3
            
            Call Chama_Tela("MovEstoqueOPLista", colSelecao)
       
        Case MENU_EST_CON_PRO_PRODUCAOOP
            colSelecao.Add MOV_EST_PRODUCAO
            colSelecao.Add MOV_EST_PRODUCAO_BENEF3

            Call Chama_Tela("MovEstoqueOPLista", colSelecao)
       
    End Select

End Sub
 
Private Sub mnuESTConRastro_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
                
        Case MENU_EST_CON_RASTRO_LOTES
            Call Chama_Tela("RastroLoteLista", colSelecao)
                
        Case MENU_EST_CON_RASTRO_SALDOS
            Call Chama_Tela("RastroLoteSaldoLista", colSelecao)

        Case MENU_EST_CON_RASTRO_MOVIMENTOS
            Call Chama_Tela("RastroMovEstoqueLista", colSelecao)
        
        Case MENU_EST_CON_RASTRO_SALDOS_PRECO
            Call Chama_Tela("RastroLotePrecoLista", colSelecao)
        
        Case MENU_EST_CON_RASTRO_MOVIMENTOS_PRECO
            Call Chama_Tela("RastroMovEstPrecoLista", colSelecao)
                
        Case MENU_EST_CON_RASTRO_ITENS_NF
            Call Chama_Tela("ItemNFRastroLista", colSelecao)
        
        Case MENU_EST_CON_RASTRO_PHAR_LM_SEP
            Call Chama_Tela("PharLMSepLista", colSelecao)
        
        Case MENU_EST_CON_RASTRO_SALDOS_CM
            Call Chama_Tela("RastroSaldoLotePharLista", colSelecao)
        
        
    End Select

End Sub

Private Sub mnuESTMov_Click(Index As Integer)

    Select Case Index
        
        Case MENU_EST_MOV_RECMATFORN
            Call Chama_Tela("RecebMaterialF")

        Case MENU_EST_MOV_RECMATCLI
            Call Chama_Tela("RecebMaterialC")

        Case MENU_EST_MOV_RECMATFORNCOM
            Call Chama_Tela("RecebMaterialFCom")

        Case MENU_EST_MOV_PRODUCAOSAIDA
            Call Chama_Tela("ProducaoSaida")

        Case MENU_EST_MOV_REQCONSUMO
            Call Chama_Tela("ReqConsumo")

        Case MENU_EST_MOV_MOVIMENTOINTERNO
            Call Chama_Tela("MovEstoque")
       
        Case MENU_EST_MOV_TRANSFERENCIAS
            Call Chama_Tela("Transfer")

        Case MENU_EST_MOV_PRODUCAOENTRADA
            Call Chama_Tela("ProducaoEntrada")

        Case MENU_EST_MOV_RESERVA
            Call Chama_Tela("Reserva")
        
        Case MENU_EST_MOV_INVENTARIO
            Call Chama_Tela("Inventario")

        Case MENU_EST_MOV_INVENTARIOLOTE
            Call Chama_Tela("InventarioLote")

        Case MENU_EST_MOV_INVENTARIOTERC
            Call Chama_Tela("InventarioCliForn")

        Case MENU_EST_MOV_NFISCENT
            Call Chama_Tela("NFiscalEntrada")

        Case MENU_EST_MOV_NFISCENTCOM
            Call Chama_Tela("NFiscalEntradaCom")

        Case MENU_EST_MOV_NFISCENTFAT
            Call Chama_Tela("NFiscalFatEntrada")

        Case MENU_EST_MOV_NFISCENTFATCOM
            Call Chama_Tela("NFiscalFatEntradaCom")

        Case MENU_EST_MOV_NFISCENTREM
            Call Chama_Tela("NFiscalEntRem")
        
        Case MENU_EST_MOV_NFISCENTDEV
            Call Chama_Tela("NFiscalEntDev")

        Case MENU_EST_MOV_CANCELAR_NF
            Call Chama_Tela("CancelaNFiscalEst")
        
        Case MENU_EST_MOV_MEDICAO
            Call Chama_Tela("ContratoMedicaoPag")
        
        Case MENU_EST_MOV_DESMEMBRAMENTO
            Call Chama_Tela("Desmembramento")
       
    End Select

End Sub

Private Sub mnuESTRel_Click(Index As Integer)

Dim lErro As Long
Dim iCancela As Integer
Dim sCodRel As String
Dim objRelSel As New AdmRelSel
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_mnuESTRel_Click

    Select Case Index

        Case MENU_EST_REL_PRODVEND
            objRelatorio.Rel_Menu_Executar ("Relação dos Produtos Vendidos")

        Case MENU_EST_REL_RESENTSAIVALOR
            objRelatorio.Rel_Menu_Executar ("Resumo das Entradas e Saídas em Valor")

        Case MENU_EST_REL_CONSUMOVENDAMES
            objRelatorio.Rel_Menu_Executar ("Consumos / Vendas Mês a Mês")

        Case MENU_EST_REL_ANALEST
            objRelatorio.Rel_Menu_Executar ("Análise do Estoque")

        Case MENU_EST_REL_ANALMOVEST
            objRelatorio.Rel_Menu_Executar ("Análise de Movimentações de Estoque")

        Case MENU_EST_REL_PONTOPEDIDO
            objRelatorio.Rel_Menu_Executar ("Produtos que atingiram o Ponto de Pedido")

        Case MENU_EST_REL_SALDOEST
            objRelatorio.Rel_Menu_Executar ("Saldo em Estoque")

        Case MENU_EST_REL_OUTROS
            lErro = Chama_Tela("RelSelecionar", objRelSel, MODULO_ESTOQUE)
            If (lErro <> 0) Then Error 7058

            If (objRelSel.iCancela <> 0) Then Exit Sub

            'Prosseguir executando relatório
            objRelatorio.Rel_Menu_Executar (objRelSel.sCodRel)

        Case MENU_EST_REL_GERREL
            Sistema_EditarRel ("")

        '####################################
        'Inserido por Wagner
        Case MENU_EST_REL_PLANILHAS
            lErro = Chama_Tela("PlanilhasSelecionar", MODULO_ESTOQUE)
            If (lErro <> 0) Then Error 7058
        '####################################
        
    End Select

    Exit Sub

Erro_mnuESTRel_Click:

    Select Case Err

        Case 7058

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165246)

    End Select

    Exit Sub

End Sub

Private Sub mnuESTRelCad_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio
    
    Select Case Index

        Case MENU_EST_REL_CAD_FORN
            objRelatorio.Rel_Menu_Executar ("Relação de Fornecedores")

        Case MENU_EST_REL_CAD_PRODUTOS
            objRelatorio.Rel_Menu_Executar ("Relação de Produtos")
        
        Case MENU_EST_REL_CAD_ALMOXARIFADO
            objRelatorio.Rel_Menu_Executar ("Relação de Almoxarifados")

        Case MENU_EST_REL_CAD_KITS
            objRelatorio.Rel_Menu_Executar ("Relação de Kits")
        
        Case MENU_EST_REL_CAD_UTILPROD
            objRelatorio.Rel_Menu_Executar ("Utilização do Produto")

    End Select

End Sub

Private Sub mnuESTRelInv_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_EST_REL_INV_ETIQINV
            objRelatorio.Rel_Menu_Executar ("Etiquetas Para Inventário")
        
        Case MENU_EST_REL_INV_INVENTARIO
            objRelatorio.Rel_Menu_Executar ("Listagem Para Inventário")
              
        Case MENU_EST_REL_INV_DEMAPINV
            objRelatorio.Rel_Menu_Executar ("Demonstrativo de Apuração de Inventário")
  
        Case MENU_EST_REL_INV_REGINVMOD7
            objRelatorio.Rel_Menu_Executar ("Registro de Inventário")

    End Select

End Sub

Private Sub mnuESTRelMov_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index
        
        Case MENU_EST_REL_MOV_MOVINT
            objRelatorio.Rel_Menu_Executar ("Relação das Movimentações Internas")

        Case MENU_EST_REL_MOV_REQCONSUMO
            objRelatorio.Rel_Menu_Executar ("Requisições Para Consumo")
        
        Case MENU_EST_REL_MOV_KARDEX
            objRelatorio.Rel_Menu_Executar ("Lista dos Movimentos")

        Case MENU_EST_REL_MOV_KARDEXDIA
            objRelatorio.Rel_Menu_Executar ("Lista dos Movimentos Diários")

        Case MENU_EST_REL_MOV_RESUMOKARDEX
            objRelatorio.Rel_Menu_Executar ("Lista dos Movimentos Resumidos Por Dia")

        Case MENU_EST_REL_MOV_BOLETIMENT
            objRelatorio.Rel_Menu_Executar ("Boletim de Entrada por Notas Fiscais")

    End Select

End Sub

Private Sub mnuESTRelPro_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_EST_REL_PRO_OP
            objRelatorio.Rel_Menu_Executar ("Relação das Ordens de Produção")

        Case MENU_EST_REL_PRO_EMPENHOS
            objRelatorio.Rel_Menu_Executar ("Lista dos Empenhos")

        Case MENU_EST_REL_PRO_LISTAFALTAS
            objRelatorio.Rel_Menu_Executar ("Lista de Faltas")
        
        Case MENU_EST_REL_PRO_MOVESTOP
            objRelatorio.Rel_Menu_Executar ("Movimentos de Estoque para cada Ordem de Produção")
        
        Case MENU_EST_REL_PRO_RESPRODOP
            objRelatorio.Rel_Menu_Executar ("Resumo dos Produtos - Ordem de Produção")
        
    End Select

End Sub

Private Sub mnuESTRot_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_EST_ROT_ATUALIZALOTE
            Call Chama_Tela("LoteEstAtualiza")
        
        Case MENU_EST_ROT_ATUALIZARASTROLOTE
            Call Chama_Tela("RastroItensNFEST")
        
        Case MENU_EST_ROT_CUSTOMEDIOPRODUCAO
            Call Chama_Tela("CustoProducao")

        Case MENU_EST_ROT_FECHAMES
            Call Chama_Tela("FechamentoMesEst")

        Case MENU_EST_ROT_CLASSIFICACAOABC
            Call Chama_Tela("ClassificacaoABC")

        Case MENU_EST_ROT_EMINFISC
            objRelatorio.Rel_Menu_Executar ("Emissão das Notas Fiscais")
        
        Case MENU_EST_ROT_EMINFISCFAT
            objRelatorio.Rel_Menu_Executar ("Emissão das Notas Fiscais Fatura")
                
        Case MENU_EST_ROT_EMISSREC
            objRelatorio.Rel_Menu_Executar ("Emissão das Notas de Recebimento")

        Case MENU_EST_ROT_REPROCESSAMENTO
            Call Chama_Tela("ReprocessamentoEST")

        Case MENU_EST_ROT_IMPORTACAOINV
            Call Chama_Tela("ImportacaoInv")
            
        Case MENU_EST_ROT_IMPORTARNFRAIZ
            Chama_Tela ("ImportarNFRaiz")

        Case MENU_EST_ROT_IMPORT_XML
            objRelatorio.Rel_Menu_Executar ("Importação de xml de NFe")

    End Select

End Sub

Private Sub mnuFATCad_Click(Index As Integer)

Dim objLote As ClassLote

    Select Case Index

        Case MENU_FAT_CAD_CLIENTES
            Call Chama_Tela("Clientes")

        Case MENU_FAT_CAD_PRODUTOS
            Call Chama_Tela("Produto")

        Case MENU_FAT_CAD_VENDEDORES
            Call Chama_Tela("Vendedores")

        Case MENU_FAT_CAD_TRANSPORTADORAS
            Call Chama_Tela("Transportadora")

        Case MENU_FAT_CAD_LOTES
            Call Chama_Tela("LoteTela", objLote, MODULO_FATURAMENTO)

        Case MENU_FAT_CAD_GERACAOSENHA
            Load GeracaoSenha
            If lErro_Chama_Tela = SUCESSO Then
                GeracaoSenha.Show vbModeless, Me
            End If

        Case MENU_FAT_CAD_FAMILIAS
            Call Chama_Tela("Familias")
            
        Case MENU_FAT_CAD_EMI
            Call Chama_Tela("TRPEmissores")
        
        Case MENU_FAT_CAD_TIPOOCR
            Call Chama_Tela("TRPTiposOcorrencia")
        
        Case MENU_FAT_CAD_ACORDO
            Call Chama_Tela("TRPAcordos")
            
    End Select

End Sub

Private Sub mnuFATCadTA_Click(Index As Integer)

    Select Case Index

        Case MENU_FAT_CAD_TA_CONTRATOPROG
            Call Chama_Tela("ContratoPropaganda")
            
'        Case MENU_FAT_CAD_TA_KITVENDA
'            Call Chama_Tela("KitVenda")

        'Inserido Pelo Wagner
        '###############
        Case MENU_FAT_CAD_TA_CONTRATOCAD
            Call Chama_Tela("ContratoCadastro")
        '#############

        Case MENU_FAT_CAD_TA_TIPOCLI
            Call Chama_Tela("TipoCliente")

        Case MENU_FAT_CAD_TA_TIPOPROD
            Call Chama_Tela("TipoProduto")

        Case MENU_FAT_CAD_TA_TIPOVEND
             Call Chama_Tela("TipoVendedor")

        Case MENU_FAT_CAD_TA_CODICAOPAG
            Call Chama_Tela("CondicoesPagto")

'        Case MENU_FAT_CAD_TA_REGIOESVENDA
'            Call Chama_Tela("RegiaoVenda")
'
'        Case MENU_FAT_CAD_TA_CANAISVENDA
'            Call Chama_Tela("CanalDeVenda")

        Case MENU_FAT_CAD_TA_TIPOBLOQUEIOS
            Call Chama_Tela("TipoDeBloqueio")

'        Case MENU_FAT_CAD_TA_TABPRECOS
'            If giFilialEmpresa = EMPRESA_TODA Or giFilialEmpresa = Abs(giFilialAuxiliar) Then
'                Call Chama_Tela("TabelaPrecoItemEmpresaToda")
'            Else
'                Call Chama_Tela("TabelaPrecoItem")
'            End If

        Case MENU_FAT_CAD_TA_FORMPRECO
            Call Chama_Tela("FormacaoPreco")
        
        Case MENU_FAT_CAD_TA_MNEMONICOFPRECO
            Call Chama_Tela("MnemonicoFPreco")

        Case MENU_FAT_CAD_TA_SERIESNFISC
            Call Chama_Tela("SerieNFiscal")

'        Case MENU_FAT_CAD_TA_PREVISAOVENDA
'            Call Chama_Tela("PrevVendaMensal")

        Case MENU_FAT_CAD_TA_ESTADOS
            Call Chama_Tela("Estados")

        Case MENU_FAT_CAD_TA_CATEGORIACLI
            Call Chama_Tela("CategoriaCliente")
            
        Case MENU_FAT_CAD_TA_CATEGORIAPROD
            Call Chama_Tela("CategoriaProduto")

'        Case MENU_FAT_CAD_TA_NATUREZAOP
'            Call Chama_Tela("NaturezaOperacao")
'
'        Case MENU_FAT_CAD_TA_CLASFISC
'            Call Chama_Tela("ClassificacaoFiscal")
'
'        Case MENU_FAT_CAD_TA_EXCECOESICMS
'            Call Chama_Tela("ExcecoesICMS")
'
'        Case MENU_FAT_CAD_TA_EXCECOESIPI
'            Call Chama_Tela("ExcecoesIPI")
'
'        Case MENU_FAT_CAD_TA_TIPOTRIB
'            Call Chama_Tela("TipoDeTributacao")
'
'        Case MENU_FAT_CAD_TA_TRIBUTACAOFORN
'            Call Chama_Tela("PadraoTribEntrada")
'
'        Case MENU_FAT_CAD_TA_TRIBUTACAOCLI
'            Call Chama_Tela("PadraoTribSaida")

        Case MENU_FAT_CAD_TA_UNIDADEMED
            Call Chama_Tela("ClasseUM")
            
        Case MENU_FAT_CAD_TA_MENSAGENS
            Call Chama_Tela("Mensagens")

        Case MENU_FAT_CAD_TA_CAMPOSGENERICOS
            Call Chama_Tela("CamposGenericos")

        Case MENU_FAT_CAD_TA_CLIENTECONTATOS
            Call Chama_Tela("ClienteContatos")

        Case MENU_FAT_CAD_TA_ATENDENTES
            Call Chama_Tela("Atendentes")

        'Inserido por Jorge Specian
        '---------------------------
        Case MENU_FAT_CAD_TA_PROJETO
            Call Chama_Tela("Projetos")
        
'        Case MENU_FAT_CAD_TA_CUSTEIOPROJETO
'            Call Chama_Tela("CusteioProjeto")
        '---------------------------

'        Case MENU_FAT_CAD_TA_DASALIQUOTAS
'            Call Chama_Tela("DASAliquotas")
        
        Case MENU_FAT_CAD_TA_CLIENTEEXPRESSO
            Call Chama_Tela("ClienteExpresso")
            
        Case MENU_FAT_CAD_TA_MODELOSEMAIL
            Call Chama_Tela("ModelosEmail")
            
        Case MENU_FAT_CAD_TA_PLANILHACOMISSOES
            Call Chama_Tela("PlanComissoesInpal")
            
        Case MENU_FAT_CAD_TA_DEINFO
            Call Chama_Tela("DEInfo")
            
    End Select

End Sub

Private Sub mnuFATCadFIS_Click(Index As Integer)

    Select Case Index
        
        Case MENU_FAT_CAD_FIS_NATUREZAOP
            Call Chama_Tela("NaturezaOperacao")

        Case MENU_FAT_CAD_FIS_CLASFISC
            Call Chama_Tela("ClassificacaoFiscal")

        Case MENU_FAT_CAD_FIS_EXCECOESICMS
            Call Chama_Tela("ExcecoesICMS")

        Case MENU_FAT_CAD_FIS_EXCECOESIPI
            Call Chama_Tela("ExcecoesIPI")

        Case MENU_FAT_CAD_FIS_TIPOTRIB
            Call Chama_Tela("TipoDeTributacao")

        Case MENU_FAT_CAD_FIS_TRIBUTACAOFORN
            Call Chama_Tela("PadraoTribEntrada")

        Case MENU_FAT_CAD_FIS_TRIBUTACAOCLI
            Call Chama_Tela("PadraoTribSaida")

        Case MENU_FAT_CAD_FIS_DASALIQUOTAS
            Call Chama_Tela("DASAliquotas")
            
        Case MENU_FAT_CAD_FIS_EXCECOESPISCOFINS
            Call Chama_Tela("ExcecoesPISCOFINS")
            
    End Select

End Sub

Private Sub mnuFATCadVEND_Click(Index As Integer)

    Select Case Index

        
        Case MENU_FAT_CAD_VEND_REGIOESVENDA
            Call Chama_Tela("RegiaoVenda")

        Case MENU_FAT_CAD_VEND_CANAISVENDA
            Call Chama_Tela("CanalDeVenda")

        Case MENU_FAT_CAD_VEND_TABPRECOS
            If giFilialEmpresa = EMPRESA_TODA Or giFilialEmpresa = Abs(giFilialAuxiliar) Then
                Call Chama_Tela("TabelaPrecoItemEmpresaToda")
            Else
                Call Chama_Tela("TabelaPrecoItem")
            End If

        Case MENU_FAT_CAD_VEND_PREVISAOVENDA
            Call Chama_Tela("PrevVendaMensal")

        Case MENU_FAT_CAD_VEND_KITVENDA
            Call Chama_Tela("KitVenda")
            
        Case MENU_FAT_CAD_VEND_CODICAOPAG
            Call Chama_Tela("CondicoesPagto")
            
        Case MENU_FAT_CAD_VEND_ROTAS
            Call Chama_Tela("Rotas")
            
        Case MENU_FAT_CAD_VEND_VEICULOS
            Call Chama_Tela("Veiculos")
            
        Case MENU_FAT_CAD_VEND_PVANDAMENTO
            Call Chama_Tela("PVAndamento")
            
        Case MENU_FAT_CAD_VEND_TABPRECOGRUPO
            Call Chama_Tela("TabelaPrecoGrupo")
            
    End Select

End Sub

Private Sub mnuFATConCad_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
        
        Case MENU_FAT_CON_CAD_CLIENTES
            Call Chama_Tela("ClientesLista")
        
        Case MENU_FAT_CON_CAD_CLIENTESCONSULTA
            Call Chama_Tela("ClienteConsulta")

        Case MENU_FAT_CON_CAD_VENDEDORES
            Call Chama_Tela("VendedorLista")
        
        Case MENU_FAT_CON_CAD_TRANSPORTADORAS
            Call Chama_Tela("TransportadoraLista")

        Case MENU_FAT_CON_CAD_CATPROD
            Call Chama_Tela("CategoriaProdutoLista", colSelecao)

        Case MENU_FAT_CON_CAD_CATPRODITEM
            Call Chama_Tela("CategoriaProdutoItemLista", colSelecao)
                   
        Case MENU_FAT_CON_CAD_EMISSORES
            Call Chama_Tela("TRPEmissoresLista")

        Case MENU_FAT_CON_CAD_ACORDOS
            Call Chama_Tela("TRPAcordosLista")

        Case MENU_FAT_CON_CAD_ITENS_CONTRATO_FAT
            
            colSelecao.Add 0
            colSelecao.Add 0
            colSelecao.Add 0
            colSelecao.Add 0
            colSelecao.Add ""
            colSelecao.Add ""
            colSelecao.Add ""
            colSelecao.Add ""
            
            Call Chama_Tela("ContratosCliItensLista", colSelecao)

    End Select
    
End Sub

Private Sub mnuFATConEST_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
        
        Case MENU_FAT_CON_EST_PRODUTOS
            Call Chama_Tela("ProdutoLista_Consulta", colSelecao)
        
        Case MENU_FAT_CON_EST_ESTPRODFILIAL
            Call Chama_Tela("EstProdFilialLista_Cons", colSelecao)
        
        Case MENU_FAT_CON_EST_ESTPROD
            Call Chama_Tela("EstProdLista_Consulta")
        
        Case MENU_FAT_CON_EST_ESTPRODTERC
            Call Chama_Tela("EstProdTercLista_Consulta", colSelecao)
    
    End Select
    
End Sub

Private Sub mnuFATConfig_Click(Index As Integer)

Dim objForm As Form
Dim lErro As Long

On Error GoTo Erro_mnuFATConfig_Click

    Select Case Index
    
        Case MENU_FAT_CONFIG_CONFIGURACAO
            Call Chama_Tela("ConfiguraFAT")
    
        Case MENU_FAT_CONFIG_AUTORIZACREDITO
            Call Chama_Tela("AlcadaFat")
    
        Case MENU_FAT_CONFIG_SEGMENTOS
            For Each objForm In Forms
                If Not (objForm Is Me) And Not (objForm Is gobjEstInicial) Then Error 59362
            Next
            
            Call Chama_Tela_Modal("SegmentosMAT")
            
        Case MENU_FAT_CONFIG_COMISSOESREGRAS
            Call Chama_Tela("ComissoesRegras")
            
        Case MENU_FAT_CONFIG_OUTRASCONFIG
            Call Chama_Tela("TRPConfig")

        Case MENU_FAT_CONFIG_EMAILCONFIG
            Call Chama_Tela("EmailConfig")
            
        Case MENU_FAT_CONFIG_REGRASMSG
            Call Chama_Tela("RegrasMsg")
            
    End Select
    
    Exit Sub
    
Erro_mnuFATConfig_Click:

    Select Case Err

        Case 59362
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FECHAR_JANELAS_FILHAS", Err, Error)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165247)

    End Select

    Exit Sub

End Sub

Private Sub mnuFATConGraf_Click(Index As Integer)

    Select Case Index

        Case MENU_FAT_CON_GRAF_AREA
            Call Chama_Tela("GrafFaturamentoCategProd")

        Case MENU_FAT_CON_GRAF_CLIENTE
            Call Chama_Tela("GrafFaturamentoCli")

        Case MENU_FAT_CON_GRAF_MENSAL_DOLAR
             Call Chama_Tela("GrafFaturamentoMensalDolar")

        Case MENU_FAT_CON_GRAF_MENSAL
            Call Chama_Tela("GrafFaturamentoMensal")
    
    End Select

End Sub

Private Sub mnuFATConNF_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
                
        Case MENU_FAT_CON_NF_NFISCTODAS
            Call Chama_Tela("NFiscalSaidaTodasLista", colSelecao)
                
        Case MENU_FAT_CON_NF_NFISC
            Call Chama_Tela("NFiscalLista", colSelecao)

        Case MENU_FAT_CON_NF_NFISCFAT
            Call Chama_Tela("NFiscalFaturaLista", colSelecao)
        
        Case MENU_FAT_CON_NF_NFISCPED
            Call Chama_Tela("NFiscalPedidoLista", colSelecao)
        
        Case MENU_FAT_CON_NF_NFISCFATPED
            Call Chama_Tela("NFiscalFaturaPedidoLista", colSelecao)
        'Janaina
        Case MENU_FAT_CON_NF_NFISCREMPED
            Call Chama_Tela("NFiscalRemPedidoLista", colSelecao)
        'Janaina
        Case MENU_FAT_CON_NF_NFISCREM
            Call Chama_Tela("NFiscalRemLista", colSelecao)

        Case MENU_FAT_CON_NF_NFISCDEV
            Call Chama_Tela("NFiscalDevLista", colSelecao)
        
        Case MENU_FAT_CON_NF_SERIESNF
            Call Chama_Tela("SerieLista", colSelecao)
    
        Case MENU_FAT_CON_NF_ITENSNFISC
            Call Chama_Tela("ItensNFiscalTodosSaida_Lista", colSelecao)
    
        Case MENU_FAT_CON_NF_CONHECTRANSP
            Call Chama_Tela("NFConhecFreteTodosLista", colSelecao)
            
        Case MENU_FAT_CON_NF_DIRECT
            Call Chama_Tela("TranspConsultaPhar")
                    
        Case MENU_FAT_CON_NF_AFAT
            Call Chama_Tela("NFiscalAFaturarLista", colSelecao)
                    
        Case MENU_FAT_CON_NF_BI
            Call Chama_Tela("NF_Guaraplus_BILista", colSelecao)
                    
    End Select

End Sub

Private Sub mnuFATConNFEnt_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index

        Case MENU_FAT_CON_NFENT_NFISCENTTODAS
            Call Chama_Tela("NFiscalEntradaTodasLista", colSelecao)

        Case MENU_FAT_CON_NFENT_NFISCENT
            Call Chama_Tela("NFiscalEntradaLista", colSelecao)

        Case MENU_FAT_CON_NFENT_NFISCENTFAT
            Call Chama_Tela("NFiscalFatEntradaLista", colSelecao)

        Case MENU_FAT_CON_NFENT_NFISCENTREM
            Call Chama_Tela("NFiscalEntRemLista", colSelecao)

        Case MENU_FAT_CON_NFENT_NFISCENTDEV
            Call Chama_Tela("NFiscalEntDevLista", colSelecao)

        Case MENU_FAT_CON_NFENT_ITENSNFISCENT
            Call Chama_Tela("ItensNFiscalTodasEnt_Lista", colSelecao)

    End Select

End Sub

Private Sub mnuFATConVen_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
        
        Case MENU_FAT_CON_VEN_PEDVEND_ATIVOS
            Call Chama_Tela("PedidoVendaLista", colSelecao)
        
        Case MENU_FAT_CON_VEN_PEDVEND_BAIXADOS
            Call Chama_Tela("PedidosDeVendaBaixadosLista", colSelecao)
        
        Case MENU_FAT_CON_VEN_TABPRECO
            Call Chama_Tela("TabelaPrecoLista", colSelecao)
        
        Case MENU_FAT_CON_VEN_TABPRECOIT
            Call Chama_Tela("TabelaPrecoItemLista", colSelecao)
        
        Case MENU_FAT_CON_VEN_PREVVEND
            Call Chama_Tela("PrevVendaLista", colSelecao)
                
        Case MENU_FAT_CON_VEN_ORCAMENTO
            Call Chama_Tela("OrcamentoVendaCGLista", colSelecao)
                
        Case MENU_FAT_CON_VEN_TABPRECOITAT
            Call Chama_Tela("TabPrecoItensAtualLista", colSelecao)
                
        Case MENU_FAT_CON_VEN_ROTA
            Call Chama_Tela("RotasLista", colSelecao)
                
        Case MENU_FAT_CON_VEN_MAPA
            Call Chama_Tela("MapaDeEntregaLista", colSelecao)
                
        Case MENU_FAT_CON_VEN_PVACOMP
            Call Chama_Tela("PVConsulta2")
                
        Case MENU_FAT_CON_VEN_PVITENS
            Call Chama_Tela("ItensPVLista")
                
        Case MENU_FAT_CON_VEN_OVITENS
            Call Chama_Tela("ItensOVLista")
                
        Case 100
            Call Chama_Tela("ArtPedidoVenda_Lista", colSelecao)
        
        Case 101
            Call Chama_Tela("ArtPedidoVendaItens_Lista", colSelecao)
    
        Case 102
            Call Chama_Tela("ArtPedQuitacaoRes_Lista", colSelecao)
        
        Case 103
            Call Chama_Tela("ArtPedQuitacaoDet_Lista", colSelecao)
            
        Case MENU_FAT_CON_VEN_PROTHEUS_HIST_VENDAS
            Call Chama_Tela("PolandHistVendaLista", colSelecao)

            
    End Select

End Sub

Private Sub mnuFATMov_Click(Index As Integer)

    Select Case Index

        'Inserido pelo Wagner
        '##########
        Case MENU_FAT_MOV_CONTRATOMED
            Call Chama_Tela("ContratoMedicao")
        '############
        
'        Case MENU_FAT_MOV_ORCVENDA
'            Call Chama_Tela("OrcamentoVenda")
'
'        Case MENU_FAT_MOV_PEDVEND
'            Call Chama_Tela("PedidoVenda")
'
'        Case MENU_FAT_MOV_LIBBLOQUEIO
'            Call Chama_Tela("LiberaBloqueio")

        Case MENU_FAT_MOV_GERANFISC
            Call Chama_Tela("GeracaoNFiscal")

'        Case MENU_FAT_MOV_BAIXAMANPED
'            Call Chama_Tela("BaixaPedido")

        Case MENU_FAT_MOV_NFISC
            Call Chama_Tela("NFiscal")

        Case MENU_FAT_MOV_NFISCFATURA
            Call Chama_Tela("NFiscalFatura")

        Case MENU_FAT_MOV_CANCELAR_NF
            Call Chama_Tela("CancelaNFiscal")
        
        Case MENU_FAT_MOV_GERAFAT
            Call Chama_Tela("GeracaoFatura")

        Case MENU_FAT_MOV_NFISCPED
            Call Chama_Tela("NFiscalPedido")

        Case MENU_FAT_MOV_NFISCFATPED
            Call Chama_Tela("NFiscalFaturaPedido")

        Case MENU_FAT_MOV_COMISSOES
            Call Chama_Tela("Comissoes")
        
        Case MENU_FAT_MOV_NFISCDEV
            Call Chama_Tela("NFiscalDev")
        'Janaina
        Case MENU_FAT_MOV_NFISCREMPED
            Call Chama_Tela("NFiscalRemPedido")
        'Janaina
        Case MENU_FAT_MOV_NFISCREM
            Call Chama_Tela("NFiscalRem")
        
'        Case MENU_FAT_MOV_CONHECIMENTOFRETE
'            Call Chama_Tela("ConhecimentoFrete")
'
'        Case MENU_FAT_MOV_CONHECIMENTOFRETEFAT
'            Call Chama_Tela("ConhecimentoFreteFatura")

        Case MENU_FAT_MOV_GERNFFAT
            Call Chama_Tela("TRPGeracaoNF")
            
        Case MENU_FAT_MOV_VEND_MAPA
            Call Chama_Tela("MapaDeEntrega")

    End Select

End Sub

Private Sub mnuFATRel_Click(Index As Integer)

Dim lErro As Long
Dim iCancela As Integer
Dim sCodRel As String
Dim objRelSel As New AdmRelSel
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_mnuFATRel_Click

    Select Case Index

        Case MENU_FAT_REL_COMISSOES
            objRelatorio.Rel_Menu_Executar ("Relatório de Comissões")
        
        Case MENU_FAT_REL_LISTAPRECOS
            objRelatorio.Rel_Menu_Executar ("Lista de Preços")

        Case MENU_FAT_REL_OUTROS
            lErro = Chama_Tela("RelSelecionar", objRelSel, MODULO_FATURAMENTO)
            If (lErro <> 0) Then Error 7058

            If (objRelSel.iCancela <> 0) Then Exit Sub

            'Prosseguir executando relatório
            objRelatorio.Rel_Menu_Executar (objRelSel.sCodRel)

        Case MENU_FAT_REL_GERREL
            Sistema_EditarRel ("")
            
        '####################################
        'Inserido por Wagner
        Case MENU_FAT_REL_PLANILHAS
            lErro = Chama_Tela("PlanilhasSelecionar", MODULO_FATURAMENTO)
            If (lErro <> 0) Then Error 7058
        '####################################
        

    End Select

    Exit Sub

Erro_mnuFATRel_Click:

    Select Case Err

        Case 7058

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165248)

    End Select

    Exit Sub

End Sub

Private Sub mnuFATRelCad_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_FAT_REL_CAD_CLI
            objRelatorio.Rel_Menu_Executar ("Relação de Clientes")
        
        Case MENU_FAT_REL_CAD_PRODUTOS
            objRelatorio.Rel_Menu_Executar ("Relação de Produtos")
                        
        Case MENU_FAT_REL_CAD_VENDEDORES
            objRelatorio.Rel_Menu_Executar ("Relação de Vendedores")

    End Select
    
End Sub

Private Sub mnuFATRelDoc_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_FAT_REL_DOC_PRENOTA
            objRelatorio.Rel_Menu_Executar ("Pré-Nota")

        Case MENU_FAT_REL_DOC_NFISC
            objRelatorio.Rel_Menu_Executar ("Relação das Notas Fiscais")
        
        Case MENU_FAT_REL_DOC_NFISCDEV
            objRelatorio.Rel_Menu_Executar ("Relação das Notas Fiscais de Devolução")
        
        Case MENU_FAT_REL_DOC_NFISCTRANSP
            objRelatorio.Rel_Menu_Executar ("Relação das Notas Fiscais para as Transportadoras")

    End Select
    
End Sub

Private Sub mnuFATRelPed_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_FAT_REL_PED_PEDIDSOAPTOSFAT
            objRelatorio.Rel_Menu_Executar ("Relação dos Pedidos Aptos a Faturar")

        Case MENU_FAT_REL_PED_PEDIDOSNENTREGUE
            objRelatorio.Rel_Menu_Executar ("Relação dos Pedidos não Faturados")

        Case MENU_FAT_REL_PED_PEDIDOPROD
            objRelatorio.Rel_Menu_Executar ("Relação de Pedidos por Produto")

        Case MENU_FAT_REL_PED_PEDIDOSVENDCLI
            objRelatorio.Rel_Menu_Executar ("Pedidos de Vendas por Vendedor/Cliente")

        Case MENU_FAT_REL_PED_PEDIDOSVENDPROD
            objRelatorio.Rel_Menu_Executar ("Relação de Pedidos por Vendedor x Produto")
        
        Case MENU_FAT_REL_PED_PEDIDOSPRODUCAO
            objRelatorio.Rel_Menu_Executar ("Pedidos Para Produção")

        Case MENU_FAT_REL_PED_PEDIDOSCLI
            objRelatorio.Rel_Menu_Executar ("Pedidos de Vendas por Cliente")

    End Select
    
End Sub

Private Sub mnuFATRelVen_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_FAT_REL_VEN_FATCLI
            objRelatorio.Rel_Menu_Executar ("Faturamento por Cliente")
        
        Case MENU_FAT_REL_VEN_FATCLIPROD
            objRelatorio.Rel_Menu_Executar ("Faturamento Cliente x Produto")

        Case MENU_FAT_REL_VEN_FATVEND
            objRelatorio.Rel_Menu_Executar ("Faturamento por Vendedor")

        Case MENU_FAT_REL_VEN_FATPRAZOPAG
            objRelatorio.Rel_Menu_Executar ("Faturamento por Prazo de Pagamento")

        Case MENU_FAT_REL_VEN_FATREALPREV
            objRelatorio.Rel_Menu_Executar ("Faturamento Real x Previsto")

        Case MENU_FAT_REL_VEN_RESVEND
            objRelatorio.Rel_Menu_Executar ("Resumo de Vendas")
               
        Case MENU_FAT_REL_VEN_DISPESTVENDA
            objRelatorio.Rel_Menu_Executar ("Disponibilidade de Estoque para Vendas")

    End Select
    
End Sub

Private Sub mnuFATRotImp_Click(Index As Integer)

Dim objRelatorio As AdmRelatorio

    Select Case Index

        Case MENU_FAT_ROT_IMPORTAR_DADOS
            Chama_Tela ("ImportarDados")

        Case MENU_FAT_ROT_IMPORTAR_DADOS_MANUAL
            Chama_Tela ("ImportarDadosArq")
            
        Case MENU_FAT_ROT_IMPORTAR_DADOS_XML
            Set objRelatorio = New AdmRelatorio
            objRelatorio.Rel_Menu_Executar ("Importação de xml de NFe")
            
    End Select
    
End Sub
    
Private Sub mnuFATRot_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_FAT_ROT_EXPORTAR_NF
            Chama_Tela ("ExportarNFiscal")
            
        Case MENU_FAT_ROT_IMPORTAR_NF
            Chama_Tela ("ImportarNFiscal")
            
        Case MENU_FAT_ROT_EXPORTAR_DADOS
            Chama_Tela ("ExportarDados")
            
'        Case MENU_FAT_ROT_IMPORTAR_DADOS
'            Chama_Tela ("ImportarDados")
            
        Case MENU_FAT_ROT_GERARQ_RPS_LOTE
            Chama_Tela ("GerArqRPSLote")
            
        Case MENU_FAT_ROT_IMPORTAR_NFE
            Chama_Tela ("ImportarNFe")
            
        'Inserido por Wagner
        '###################
        Case MENU_FAT_ROT_CONTRATOFATLOTE
            Chama_Tela ("ContratoFatLote")
        '##################
        
        Case MENU_FAT_ROT_REAJUSTEPRECO
            Chama_Tela ("AtualizacaoPreco")
        
        Case MENU_FAT_ROT_ATUALIZARASTRO
            Chama_Tela ("RastroItensNFFAT")
            
        Case MENU_FAT_ROT_ROMANEIODESPACHO
            objRelatorio.Rel_Menu_Executar ("Romaneio de Separação")
        
        Case MENU_FAT_ROT_EMINFISC
            objRelatorio.Rel_Menu_Executar ("Emissão das Notas Fiscais")
        
        Case MENU_FAT_ROT_EMINFISCFAT
            objRelatorio.Rel_Menu_Executar ("Emissão das Notas Fiscais Fatura")
                
        Case MENU_FAT_ROT_EMIFATURAS
            objRelatorio.Rel_Menu_Executar ("Emissão de Faturas")

        Case MENU_FAT_ROT_EMIDUPLICATAS
            objRelatorio.Rel_Menu_Executar ("Emissão de Duplicatas")
            
        'Inserido por Wagner
        '###################
        Case MENU_FAT_ROT_GERARQPV
            Chama_Tela ("GeracaoArqPVLote")
        '###################

        Case MENU_FAT_ROT_EXPORTAR_NF_HARMONIA
            Chama_Tela ("NFiscalExportaHar")
            
        Case MENU_FAT_ROT_NF_PAULISTA
            Chama_Tela ("NFiscalPaulista")
            
        Case MENU_FAT_ROT_IMPORTAR_PV_SET
            Chama_Tela ("ImportaPV")
            
        Case MENU_FAT_ROT_REGFATHTML
            Call Chama_Tela("TRPRegerarFaturas")
            
        Case MENU_FAT_ROT_TRP_VEND_REMONTA
            Call Chama_Tela("TRPVendRemonta")
            
        Case MENU_FAT_ROT_ARQCOMISSOES
            Chama_Tela ("ArqComissoes")
            
        Case MENU_FAT_ROT_EXPORTLOREAL
            objRelatorio.Rel_Menu_Executar ("Exportação Loreal")
            
        Case MENU_FAT_ROT_EXPORTGENVEND
            objRelatorio.Rel_Menu_Executar ("Exportação Gerenciador de Vendas")
            
    End Select

End Sub

Private Sub mnuFISCon_Click(Index As Integer)

    Select Case Index

        Case MENU_FIS_CON_EXCECOES_ICMS
            Call Chama_Tela("ExcecoesICMSLista")
        
        Case MENU_FIS_CON_EXCECOES_IPI
            Call Chama_Tela("ExcecoesIPILista")
        
        Case MENU_FIS_CON_NATUREZA_OPERACAO
            Call Chama_Tela("NaturezaOperacaoLista")

        Case MENU_FIS_CON_PRODUTOS
            Call Chama_Tela("ProdutoLista_Consulta")
        
        Case MENU_FIS_CON_TIPO_TRIBUTACAO
            Call Chama_Tela("TiposTributacaoTodosLista")
        
        Case MENU_FIS_CON_TIPO_REG_APUR_ICMS
            Call Chama_Tela("TiposRegApuracaoICMSLista")

        Case MENU_FIS_CON_TIPO_REG_APUR_IPI
            Call Chama_Tela("TiposRegApuracaoIPILista")
        
        Case MENU_FIS_CON_LIVRO_ABERTOS
            Call Chama_Tela("LivrosAbertosTodosLista")
        
        Case MENU_FIS_CON_LIVROS_FECHADOS
            Call Chama_Tela("LivrosFechadosTodosLista")
        
        Case MENU_FIS_CON_APUR_ICMS
            Call Chama_Tela("ApuracaoICMSLista")
        
        Case MENU_FIS_CON_APUR_IPI
            Call Chama_Tela("ApuracaoIPILista")
        
        Case MENU_FIS_CON_LANC_APUR_ICMS
            Call Chama_Tela("ApuracaoICMSItensLista")

        Case MENU_FIS_CON_LANC_APUR_IPI
            Call Chama_Tela("ApuracaoIPIItensLista")
        
        Case MENU_FIS_CON_REG_ENTRADA
            Call Chama_Tela("EdicaoRegEntrada_Lista")
        
        Case MENU_FIS_CON_REG_SAIDA
            Call Chama_Tela("EdicaoRegSaida_Lista")

        Case MENU_FIS_CON_ITEM_NF
            Call Chama_Tela("ItensNF_ICMSLista")

        Case MENU_FIS_CON_ITEM_NF_TRIB
            Call Chama_Tela("ItensNFTribLista")
    
    End Select

End Sub

Private Sub mnuFISConfig_Click(Index As Integer)
    
    Select Case Index

        Case MENU_FIS_CONFIG_CONFIGURACAO
            Call Chama_Tela("ConfiguraFIS")
        
        Case MENU_FIS_CONFIG_TRIBUTO
            Call Chama_Tela("TributoConfigura")
        
        Case MENU_FIS_CONFIG_ICMS
            Call Chama_Tela("Estados")

    End Select
    
End Sub

Private Sub mnuFISRel_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio
Dim objRelSel As New AdmRelSel
Dim lErro As Long

On Error GoTo Erro_mnuFISRel_Click

    Select Case Index
    
        Case MENU_FIS_REL_APURACAO_COFINS
            objRelatorio.Rel_Menu_Executar ("Apuração do COFINS")
        
        Case MENU_FIS_REL_APURACAO_PIS
            objRelatorio.Rel_Menu_Executar ("Apuração do PIS")
        
        Case MENU_FIS_REL_RESUMO_DECLAN
            objRelatorio.Rel_Menu_Executar ("Resumo para o preenchimento do Declan")
            
        Case MENU_FIS_REL_TERMO
            objRelatorio.Rel_Menu_Executar ("Termo de Abertura dos Livros")

        Case MENU_FIS_REL_IR
            objRelatorio.Rel_Menu_Executar ("Relatório para Recolhimento de IRRF")
        
        Case MENU_FIS_REL_INSS
            objRelatorio.Rel_Menu_Executar ("Relatório para Recolhimento de INSS Retido")
    
        Case MENU_FIS_REL_PISCOFINSCSLL
            objRelatorio.Rel_Menu_Executar ("Relatório de Retenção CSLL PIS COFINS")
        
        Case MENU_FIS_REL_OUTROS
            lErro = Chama_Tela("RelSelecionar", objRelSel, MODULO_LIVROSFISCAIS)
            If (lErro <> 0) Then Error 7058

            If (objRelSel.iCancela <> 0) Then Exit Sub

            'Prosseguir executando relatório
            objRelatorio.Rel_Menu_Executar (objRelSel.sCodRel)

        Case MENU_FIS_REL_GERREL
            Sistema_EditarRel ("")
        
        '####################################
        'Inserido por Wagner
        Case MENU_FIS_REL_PLANILHAS
            lErro = Chama_Tela("PlanilhasSelecionar", MODULO_LIVROSFISCAIS)
            If (lErro <> 0) Then Error 7058
        '####################################
        
    End Select

    Exit Sub
    
Erro_mnuFISRel_Click:

    Select Case Err

        Case 7058

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165249)

    End Select

    Exit Sub
    
End Sub

Private Sub mnuFISRelGerenciais_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio
    
    Select Case Index
    
        Case MENU_FIS_REL_LANC_NATOPERACAO
            objRelatorio.Rel_Menu_Executar ("Lista de Reg. de Entrada/Saída p/ Nat. de Operação")
            
        Case MENU_FIS_REL_LANC_ESTADO
            objRelatorio.Rel_Menu_Executar ("Lista de Reg. de Entrada/Saída por Estado")
                    
        Case MENU_FIS_REL_LANC_TIPOICMS
            objRelatorio.Rel_Menu_Executar ("Lista de Reg. de Entrada/Saída por Tipo ICMS")
            
        Case MENU_FIS_REL_LANC_CLIENTE
            objRelatorio.Rel_Menu_Executar ("Lista de Reg. de Entrada/Saída por Cliente")
                    
        Case MENU_FIS_REL_LANC_FORNECEDOR
            objRelatorio.Rel_Menu_Executar ("Lista de Reg. de Entrada/Saída por Fornecedor")
                    
        Case MENU_FIS_REL_LANC_TODOS
            objRelatorio.Rel_Menu_Executar ("Lista de Reg. de Entrada/Saída")
                    
    End Select
    
End Sub

Private Sub mnuFISRelLivICMSIPI_Click(Index As Integer)
    
Dim objRelatorio As New AdmRelatorio
    
    Select Case Index
        
        Case MENU_FIS_REL_LIVREG_ENTRADA
            objRelatorio.Rel_Menu_Executar ("Livro de Reg. de Entradas")
        
        Case MENU_FIS_REL_SAIDA
            objRelatorio.Rel_Menu_Executar ("Livro de Reg. de Saídas")
        
        Case MENU_FIS_REL_RCPE
            objRelatorio.Rel_Menu_Executar ("Registro de Controle da Produção e do Estoque")
        
        Case MENU_FIS_REL_REGINVENTARIO
            objRelatorio.Rel_Menu_Executar ("Livro de Reg. de Inventário")
           
        Case MENU_FIS_REL_APURACAO_ICMS
            objRelatorio.Rel_Menu_Executar ("Apuração do ICMS")
    
        Case MENU_FIS_REL_RESUMOICMS
            objRelatorio.Rel_Menu_Executar ("Resumo da Apuração ICMS")
    
        Case MENU_FIS_REL_APURACAO_IPI
            objRelatorio.Rel_Menu_Executar ("Apuração do IPI")

        Case MENU_FIS_REL_RESUMOIPI
            objRelatorio.Rel_Menu_Executar ("Resumo da Apuração IPI")

        Case MENU_FIS_REL_EMITENTES
            objRelatorio.Rel_Menu_Executar ("Lista de Códigos de Emitentes")

        Case MENU_FIS_REL_MERCARDORIAS
            objRelatorio.Rel_Menu_Executar ("Tabela de Códigos de Mercadorias")
    
        Case MENU_FIS_REL_OPERINTEREST
            objRelatorio.Rel_Menu_Executar ("Listagem de Operações Interestaduais")
            
        Case MENU_FIS_REL_GNRICMS
            objRelatorio.Rel_Menu_Executar ("Guia de Recolhimento Nacional de ICMS")
    
    End Select
    
End Sub

Private Sub mnuFISRelLivroISS_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio
    
    Select Case Index
    
        Case MENU_FIS_REL_APURACAO_ISS
            objRelatorio.Rel_Menu_Executar ("Apuração do ISS")
    
    End Select
    
End Sub

Private Sub mnuJanelasCascade_Click(Index As Integer)
    gbTelaReordenando = True
    Me.Arrange (vbCascade)
    gbTelaReordenando = False
End Sub

Private Sub mnuLJCad_Click(Index As Integer)
    
        Select Case Index
    
            Case MENU_LJ_CAD_CLIENTE_LOJA
                If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                    Call Chama_Tela("ClienteLoja")
                Else
                    Call Chama_Tela("Clientes")
                End If
            
            Case MENU_LJ_CAD_PROD
                If giLocalOperacao <> LOCALOPERACAO_CAIXA_CENTRAL Then
                    Call Chama_Tela("Produto")
                End If
                
            Case MENU_LJ_CAD_OPERADOR
                Call Chama_Tela("Operador")
    
            Case MENU_LJ_CAD_CAIXA
                Call Chama_Tela("Caixa")
    
            Case MENU_LJ_CAD_VENDEDOR
                If giLocalOperacao <> LOCALOPERACAO_CAIXA_CENTRAL Then
                    Call Chama_Tela("Vendedores")
                End If
            Case MENU_LJ_CAD_ECF
                Call Chama_Tela("ECF")
            
        End Select
    
End Sub

Private Sub mnuLJCadTA_Click(Index As Integer)
    
        Select Case Index
    
            Case MENU_LJ_CAD_TA_TABPRECO
            If giFilialEmpresa = EMPRESA_TODA Or giFilialEmpresa = Abs(giFilialAuxiliar) Then
                Call Chama_Tela("TabelaPrecoItemEmpresaToda")
            Else
                Call Chama_Tela("TabelaPrecoItem")
            End If

            Case MENU_LJ_CAD_TA_PRODDESC
                Call Chama_Tela("ProdutoDesconto")
                            
            Case MENU_LJ_CAD_TA_ADM_MEIO_PAG
                Call Chama_Tela("AdmMeioPagto")
    
            Case MENU_LJ_CAD_TA_REDE
                Call Chama_Tela("Rede")
            
            Case MENU_LJ_CAD_TA_TECLADO
                Call Chama_Tela("Teclado")
                    
            Case MENU_LJ_CAD_TA_IMPRESSORA
                Call Chama_Tela("ImpressoraECF")
    
        End Select
    
End Sub

Private Sub mnuLJCon_Click(Index As Integer)

Dim objMovCaixa As New ClassMovimentoCaixa
Dim colSelecao As New Collection

    Select Case Index
        
        Case MENU_LJ_CON_MOVCAIXA
             Call Chama_Tela("MovCaixaTipoMovLista")
            
        Case MENU_LJ_CON_MOVCAIXA_SESSOES
            Call Chama_Tela("MovCaixaSessoesLista", colSelecao, objMovCaixa, Nothing, "Tipo IN (35,68,32)")
    
        Case MENU_LJ_CON_MOVCAIXACF
             Call Chama_Tela("MovCxCFLista")
             
        Case MENU_LJ_CON_CUPOMFISCAL
            If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then
                Call Chama_Tela("CupomFiscalLista")
            Else
                Call Chama_Tela("CupomFiscalLista")
            End If
            
        Case MENU_LJ_CON_ITEMCUPOMFISCAL
            If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then
                Call Chama_Tela("ItensCupomFiscalLista")
            Else
                Call Chama_Tela("ItensCupomFiscalLista")
            End If
    
    End Select
        
End Sub

Private Sub mnuLJConfig_Click(Index As Integer)
    
    Select Case Index

        Case MENU_LJ_CONFIG_CONFIGURACAO
            Call Chama_Tela("LojaConfig")

        Case MENU_LJ_CONFIG_TECLADO
            Call Chama_Tela("TecladoProduto")

        Case MENU_LJ_CONFIG_VENDEDOR_LOJA
            Call Chama_Tela("VendedorFilial")

    End Select
    
End Sub

Private Sub mnuLJMov_Click(Index As Integer)

    Select Case Index
        
        Case MENU_LJ_MOV_BORD_BOLETO
            Call Chama_Tela("BorderoBoleto")
        
        Case MENU_LJ_MOV_CHEQUENESP
            Call Chama_Tela("ChequeNEsp")
        
        Case MENU_LJ_MOV_BORD_CHEQUE
                Call Chama_Tela("BorderoCheque")
            
        Case MENU_LJ_MOV_BORD_VALE_TICKET
            Call Chama_Tela("BorderoValeTicket")
        
        Case MENU_LJ_MOV_DEPOSITO_BANCARIO
            Call Chama_Tela("DepositoBancario")

        Case MENU_LJ_MOV_DEPOSITO_CAIXA
            Call Chama_Tela("DepositoCaixa")

        Case MENU_LJ_MOV_SAQUE_CAIXA
            Call Chama_Tela("SaqueCaixa")

        Case MENU_LJ_MOV_TRANSFERENCIA
            Call Chama_Tela("TransfCentral")
            
        Case MENU_LJ_MOV_BORDEROOUTROS
            Call Chama_Tela("BorderoOutros")
        
        Case MENU_LJ_MOV_RECEB_CARNE
            Call Chama_Tela("RecebimentoCarne")

    End Select
        
End Sub

Private Sub mnuLJRot_Click(Index As Integer)
    
    Select Case Index

        Case MENU_LJ_ROT_GERACAOARQCC
            Call Chama_Tela("GeracaoArqCC")
            
        Case MENU_LJ_ROT_GERACAOARQBACK
            Call Chama_Tela("GeracaoArqBack")
            
        Case MENU_LJ_ROT_CARGABALANCA
            Call Chama_Tela("CargaBalanca")
    
    End Select
        
    Exit Sub
    
End Sub

Private Sub mnuPCPConCur_Click(Index As Integer)

    Select Case Index

        Case MENU_PCP_CON_CUR_CERTIFICADO
            Call Chama_Tela("CertificadosLista")
            
        Case MENU_PCP_CON_CUR_CURSO
            Call Chama_Tela("CursosLista")
            
        Case MENU_PCP_CON_CUR_CURSOMO
            Call Chama_Tela("CursoMOLista")
            
        Case MENU_PCP_CON_CUR_CERTIFICADOMO
            Call Chama_Tela("CertificadoMOLista")
            
    End Select
    
End Sub

Private Sub mnuQUACad_Click(Index As Integer)

    Select Case Index

        Case MENU_QUA_CAD_PRODUTOS
            Call Chama_Tela("Produto")
            
        Case MENU_QUA_CAD_TESTES
            Call Chama_Tela("TestesQualidade")
            
        Case MENU_QUA_CAD_PRODUTOTESTES
            Call Chama_Tela("ProdutoTeste")
            
    End Select

End Sub

Private Sub mnuQUACon_Click(Index As Integer)

    Select Case Index
        
        Case MENU_QUA_CON_RESULTADO_TESTES
            Call Chama_Tela("RastreamentoLoteTesteLista")
        
    End Select

End Sub

Private Sub mnuQUAConCad_Click(Index As Integer)

    Select Case Index

        Case MENU_QUA_CON_CAD_PRODUTOS
            Call Chama_Tela("ProdutoLista_Consulta")
            
        Case MENU_QUA_CON_CAD_TESTES
            Call Chama_Tela("TestesQualidadeLista")
            
        Case MENU_QUA_CON_CAD_PRODUTOTESTES
            Call Chama_Tela("ProdutoTesteLista")
        
    End Select

End Sub

Private Sub mnuQUAMov_Click(Index As Integer)

    Select Case Index

        Case MENU_QUA_MOV_RESULTADOS
            Call Chama_Tela("RastreamentoLote")
            
    End Select

End Sub

Private Sub mnuQUARel_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio
Dim objRelSel As New AdmRelSel
Dim lErro As Long

On Error GoTo Erro_mnuQUARel_Click

    Select Case Index
    
        Case MENU_QUA_REL_FICHA_CONTROLE
            objRelatorio.Rel_Menu_Executar ("Ficha de Controle de Qualidade")
            
        Case MENU_QUA_REL_LAUDOS_NF
            objRelatorio.Rel_Menu_Executar ("Laudos do Controle de Qualidade de NFs")
            
        Case MENU_QUA_REL_NAO_CONFORME
            Call objRelatorio.Rel_Menu_Executar("Não Conformidade")
            
        Case MENU_QUA_REL_OUTROS
            lErro = Chama_Tela("RelSelecionar", objRelSel, MODULO_QUALIDADE)
            If (lErro <> 0) Then Error 7058

            If (objRelSel.iCancela <> 0) Then Exit Sub

            'Prosseguir executando relatório
            objRelatorio.Rel_Menu_Executar (objRelSel.sCodRel)

        Case MENU_QUA_REL_GERREL
            Sistema_EditarRel ("")
    
        '####################################
        'Inserido por Wagner
        Case MENU_QUA_REL_PLANILHAS
            lErro = Chama_Tela("PlanilhasSelecionar", MODULO_QUALIDADE)
            If (lErro <> 0) Then Error 7058
        '####################################
            
    End Select

    Exit Sub
    
Erro_mnuQUARel_Click:

    Select Case Err

        Case 7058

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165250)

    End Select

    Exit Sub
    
End Sub

Private Sub mnuQUARelCad_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_QUA_REL_CAD_PRODUTO
            objRelatorio.Rel_Menu_Executar ("Relação de Produtos")

        Case MENU_QUA_REL_CAD_PRODUTOTESTE
            objRelatorio.Rel_Menu_Executar ("Relação de Testes por Produto")

        Case MENU_QUA_REL_CAD_TESTE
            objRelatorio.Rel_Menu_Executar ("Relação de Testes")

    End Select

End Sub

Private Sub mnuSRVCadTA_Click(Index As Integer)

    Select Case Index

        Case 5
            Call Chama_Tela("TipoDeBloqueioGen")

    End Select
    
End Sub

Private Sub mnuSRVConfig_Click(Index As Integer)

    Select Case Index

        Case MENU_SRV_CONFIG_CONFIGURACAO
            Call Chama_Tela("ConfiguraSRV")

    End Select

End Sub

Private Sub mnuSRVRel_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_SRV_REL_SOLIC
            objRelatorio.Rel_Menu_Executar ("Solicitação de Serviço")

        Case MENU_SRV_REL_OS
            objRelatorio.Rel_Menu_Executar ("Ordem de Serviço")

    End Select
    
End Sub

Private Sub mnuSRVRelPS_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index

        Case MENU_SRV_REL_PS_APTOSFAT
            objRelatorio.Rel_Menu_Executar ("Pedidos de Serviço aptos a faturar")

        Case MENU_SRV_REL_PS_NAOENTR
            objRelatorio.Rel_Menu_Executar ("Pedidos de Serviço não entregues")

        Case MENU_SRV_REL_PS_PORCLI
            objRelatorio.Rel_Menu_Executar ("Pedidos por cliente")

        Case MENU_SRV_REL_PS_PORSRV
            objRelatorio.Rel_Menu_Executar ("Pedidos por serviço")

    End Select
    
End Sub

Private Sub mnuTESCad_Click(Index As Integer)

Dim objLote As ClassLote

    Select Case Index

        Case MENU_TES_CAD_BANCOS
            Call Chama_Tela("Bancos")

        Case MENU_TES_CAD_CONTACORRENTE
            Call Chama_Tela("CtaCorrenteInt")

        Case MENU_TES_CAD_FAVORECIDOS
            Call Chama_Tela("Favorecidos")

        Case MENU_TES_CAD_LOTES
            Call Chama_Tela("LoteTela", objLote, MODULO_TESOURARIA)

    End Select

End Sub

Private Sub mnuTESCadTA_Click(Index As Integer)

    Select Case Index

        Case MENU_TES_CAD_TA_TIPOAPLIC
            Call Chama_Tela("TipoAplicacao")

        Case MENU_TES_CAD_TA_HISTEXTRATO
            Call Chama_Tela("HistMovCta")

        Case MENU_TES_CAD_TA_NATMOVCTA
            Call Chama_Tela("TiposMovtoCtaCorrente1")

    End Select

End Sub

Private Sub mnuTESCon_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index

        Case MENU_TES_CON_FLUXOCAIXA
            Call Chama_Tela("FluxoDeCaixa")

        Case MENU_TES_CON_APLICACOES
            Call Chama_Tela("AplicacaoLista", colSelecao)

        Case MENU_TES_CON_BANCOS
            Call Chama_Tela("BancoLista")

        Case MENU_TES_CON_CONTASCORRENTES
            Call Chama_Tela("CtaCorrenteLista", colSelecao)

        Case MENU_TES_CON_CONTASCORRENTESFILIAIS
            Call Chama_Tela("CtaCorrenteTodasLista")

        Case MENU_TES_CON_DEPOSITOS
            Call Chama_Tela("DepositoLista", colSelecao)

        Case MENU_TES_CON_SAQUES
            Call Chama_Tela("SaqueLista", colSelecao)

''        Case MENU_TES_CON_TRANSPORTADORAS
''            Call Chama_Tela("TransportadoraLista")

        Case MENU_TES_CON_TIPOAPLICACAO
            Call Chama_Tela("TipoAplicacaoLista")

        Case MENU_TES_CON_TRANSFERENCIA
            Call Chama_Tela("TransferenciaLista")
            
        Case MENU_TES_CON_MOVCC
        
            colSelecao.Add giFilialEmpresa
        
            Call Chama_Tela("MovCtaCorrenteLista", colSelecao, Nothing, Nothing, "FilialEmpresa = ?")
        
        Case MENU_TES_CON_MOVCC_TF
            Call Chama_Tela("MovCtaCorrenteLista", colSelecao)

        Case MENU_TES_CON_FLUXOCAIXACTB
            Call Chama_Tela("FlCxCtb")
            
        Case MENU_TES_CON_FLUXOCAIXACTB1
            Call Chama_Tela("FlCxCtb1")
            
    End Select

End Sub

Private Sub mnuTESConfig_Click(Index As Integer)
        
    Select Case Index
        
        Case MENU_TES_CONFIG_CONFIGURACOES
            Call Chama_Tela("ConfiguraTES")

    End Select

End Sub

Private Sub mnuTESMov_Click(Index As Integer)

    Select Case Index

        Case MENU_TES_MOV_SAQUE
            Call Chama_Tela("Saque")

        Case MENU_TES_MOV_DEPOSITO
            Call Chama_Tela("Deposito")

        Case MENU_TES_MOV_TRANFERENCIA
            Call Chama_Tela("Transferencia")

        Case MENU_TES_MOV_APLICACAO
            Call Chama_Tela("Aplicacao")

        Case MENU_TES_MOV_RESGATE
            Call Chama_Tela("Resgate")

        Case MENU_TES_MOV_CONCILIACAO
            Call Chama_Tela("ConciliacaoBancaria")

    End Select

End Sub

Private Sub mnuTESRel_Click(Index As Integer)

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim iCancela As Integer
Dim sCodRel As String
Dim objRelSel As New AdmRelSel

On Error GoTo Erro_mnuTESRel_Click

    Select Case Index

        Case MENU_TES_REL_EXTRATOTES
            objRelatorio.Rel_Menu_Executar ("Extrato de Tesouraria")

        Case MENU_TES_REL_EXTRATOBANC
            objRelatorio.Rel_Menu_Executar ("Extrato Bancário")

        Case MENU_TES_REL_POSAPLIC
            objRelatorio.Rel_Menu_Executar ("Posição das Aplicações")

        Case MENU_TES_REL_BORDEROPRE
            objRelatorio.Rel_Menu_Executar ("Borderô de Pré-Datados")

        Case MENU_TES_REL_PREVFLUXOCAIXA
            '???

        Case MENU_TES_REL_MOVFINRED
            objRelatorio.Rel_Menu_Executar ("Movimentação Financeira")

        Case MENU_TES_REL_MOVFINDET
            objRelatorio.Rel_Menu_Executar ("Movimentação Financeira Detalhada")

        Case MENU_TES_REL_ERROSCONAUTO
            objRelatorio.Rel_Menu_Executar ("Erros na Conciliação Automática")

        Case MENU_TES_REL_CONCPEND
            objRelatorio.Rel_Menu_Executar ("Conciliações Pendentes")

        Case MENU_TES_REL_OUTROS
            lErro = Chama_Tela("RelSelecionar", objRelSel, MODULO_TESOURARIA)
            If (lErro <> SUCESSO) Then Error 7716

            If (objRelSel.iCancela <> SUCESSO) Then Exit Sub

            'Prosseguir executando relatório
            objRelatorio.Rel_Menu_Executar (objRelSel.sCodRel)

        Case MENU_TES_REL_GERREL
            Sistema_EditarRel ("")

        '####################################
        'Inserido por Wagner
        Case MENU_TES_REL_PLANILHAS
            lErro = Chama_Tela("PlanilhasSelecionar", MODULO_TESOURARIA)
            If (lErro <> 0) Then Error 7058
        '####################################
        
    End Select

        Exit Sub

Erro_mnuTESRel_Click:

    Select Case Err

        Case 7716

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165251)

    End Select

    Exit Sub

End Sub

Private Sub mnuTESRot_Click(Index As Integer)

    Select Case Index

        Case MENU_TES_ROT_LIMPAARQ
            '???

        Case MENU_TES_ROT_RECEXTRATOCONCILIA
            Call Chama_Tela("ExtratoBancarioCNAB")

        Case MENU_TES_ROT_CONCILIAEXTRATO
            Call Chama_Tela("ConciliarExtratoBancario")

    End Select

End Sub

Private Sub Primeiro_Click()

Dim lErro As Long

    lErro = Seta_Click(BOTAO_PRIMEIRO)

End Sub

Private Sub Proximo_Click()

Dim lErro As Long

    lErro = Seta_Click(BOTAO_PROXIMO)

End Sub

Private Sub SuporteOnline_Click()
    Call mnuAjudaSuporte_Click
End Sub

Private Sub Timer1_Timer()

    If GL_lTransacao = 0 And Timer1.Interval > 0 Then
        Call Chama_Tela("AvisoWFW", 0)
    End If
    
End Sub

Private Sub Timer2_Timer()
    iCountBkp = iCountBkp + 1
    'Só executa o teste para ver se tem ou não que fazer o backup se:
    '1 - Já passou 20 minutos do último teste
    '2 - Se não tem nenhum transação aberta que possa ser prejudicada pela demora do backup
    '3 - Se o timer está ativo por conta de um backup habilitado
    '4 - Se não iniciou o processo, pode estar no meio de uma execução
    If iCountBkp >= 20 And GL_lTransacao = 0 And Timer2.Interval > 0 And giExeBkp = DESMARCADO Then
        iCountBkp = 0
        giExeBkp = MARCADO
        Call CF("Backup_Executa")
        giExeBkp = DESMARCADO
    End If
End Sub

Private Sub Ultimo_Click()

Dim lErro As Long

    lErro = Seta_Click(BOTAO_ULTIMO)

End Sub

Private Sub Anterior_Click()

Dim lErro As Long

    lErro = Seta_Click(BOTAO_ANTERIOR)

End Sub

Private Sub mnuArqSair_Click()

    Unload Me

End Sub

Private Function Keep_Rotinas_Alive() As Long

Dim objAux As Object
Dim lErro As Long
    
    
On Error GoTo Erro_Keep_Rotinas_Alive

    lErro = CF("Retorna_ColFiliais")
    If lErro <> SUCESSO Then gError 55177

    Set objAux = CreateObject("GlobaisAdm.AdmAdm")
    If objAux Is Nothing Then gError 55178
    GL_objKeepAlive.Add objAux

    If SGE_DATA_SIMULADO = DATA_NULA Then
        gdtDataHoje = Date
    Else
        gdtDataHoje = SGE_DATA_SIMULADO
    End If
    gdtDataAtual = gdtDataHoje

    Set objAux = CreateObject("AdmLib.Adm")
    If objAux Is Nothing Then gError 55180
    GL_objKeepAlive.Add objAux

'    Call Rotinas_Pre_Carga
    
    Set objAux = CreateObject("RotinasMAT.ClassMATSelect")
    If objAux Is Nothing Then gError 60795
    GL_objKeepAlive.Add objAux

'    Set objAux = CreateObject("GlobaisTelasEst.CTClasseUM")
'    If objAux Is Nothing Then Error 60796
'    GL_objKeepAlive.Add objAux
'
'    Set objAux = CreateObject("GlobaisTelasFAT.CTClientes")
'    If objAux Is Nothing Then Error 60797
'    GL_objKeepAlive.Add objAux
'
'    Set objAux = CreateObject("GlobaisTelasCPR.CTNFFATPAG")
'    If objAux Is Nothing Then Error 60798
'    GL_objKeepAlive.Add objAux
'
'    Set objAux = CreateObject("GlobaisTelasCTB.CTLancamentosAt")
'    If objAux Is Nothing Then Error 60799
'    GL_objKeepAlive.Add objAux
'
'    Set objAux = CreateObject("TelasEst.ClassTelasEst")
'    If objAux Is Nothing Then Error 60800
'    GL_objKeepAlive.Add objAux
'
    Set objAux = New EstInicial1
    If objAux Is Nothing Then gError 59348
    Load objAux
    
    '***** Edicao Telas
    Set gobjEstInicial = objAux
    Set gobjmenuEdicao = mnuArqSub(MENU_ARQ_EDICAO)
    Call CF("EdicaoTela_Le")

    '***** fim Edicao Telas

'    Set objAux = CreateObject("TelasFat.ClassTelasFat")
'    If objAux Is Nothing Then Error 60801
'    GL_objKeepAlive.Add objAux
'
'    Set objAux = CreateObject("TelasCpr.ClassTelasCpr")
'    If objAux Is Nothing Then Error 60802
'    GL_objKeepAlive.Add objAux

'    Set objAux = CreateObject("Telas.ClassTelasCTB")
'    If objAux Is Nothing Then Error 60803
'    GL_objKeepAlive.Add objAux
'
'    Set objAux = CreateObject("TelasAdm.ClassTelasAdm")
'    If objAux Is Nothing Then Error 60804
'    GL_objKeepAlive.Add objAux
'
    Set objAux = CreateObject("Adrelvb.Classrelaux")
    If objAux Is Nothing Then gError 59311
    GL_objKeepAlive.Add objAux

    Set objAux = CreateObject("GlobaisPV.ClassPedidoDeVenda")
    If objAux Is Nothing Then gError 59352
    GL_objKeepAlive.Add objAux
    
    Set objAux = gobjCRFAT
    If objAux Is Nothing Then gError 211089
    Set objAux = gobjMAT
    If objAux Is Nothing Then gError 211089
    Set objAux = gobjCP
    If objAux Is Nothing Then gError 211089
    Set objAux = gobjCR
    If objAux Is Nothing Then gError 211089
    Set objAux = gobjTES
    If objAux Is Nothing Then gError 211089
    Set objAux = gobjEST
    If objAux Is Nothing Then gError 211089
    Set objAux = gobjFAT
    If objAux Is Nothing Then gError 211089
    Set objAux = gobjTributacao
    If objAux Is Nothing Then gError 211089
    Set objAux = gcolUFs
    If objAux Is Nothing Then gError 211089
    Set objAux = gobjCOM
    If objAux Is Nothing Then gError 211089
    If gcolModulo.Ativo(MODULO_LOJA) = MODULO_ATIVO Then Set objAux = gobjLoja
    If objAux Is Nothing Then gError 211089
    Set objAux = Nothing
        
    Keep_Rotinas_Alive = SUCESSO
    
    Exit Function

Erro_Keep_Rotinas_Alive:

    Keep_Rotinas_Alive = gErr
    
    Select Case gErr
    
        Case 55177, 55178, 55180, 59311, 59348, 60795, 60796, 60797, 60798, 60799, 60800, 60801, 60802, 60803, 60804
    
        Case 211089
            Call Rotina_Erro(vbOKOnly, "ERRO_INICIALIZACAO_CONFIG_SISTEMA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165252)
    
    End Select
    
    Exit Function

End Function

Private Function Retorna_ColFiliais() As Long

Dim lErro As Long

On Error GoTo Erro_Retorna_ColFiliais

    Set gcolFiliais = New Collection
    
    lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, gcolFiliais)
    If lErro <> SUCESSO Then Error 55179

    Retorna_ColFiliais = SUCESSO
    
    Exit Function

Erro_Retorna_ColFiliais:

    Retorna_ColFiliais = Err

    Select Case Err
    
        Case 55179
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165253)

    End Select

    Exit Function

End Function

Private Sub MenuCadCTB_Contabil_ExtraContabil()
'torna visivel/invisivel as opcoes do menu de cadastros do CTB relativo a associacao de conta x centro de custo contabil/extra-contabil

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
            mnuCTBCad(MENU_CTB_CAD_ASSOCCCL).Visible = True
            mnuCTBCad(MENU_CTB_CAD_ASSOCCCLCTB).Visible = False
            mnuCTBCad(MENU_CTB_CAD_RATEIOOFF).Visible = True
        ElseIf giSetupUsoCcl = CCL_USA_CONTABIL Then
            mnuCTBCad(MENU_CTB_CAD_ASSOCCCL).Visible = False
            mnuCTBCad(MENU_CTB_CAD_ASSOCCCLCTB).Visible = True
            mnuCTBCad(MENU_CTB_CAD_RATEIOOFF).Visible = True
        Else
            mnuCTBCad(MENU_CTB_CAD_ASSOCCCL).Visible = False
            mnuCTBCad(MENU_CTB_CAD_ASSOCCCLCTB).Visible = False
            mnuCTBCad(MENU_CTB_CAD_RATEIOOFF).Visible = False
        End If
    
    End If
    
End Sub

'FERNANDO subir para ADMselect
Function MenuItem_Le_Titulo(sTitulo As String, objMenuItem As ClassMenuItens) As Long
'Lê item de menu a partir do nome da tela

Dim tMenuItens As typeMenuItens
Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_MenuItem_Le_Titulo

    'Iniciliza comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 25961

    tMenuItens.sNomeTela = String(STRING_NOME_TELA, 0)
    tMenuItens.sNomeControle = String(STRING_NOME_CONTROLE, 0)
    tMenuItens.sNomeControlePai = String(STRING_NOME_CONTROLE, 0)
    tMenuItens.sSiglaRotina = String(STRING_SIGLA_ROTINA, 0)
    tMenuItens.sTitulo = String(STRING_TITULO_MENU, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Identificador, Titulo, SiglaRotina, NomeTela, NomeControle, IndiceControle, NomeControlePai, IndiceControlePai FROM MenuItens WHERE Titulo = ? ", _
        tMenuItens.iIdentificador, tMenuItens.sTitulo, tMenuItens.sSiglaRotina, tMenuItens.sNomeTela, tMenuItens.sNomeControle, tMenuItens.iIndiceControle, tMenuItens.sNomeControlePai, tMenuItens.iIndiceControlePai, sTitulo)
    If lErro <> AD_SQL_SUCESSO Then Error 25962

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25963
    If lErro = AD_SQL_SEM_DADOS Then Error 25964

    objMenuItem.iIdentificador = tMenuItens.iIdentificador
    objMenuItem.sTitulo = tMenuItens.sTitulo
    objMenuItem.sSiglaRotina = tMenuItens.sSiglaRotina
    objMenuItem.sNomeTela = tMenuItens.sNomeTela
    objMenuItem.sNomeControle = tMenuItens.sNomeControle
    objMenuItem.iIndiceControle = tMenuItens.iIndiceControle
    objMenuItem.sNomeControlePai = tMenuItens.sNomeControlePai
    objMenuItem.iIndiceControlePai = tMenuItens.iIndiceControlePai

    Call Comando_Fechar(lComando)

    MenuItem_Le_Titulo = SUCESSO

    Exit Function

Erro_MenuItem_Le_Titulo:

    MenuItem_Le_Titulo = Err

    Select Case Err

        Case 25961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25962, 25963
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MENUITENS", Err)

        Case 25964
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MENUITEM_NAO_CADASTRADO", Err, sTitulo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165254)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'CODIGO PARA O MENU DO FIS

Private Sub mnuFISCad_Click(Index As Integer)
'Telas de Cadastro

    Select Case Index

        Case MENU_FIS_CAD_NATUREZAOP
            Call Chama_Tela("NaturezaOperacao")
        
        Case MENU_FIS_CAD_TIPOTRIB
            Call Chama_Tela("TipoDeTributacao")
        
        Case MENU_FIS_CAD_EXCECOESICMS
            Call Chama_Tela("ExcecoesICMS")

        Case MENU_FIS_CAD_EXCECOESIPI
            Call Chama_Tela("ExcecoesIPI")

        Case MENU_FIS_CAD_TRIBUTACAOFORN
            Call Chama_Tela("PadraoTribEntrada")

        Case MENU_FIS_CAD_TRIBUTACAOCLI
            Call Chama_Tela("PadraoTribSaida")
        
        Case MENU_FIS_CAD_TIPOREGAPURACAOICMS
            Call Chama_Tela("TiposRegApuracaoICMS")

        Case MENU_FIS_CAD_TIPOREGAPURACAOIPI
            Call Chama_Tela("TiposRegApuracaoIPI")
        
    End Select
    
End Sub

Private Sub mnuFISMov_Click(Index As Integer)
'Telas do Movimentos
    
    Select Case Index
    
        Case MENU_FIS_MOV_REGENTRADA
            Call Chama_Tela("EdicaoRegEntrada")
          
        Case MENU_FIS_MOV_REGSAIDA
            Call Chama_Tela("EdicaoRegSaida")
       
    End Select
    
End Sub

Private Sub mnuFISMovICMS_Click(Index As Integer)
'Telas do Movimentos

    Select Case Index
    
        Case MENU_FIS_MOV_ICMS_APURACAO
            Call Chama_Tela("ApuracaoICMS")
        
        Case MENU_FIS_MOV_ICMS_REGINVENTARIO
            Call Chama_Tela("EdicaoRegInventario")
        
        Case MENU_FIS_MOV_ICMS_REGEMITENTES
            Call Chama_Tela("RegESEmitentes")
    
        Case MENU_FIS_MOV_ICMS_REGCADPRODUTOS
            Call Chama_Tela("RegESCadProd")
        
        Case MENU_FIS_MOV_ICMS_LANCAPURACAO
            Call Chama_Tela("ApuracaoICMSItens")
            
        Case MENU_FIS_MOV_ICMS_GNRICMS
            Call Chama_Tela("CadastrarGNRICMS")
            
        Case MENU_FIS_MOV_ICMS_GUIASICMS
            Call Chama_Tela("GuiaICMS")
        
        Case MENU_FIS_MOV_ICMS_GUIASICMSST
            Call Chama_Tela("GuiaICMS", Nothing, 1)
        
    End Select
    
End Sub

Private Sub mnuFISMovIPI_Click(Index As Integer)
'Telas do Movimentos

    Select Case Index
    
        Case MENU_FIS_MOV_IPI_APURACAO
            Call Chama_Tela("ApuracaoIPI")
        
        Case MENU_FIS_MOV_IPI_LANCAPURACAO
            Call Chama_Tela("ApuracaoIPIItens")
        
    End Select

End Sub

Private Sub mnuFISRot_Click(Index As Integer)
'Telas das Rotinas

    Select Case Index

        Case MENU_FIS_ROT_FECHAMENTOLIVRO
            Call Chama_Tela("EscrituracaoFechamento")
        
        Case MENU_FIS_ROT_GERACAOARQICMS
            Call Chama_Tela("GeracaoArqICMSFIS")
            
        Case MENU_FIS_ROT_GERACAOARQIN86
            Call Chama_Tela("GeraIN86")
    
        Case MENU_FIS_ROT_REABERTURALIVRO
            Call Chama_Tela("ReaberturaLivroFIS")
    
        Case MENU_FIS_ROT_LIVREGESATUALIZA
            Call Chama_Tela("LivRegESAtualiza")
    
        Case MENU_FIS_ROT_SPEDFISCAL
            Call Chama_Tela("SpedFiscal")
    
        Case MENU_FIS_ROT_SPEDFISCALPIS
            Call Chama_Tela("SpedFiscalPis")
            
        Case MENU_FIS_ROT_DFC
            Call Chama_Tela("DFC")
    
        Case MENU_FIS_ROT_SPEDECF
            Call Chama_Tela("SpedECF")
    
    End Select
    
End Sub

'Private Function Jones_Gera_TabConfere_MenuItens() As Long
'
'Dim lErro As Long, iIndice As Integer, tmenuitem As typeMenuItens, sCaptionReal As String
'Dim alComando(1 To 10) As Long, lTransacao As Long, objObjeto As Object
'
'On Error GoTo Erro_Jones_Gera_TabConfere_MenuItens
'
'    'Abrir comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'
'        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
'        If alComando(iIndice) = 0 Then Error 12633
'
'    Next
'
'    'Inicializa transação
'    lTransacao = Transacao_AbrirExt(GL_lConexaoDic)
'    If lTransacao = 0 Then Error 11817
'
'    With tmenuitem
'
'        .sNomeControle = String(255, 0)
'        .sNomeControlePai = String(255, 0)
'        .sTitulo = String(255, 0)
'        .sNomeTela = String(255, 0)
'
'        lErro = Comando_Executar(alComando(1), "SELECT Identificador, Titulo, NomeControle, IndiceControle, NomeControlePai, IndiceControlePai, NomeTela FROM MenuItens", _
'            .iIdentificador, .sTitulo, .sNomeControle, .iIndiceControle, .sNomeControlePai, .iIndiceControlePai, .sNomeTela)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then Error 9999
'
'    lErro = Comando_BuscarProximo(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9999
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        Set objObjeto = Me.Controls(tmenuitem.sNomeControle)
'        If tmenuitem.iIndiceControle <> 0 Then
'            sCaptionReal = objObjeto(tmenuitem.iIndiceControle).Caption
'        Else
'            sCaptionReal = objObjeto.Caption
'        End If
'        With tmenuitem
'            lErro = Comando_Executar(alComando(2), "INSERT INTO JonesMenuItens (Identificador, Titulo, NomeControle, IndiceControle, NomeControlePai, IndiceControlePai, NomeTela, CaptionReal) VALUES (?,?,?,?,?,?,?,?)", _
'                .iIdentificador, .sTitulo, .sNomeControle, .iIndiceControle, .sNomeControlePai, .iIndiceControlePai, .sNomeTela, sCaptionReal)
'        End With
'
'        If lErro <> AD_SQL_SUCESSO Then
'            Error 9999
'        End If
'
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9999
'
'    Loop
'
'    'Finaliza Transação
'    lErro = Transacao_CommitExt(lTransacao)
'    If lErro <> AD_SQL_SUCESSO Then Error 11836
'
'    'Libera comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'
'        Call Comando_Fechar(alComando(iIndice))
'
'    Next
'
'    Jones_Gera_TabConfere_MenuItens = SUCESSO
'
'    Exit Function
'
'Erro_Jones_Gera_TabConfere_MenuItens:
'
'    sCaptionReal = "<lixo>"
'
'    Resume Next
'
'    Jones_Gera_TabConfere_MenuItens = Err
'
'    Select Case Err
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165255)
'
'    End Select
'
'    Call Transacao_RollbackExt(lTransacao)
'
'    'Libera comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'
'        Call Comando_Fechar(alComando(iIndice))
'
'    Next
'
'    Exit Function
'
'End Function


Private Sub mnuPCPCad_Click(Index As Integer)

    Select Case Index
    
        Case MENU_PCP_CAD_PRODUTOS
            Call Chama_Tela("Produto")
        
        Case MENU_PCP_CAD_ALMOXARIFADOS
            Call Chama_Tela("Almoxarifado")
                    
        Case MENU_PCP_CAD_KIT
            Call Chama_Tela("Kit")
            
        Case MENU_PCP_CAD_LOTERASTRO
            Call Chama_Tela("RastreamentoLote")
        
        Case MENU_PCP_CAD_PRODALM
            Call Chama_Tela("EstoqueProduto")
                    
        Case MENU_PCP_CAD_EMBALAGEM
            Call Chama_Tela("Embalagem")
        
        Case MENU_PCP_CAD_MAQUINA
            Call Chama_Tela("Maquinas")
                    
    End Select

End Sub

Private Sub mnuPCPCadTA_Click(Index As Integer)

    Select Case Index

        Case MENU_PCP_CAD_TA_ESTOQUE
            Call Chama_Tela("Estoque")

        Case MENU_PCP_CAD_TA_TIPOPROD
            Call Chama_Tela("TipoProduto")

        Case MENU_PCP_CAD_TA_UNIDADEMED
            Call Chama_Tela("ClasseUM")

        Case MENU_PCP_CAD_TA_ESTOQUEINI
            Call Chama_Tela("EstoqueInicial")

        Case MENU_PCP_CAD_TA_CATEGORIAPROD
            Call Chama_Tela("CategoriaProduto")
        
        Case MENU_PCP_CAD_TA_PRODUTOEMBALAGEM
            Call Chama_Tela("ProdutoEmbalagem")
            
        Case MENU_PCP_CAD_TA_TESTESQUALIDADE
            Call Chama_Tela("TestesQualidade")
            
        'Inserido por Jorge Specian
        '--------------------------------
        Case MENU_PCP_CAD_TA_COMPETENCIAS
            Call Chama_Tela("Competencias")
            
        Case MENU_PCP_CAD_TA_CENTROSDETRABALHOS
            Call Chama_Tela("CentrodeTrabalho")
            
        Case MENU_PCP_CAD_TA_TAXADEPRODUCAO
            Call Chama_Tela("TaxaDeProducao")
            
        Case MENU_PCP_CAD_TA_ROTEIROSDEFABRICACAO
            Call Chama_Tela("RoteirosDeFabricacao")
            
        Case MENU_PCP_CAD_TA_TIPOSDEMAODEOBRA
            Call Chama_Tela("TiposDeMaodeObra")
        
        Case MENU_PCP_CAD_TA_USUPRODARTX
            Call Chama_Tela("UsuProdArtlux")
        
        '--------------------------------
        Case MENU_PCP_CAD_TA_CERTIFICADOS
            Call Chama_Tela("Certificados")
        
        Case MENU_PCP_CAD_TA_CURSOS
            Call Chama_Tela("Cursos")
        
    End Select

End Sub

Private Sub mnuPCPConCad_Click(Index As Integer)
        
Dim colSelecao As New Collection

    Select Case Index
        
        Case MENU_PCP_CON_CAD_PRODUTOS
            Call Chama_Tela("ProdutoLista_Consulta", colSelecao)

        Case MENU_PCP_CON_CAD_ALMOXARIFADO
            Call Chama_Tela("AlmoxarifadoLista_Consulta", colSelecao)
        
        Case MENU_PCP_CON_CAD_KIT
            Call Chama_Tela("KitLista")

        Case MENU_PCP_CON_CAD_TIPOPROD
            Call Chama_Tela("TipoProdutoLista")

        Case MENU_PCP_CON_CAD_CLASSEUM
            Call Chama_Tela("ClasseUMLista", colSelecao)

        '######################################
        'Inserido por Wagner
        Case MENU_PCP_CON_CAD_COMPETENCIA
            Call Chama_Tela("CompetenciasLista", colSelecao)
        
        Case MENU_PCP_CON_CAD_MAQUINA
            Call Chama_Tela("MaquinasLista", colSelecao)
        
        Case MENU_PCP_CON_CAD_CT
            Call Chama_Tela("CentrodeTrabalhoLista", colSelecao)
        
        Case MENU_PCP_CON_CAD_TIPOMAODEOBRA
            Call Chama_Tela("TiposDeMaodeObraLista", colSelecao)
        '######################################
        
    End Select


End Sub

Private Sub mnuPCPConEst_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
                
        Case MENU_PCP_CON_EST_ESTPROD
            Call Chama_Tela("EstProdLista_Consulta")

        Case MENU_PCP_CON_EST_ESTPRODFILIAL
            Call Chama_Tela("EstProdFilialLista_Cons", colSelecao)

        Case MENU_PCP_CON_EST_ESTPRODTERC
            Call Chama_Tela("EstProdTercLista_Consulta", colSelecao)

    End Select

End Sub

Private Sub mnuPCPConfig_Click(Index As Integer)

Dim objForm As Form
Dim lErro As Long

On Error GoTo Erro_mnuPCPConfig_Click

    Select Case Index

        Case MENU_PCP_CONFIG_SEGMENTOS
            
            For Each objForm In Forms
                If Not (objForm Is Me) And Not (objForm Is gobjEstInicial) Then Error 59361
            Next

            Call Chama_Tela_Modal("SegmentosMAT")
        
        Case MENU_PCP_CONFIG_CONFIGURACAO
            Call Chama_Tela("ConfiguraEST")

    End Select

    Exit Sub
    
Erro_mnuPCPConfig_Click:

    Select Case Err

        Case 59361
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FECHAR_JANELAS_FILHAS", Err, Error)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165256)

    End Select

    Exit Sub

End Sub

Private Sub mnuPCPConPro_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index

        Case MENU_PCP_CON_PRO_OP
            Call Chama_Tela("OrdemProducaoLista", colSelecao)
            
        Case MENU_PCP_CON_PRO_OP_BAIXA
            Call Chama_Tela("OrdemProdBaixadasLista", colSelecao)
            
        Case MENU_PCP_CON_PRO_EMPENHO
            Call Chama_Tela("EmpenhoLista", colSelecao)

        Case MENU_PCP_CON_PRO_ITENSOP
            Call Chama_Tela("ItemOrdemProducaoLista", colSelecao)
        
        Case MENU_PCP_CON_PRO_REQPROD
            colSelecao.Add MOV_EST_REQ_PRODUCAO
            colSelecao.Add MOV_EST_REQ_PRODUCAO_BENEF3
            colSelecao.Add MOV_EST_REQ_PRODUCAO_OUTROS
            
            Call Chama_Tela("MovEstoqueLista1", colSelecao)
            
        Case MENU_PCP_CON_PRO_PRODUCAO
            colSelecao.Add MOV_EST_PRODUCAO
            colSelecao.Add MOV_EST_PRODUCAO_BENEF3
            colSelecao.Add MOV_EST_PRODUCAO_OUTROS
            
            Call Chama_Tela("MovEstoqueLista1", colSelecao)

        Case MENU_PCP_CON_PRO_REQPRODOP
            colSelecao.Add MOV_EST_REQ_PRODUCAO
            colSelecao.Add MOV_EST_REQ_PRODUCAO_BENEF3
            
            Call Chama_Tela("MovEstoqueOPLista", colSelecao)
       
        Case MENU_PCP_CON_PRO_PRODUCAOOP
            colSelecao.Add MOV_EST_PRODUCAO
            colSelecao.Add MOV_EST_PRODUCAO_BENEF3

            Call Chama_Tela("MovEstoqueOPLista", colSelecao)

        '########################################
        'Inserido por Wagner
        Case MENU_PCP_CON_PRO_ROTEIRO
            Call Chama_Tela("RoteirosDeFabricacaoLista", colSelecao)
            
        Case MENU_PCP_CON_PRO_TAXA
            Call Chama_Tela("TaxaDeProducaoLista", colSelecao)
            
        Case MENU_PCP_CON_PRO_PMP
            Call Chama_Tela("PMPLista", colSelecao)
        '########################################
        
    End Select

End Sub

Private Sub mnuPCPConRastro_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
                
        Case MENU_PCP_CON_RASTRO_LOTES
            Call Chama_Tela("RastroLoteLista", colSelecao)
                
        Case MENU_PCP_CON_RASTRO_SALDOS
            Call Chama_Tela("RastroLoteSaldoLista", colSelecao)

        Case MENU_PCP_CON_RASTRO_MOVIMENTOS
            Call Chama_Tela("RastroMovEstoqueLista", colSelecao)
        
    End Select

End Sub

Private Sub mnuPCPMov_Click(Index As Integer)

Dim sNomeTela As String

    Select Case Index

        Case MENU_PCP_MOV_MANUAL
            Call Chama_Tela("OrdemProducao")

        Case MENU_PCP_MOV_AUTOMATICO
            Call Chama_Tela("GeracaoOP")

        Case MENU_PCP_MOV_EMPENHO
            Call Chama_Tela("Empenho")
        
        Case MENU_PCP_MOV_ORDPRODBLOQUEIO
            Call Chama_Tela("OrdemProducaoBloqueio")
        
        Case MENU_PCP_MOV_REQPROD
            sNomeTela = "ProducaoSaida"
            Call CF("Chama_Tela_ProducaoSaida", sNomeTela)
            Call Chama_Tela(sNomeTela)

        Case MENU_PCP_MOV_TRANSF
            Call Chama_Tela("Transfer")

        Case MENU_PCP_MOV_PRODENT
            Call Chama_Tela("ProducaoEntrada")

        'Inserido por Jorge Specian
        '--------------------------------
        Case MENU_PCP_MOV_APONTAMENTOPRODUCAO
            Call Chama_Tela("ApontamentoProducao")
        '--------------------------------
        
        Case MENU_PCP_MOV_PMP
            Call Chama_Tela("PMP")
            
        Case MENU_PCP_MOV_ORDEMCORTE
            Call Chama_Tela("OrdemCorte")
        
        Case MENU_PCP_MOV_OCARTX
            Call Chama_Tela("OCArtlux")
        
        Case MENU_PCP_MOV_OCMANUALARTX
            Call Chama_Tela("OCManualArtlux")
        
        Case MENU_PCP_MOV_REQPROD_DPACK
            Call Chama_Tela("ProducaoSaida")
    
    End Select

End Sub


Private Sub mnuPCPRel_Click(Index As Integer)

Dim lErro As Long
Dim iCancela As Integer
Dim sCodRel As String
Dim objRelSel As New AdmRelSel
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_mnuPCPRel_Click

    Select Case Index
    
        Case MENU_PCP_REL_OP
            objRelatorio.Rel_Menu_Executar ("Relação das Ordens de Produção")

        Case MENU_PCP_REL_EMPENHOS
            objRelatorio.Rel_Menu_Executar ("Lista dos Empenhos")

        Case MENU_PCP_REL_LISTAFALTAS
            objRelatorio.Rel_Menu_Executar ("Lista de Faltas")
        
        Case MENU_PCP_REL_DISTREATOR
            objRelatorio.Rel_Menu_Executar ("Distribuição de Matéria-Prima por Máquina")
        
        Case MENU_PCP_REL_PRODXOP
            objRelatorio.Rel_Menu_Executar ("Produtos x Ordens de Produção")
       
        Case MENU_PCP_REL_MOVESTOP
            objRelatorio.Rel_Menu_Executar ("Movimentos de Estoque para cada Ordem de Produção")

        Case MENU_PCP_REL_OPXREQPROD
            objRelatorio.Rel_Menu_Executar ("Ordem de Produção x Requisição p/ Produção")
                
        Case MENU_PCP_REL_ANALRENDOP
            objRelatorio.Rel_Menu_Executar ("Análise de Rendimento por Ordem de Produção")
        
        Case MENU_PCP_REL_PVENDAXPCONSUMO
            objRelatorio.Rel_Menu_Executar ("Previsão de Vendas x Previsão de Consumo")

        Case MENU_PCP_REL_FORPADRAOCUSTO
            objRelatorio.Rel_Menu_Executar ("Fórmula Padrão para Custo")
        
        Case MENU_PCP_REL_OUTROS
            lErro = Chama_Tela("RelSelecionar", objRelSel, MODULO_ESTOQUE)
            If (lErro <> 0) Then Error 7058

            If (objRelSel.iCancela <> 0) Then Exit Sub

            'Prosseguir executando relatório
            objRelatorio.Rel_Menu_Executar (objRelSel.sCodRel)

        Case MENU_PCP_REL_GERREL
            Sistema_EditarRel ("")

        '####################################
        'Inserido por Wagner
        Case MENU_PCP_REL_PLANILHAS
            lErro = Chama_Tela("PlanilhasSelecionar", MODULO_ESTOQUE)
            If (lErro <> 0) Then Error 7058
        '####################################
        
    End Select

    Exit Sub

Erro_mnuPCPRel_Click:

    Select Case Err

        Case 7058

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165257)

    End Select

    Exit Sub

End Sub

Private Sub mnuPCPRelCad_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio
    
    Select Case Index

        Case MENU_PCP_REL_CAD_PRODUTOS
            objRelatorio.Rel_Menu_Executar ("Relação de Produtos")
        
        Case MENU_PCP_REL_CAD_ALMOXARIFADO
            objRelatorio.Rel_Menu_Executar ("Relação de Almoxarifados")

        Case MENU_PCP_REL_CAD_KITS
            objRelatorio.Rel_Menu_Executar ("Relação de Kits")
        
        Case MENU_PCP_REL_CAD_UTILPROD
            objRelatorio.Rel_Menu_Executar ("Utilização do Produto")

    End Select

End Sub

Private Sub mnuPCPRot_Click(Index As Integer)

    Select Case Index
        
        Case MENU_PCP_ROT_CUSTOMEDIOPRODUCAO
            Call Chama_Tela("CustoProducao")
    
        Case MENU_PCP_ROT_MRP
            Call Chama_Tela("MRP")
        
    End Select

End Sub

Private Sub mnuPRJMov_Click(Index As Integer)

    Select Case Index

        Case MENU_PRJ_MOV_PROPOSTA
            Call Chama_Tela("PropostaPRJ")
            
        Case MENU_PRJ_MOV_CONTRATO
            Call Chama_Tela("ContratoPRJ")
            
        Case MENU_PRJ_MOV_PAGTO
            Call Chama_Tela("PagamentoPRJ")
        
        Case MENU_PRJ_MOV_RECEB
            Call Chama_Tela("RecebimentoPRJ")
            
        Case MENU_PRJ_MOV_APONT
            Call Chama_Tela("ApontamentoPRJ")
    
    End Select
        
    Exit Sub
    
End Sub

Private Sub mnuPRJConCad_Click(Index As Integer)

    Select Case Index

        Case MENU_PRJ_CON_CAD_PRODUTO
            Call Chama_Tela("ProdutoLista_Consulta")
            
        Case MENU_PRJ_CON_CAD_CLIENTE
            Call Chama_Tela("ClientesLista")
            
        Case MENU_PRJ_CON_CAD_FORNECEDOR
            Call Chama_Tela("FornecedorLista")
        
    End Select

End Sub

Private Sub mnuPRJCon_Click(Index As Integer)

    Select Case Index

        Case MENU_PRJ_CON_PROJETO
            Call Chama_Tela("ProjetosLista")
            
        Case MENU_PRJ_CON_PROPOSTA
            Call Chama_Tela("PRJPropostasLista")
            
        Case MENU_PRJ_CON_ETAPA
            Call Chama_Tela("PRJEtapasLista")
            
        Case MENU_PRJ_CON_PAGTO
            Call Chama_Tela("PagamentoPRJLista")
        
        Case MENU_PRJ_CON_RECEB
            Call Chama_Tela("RecebimentoPRJLista")
        
    End Select

End Sub

Private Sub mnuPRJCad_Click(Index As Integer)
    
        Select Case Index
            
            Case MENU_PRJ_CAD_CLIENTE
                If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                    Call Chama_Tela("ClienteLoja")
                Else
                    Call Chama_Tela("Clientes")
                End If
            
            Case MENU_PRJ_CAD_PRODUTO
                If giLocalOperacao <> LOCALOPERACAO_CAIXA_CENTRAL Then
                    Call Chama_Tela("Produto")
                End If
         
            Case MENU_PRJ_CAD_FORNECEDOR
                Call Chama_Tela("Fornecedores")
    
            Case MENU_PRJ_CAD_PROJETO
                Call Chama_Tela("Projetos")

            Case MENU_PRJ_CAD_ETAPA
                Call Chama_Tela("EtapaPRJ")
                
            Case MENU_PRJ_CAD_ORGANOGRAMA
                Call Chama_Tela("OrganogramaPRJ")
            
            
        End Select
    
End Sub

Private Sub mnuPRJRel_Click(Index As Integer)

Dim lErro As Long
Dim iCancela As Integer
Dim sCodRel As String
Dim objRelSel As New AdmRelSel
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_mnuPRJRel_Click

    Select Case Index
    
        Case MENU_PRJ_REL_PREV_REAL
            objRelatorio.Rel_Menu_Executar ("Realizado x Previsto do Projeto")
        
        Case MENU_PRJ_REL_FLUXO
            objRelatorio.Rel_Menu_Executar ("Fluxo Financeiro do Projeto")
        
        Case MENU_PRJ_REL_MAT
            objRelatorio.Rel_Menu_Executar ("Materiais Utilizados no Projeto")
        
        Case MENU_PRJ_REL_MO
            objRelatorio.Rel_Menu_Executar ("Mãos de obra Utilizadas no Projeto")
        
        Case MENU_PRJ_REL_ACOMP
            objRelatorio.Rel_Menu_Executar ("Acompanhamento do Projeto")
            
        Case MENU_PRJ_REL_OUTROS
            lErro = Chama_Tela("RelSelecionar", objRelSel, MODULO_PROJETO)
            If (lErro <> 0) Then gError 189138

            If (objRelSel.iCancela <> 0) Then Exit Sub

            'Prosseguir executando relatório
            objRelatorio.Rel_Menu_Executar (objRelSel.sCodRel)

        Case MENU_PRJ_REL_GERREL
            Sistema_EditarRel ("")

        Case MENU_PRJ_REL_PLANILHAS
            lErro = Chama_Tela("PlanilhasSelecionar", MODULO_PROJETO)
            If (lErro <> 0) Then gError 189139
        
    End Select

    Exit Sub

Erro_mnuPRJRel_Click:

    Select Case gErr

        Case 189138, 189139

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189142)

    End Select

    Exit Sub

End Sub

Private Sub mnuPRJRot_Click(Index As Integer)

    Select Case Index

        Case MENU_PRJ_ROT_EXPORTEXCEL
            Call Chama_Tela("PRJRelsEmExcel")
                       
    End Select
    
End Sub

Private Sub mnuPRJConfig_Click(Index As Integer)

Dim objForm As Form
Dim lErro As Long

On Error GoTo Erro_mnuPRJConfig_Click

    Select Case Index

        Case MENU_PRJ_CONFIG_SEGMENTOS
            
            For Each objForm In Forms
                If Not (objForm Is Me) And Not (objForm Is gobjEstInicial) Then gError 189140
            Next

            Call Chama_Tela_Modal("SegmentosPRJ")

    End Select

    Exit Sub
    
Erro_mnuPRJConfig_Click:

    Select Case gErr

        Case 189140
            Call Rotina_Erro(vbOKOnly, "ERRO_FECHAR_JANELAS_FILHAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189141)

    End Select

    Exit Sub

End Sub

Private Sub BotaoAnotacao_Click()

Dim lErro As Long, objAnotacao As New ClassAnotacoes
Dim obj As Object

On Error GoTo Erro_BotaoAnotacao_Click

Dim bApenasSaldosFAT As Boolean, bEST As Boolean

    If gobjEST.iPrioridadeProduto = 171 And LCase(left(gsUsuario, 5)) = "super" Then
    
        Set obj = CreateObject("RotinasEST.ClassESTGrava")
    
        Call obj.NF_Grava_Trib_E_MovEst
    
    Else

        If gi_ST_SetaIgnoraClick = 0 And (Not gobj_ST_TelaAtiva Is Nothing) And Indice.ListIndex > -1 Then
    
            'Extrai da tela o ID
            Call gobj_ST_TelaAtiva.Anotacao_Extrai(objAnotacao)
    
            If objAnotacao.sID <> "" Then
                Call Chama_Tela("Anotacoes", objAnotacao)
            Else
                Call Chama_Tela_Modal("Anotacoes2", objAnotacao)
                Call gobj_ST_TelaAtiva.Anotacao_Preenche(objAnotacao)
            End If
    
        End If
        
    End If

    Exit Sub
     
Erro_BotaoAnotacao_Click:

    Select Case gErr
          
        Case 438 'para telas que ainda nao tiveram anotacao_extrai implementada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165258)
     
    End Select
     
    Exit Sub

End Sub

Private Sub TransfereDados()

Dim lErro As Long, sTabela As String, sColuna As String, sTabelaAnt As String, sColunaAnt As String
Dim sCampos As String, lNumReg As Long, sQualDefDest As String, alComando(1 To 4) As Long, iIndice As Integer
Dim lTransacao As Long, lNumReg2 As Long

On Error GoTo Erro_TransfereDados

    sTabela = String(255, 0)
    sColuna = String(255, 0)
    sQualDefDest = "SGEDados1_Fox.dbo."
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 1111
    
    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 1111
    
    Next
    
    lErro = Comando_Executar(alComando(1), "SELECT INFORMATION_SCHEMA.COLUMNS.TABLE_NAME, COLUMN_NAME FROM INFORMATION_SCHEMA.TABLES, INFORMATION_SCHEMA.COLUMNS WHERE TABLE_TYPE = 'BASE TABLE' AND INFORMATION_SCHEMA.TABLES.TABLE_NAME = INFORMATION_SCHEMA.COLUMNS.TABLE_NAME ORDER BY INFORMATION_SCHEMA.COLUMNS.TABLE_NAME, COLUMN_NAME", sTabela, sColuna)
    If lErro <> AD_SQL_SUCESSO Then gError 1111
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1111
    
    If lErro = AD_SQL_SUCESSO Then
    
        sTabelaAnt = sTabela
    
        Do While lErro = AD_SQL_SUCESSO
        
            'se trocou de tabela
            If sTabelaAnt <> sTabela Then
            
                'verifica se existe conteudo na tabela origem
                lErro = Comando_Executar(alComando(2), "SELECT COUNT(*) FROM " & sTabelaAnt, lNumReg)
                If lErro <> AD_SQL_SUCESSO Then
                    MsgBox (sTabelaAnt)
                Else
                
                    lErro = Comando_BuscarProximo(alComando(2))
                    If lErro <> AD_SQL_SUCESSO Then gError 1111
                    
                    'se a tabela origem está vazia
                    If lNumReg <> 0 Then
                    
                        'verifica se existe conteudo pré-cadastrado na tabela destino
                        lErro = Comando_Executar(alComando(4), "SELECT COUNT(*) FROM " & sQualDefDest & sTabelaAnt, lNumReg2)
                        If lErro <> AD_SQL_SUCESSO Then
                            MsgBox (sTabelaAnt)
                        Else
                            
                            lErro = Comando_BuscarProximo(alComando(4))
                            If lErro <> AD_SQL_SUCESSO Then gError 1111
                                
                            If lNumReg2 = 0 Then
                                If UCase(sTabelaAnt) <> "DTPROPERTIES" And UCase(sTabelaAnt) <> "PEDIDOSDEVENDA" And UCase(sTabelaAnt) <> "NFISCAL" And UCase(sTabelaAnt) <> "PEDIDOSDEVENDABAIXADOS" Then
                                    lErro = Comando_Executar(alComando(3), "INSERT INTO " & sQualDefDest & sTabelaAnt & "(" & sCampos & ") (SELECT " & sCampos & " FROM " & sTabelaAnt & ")")
                                    If lErro <> AD_SQL_SUCESSO Then MsgBox (sTabelaAnt)
                                End If
                            End If
                        End If
                    End If
                
                End If
                
                sTabelaAnt = sTabela
                sCampos = ""
            
            End If
        
            If sCampos = "" Then
                sCampos = sColuna
            Else
                sCampos = sCampos & "," & sColuna
            End If
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1111
            
        Loop
    
        'verifica se existe conteudo pré-cadastrado na tabela destino
        lErro = Comando_Executar(alComando(2), "SELECT COUNT(*) FROM " & sQualDefDest & sTabelaAnt, lNumReg)
        If lErro <> AD_SQL_SUCESSO Then gError 1111
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 1111
            
        'se a tabela destino está vazia
        If lNumReg = 0 Then
        
            If UCase(sTabelaAnt) <> "DTPROPERTIES" And UCase(sTabelaAnt) <> "PEDIDOSDEVENDA" And UCase(sTabelaAnt) <> "NFISCAL" And UCase(sTabelaAnt) <> "PEDIDOSDEVENDABAIXADOS" Then
                lErro = Comando_Executar(alComando(3), "INSERT INTO " & sQualDefDest & sTabelaAnt & "(" & sCampos & ") (SELECT " & sCampos & " FROM " & sTabelaAnt & ")")
                If lErro <> AD_SQL_SUCESSO Then MsgBox (sTabelaAnt)
            End If
        End If
    
    End If
    
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 1111
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Sub
     
Erro_TransfereDados:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165259)
     
    End Select
     
    Exit Sub

End Sub

Private Sub IdentificaTabelasPreCadastradas()

Dim lErro As Long, sTabela As String
Dim sCampos As String, lNumReg As Long, sQualDefDest As String, alComando(1 To 3) As Long, iIndice As Integer

On Error GoTo Erro_IdentificaTabelasPreCadastradas

    
    Open "TabelasPreCadastradas.txt" For Output As #1

    sTabela = String(255, 0)
    sQualDefDest = "SGEPadrao_Versao_Nova.dbo."
    
    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 1111
    
    Next
    
    lErro = Comando_Executar(alComando(1), "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME", sTabela)
    If lErro <> AD_SQL_SUCESSO Then gError 1111
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1111
    
    Do While lErro = AD_SQL_SUCESSO
    
        'verifica se existe conteudo pré-cadastrado na tabela destino
        lErro = Comando_Executar(alComando(2), "SELECT COUNT(*) FROM " & sQualDefDest & sTabela, lNumReg)
        If lErro <> AD_SQL_SUCESSO Then gError 1111
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 1111
        
        'se a tabela destino nao está vazia
        If lNumReg <> 0 Then
        
            Print #1, sTabela
            
        End If
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1111
    
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Close
    
    Exit Sub
     
Erro_IdentificaTabelasPreCadastradas:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165260)
     
    End Select
     
    Close
    
    Exit Sub

End Sub

Private Sub CriaScriptCopiaDados()

Dim lErro As Long, sTabela As String, sColuna As String, sTabelaAnt As String, sColunaAnt As String
Dim sCampos As String, lNumReg As Long, sQualDefDest As String, alComando(1 To 3) As Long, iIndice As Integer
Dim lTransacao As Long

On Error GoTo Erro_CriaScriptCopiaDados

    Open "ScriptCopiaDados.txt" For Output As #1

    sTabela = String(255, 0)
    sColuna = String(255, 0)
    sQualDefDest = "SGEDados_Demo_Novo.dbo."
    
    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 1111
    
    Next
    
    lErro = Comando_Executar(alComando(1), "SELECT INFORMATION_SCHEMA.COLUMNS.TABLE_NAME, COLUMN_NAME FROM INFORMATION_SCHEMA.TABLES, INFORMATION_SCHEMA.COLUMNS WHERE TABLE_TYPE = 'BASE TABLE' AND INFORMATION_SCHEMA.TABLES.TABLE_NAME = INFORMATION_SCHEMA.COLUMNS.TABLE_NAME ORDER BY INFORMATION_SCHEMA.COLUMNS.TABLE_NAME, COLUMN_NAME", sTabela, sColuna)
    If lErro <> AD_SQL_SUCESSO Then gError 1111
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1111
    
    If lErro = AD_SQL_SUCESSO Then
    
        sTabelaAnt = sTabela
    
        Do While lErro = AD_SQL_SUCESSO
        
            'se trocou de tabela
            If sTabelaAnt <> sTabela Then
            
                Print #1, "DELETE FROM " & sQualDefDest & sTabelaAnt
                Print #1, "GO"
                Print #1, ""
                Print #1, "INSERT INTO " & sQualDefDest & sTabelaAnt & "(" & sCampos & ") (SELECT " & sCampos & " FROM " & sTabelaAnt & ")"
                Print #1, "GO"
                Print #1, ""
                
                sTabelaAnt = sTabela
                sCampos = ""
            
            End If
        
            If sCampos = "" Then
                sCampos = sColuna
            Else
                sCampos = sCampos & "," & sColuna
            End If
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1111
            
        Loop
    
        Print #1, "DELETE FROM " & sQualDefDest & sTabelaAnt
        Print #1, "GO"
        Print #1, ""
        Print #1, "INSERT INTO " & sQualDefDest & sTabelaAnt & "(" & sCampos & ") (SELECT " & sCampos & " FROM " & sTabelaAnt & ")"
        Print #1, "GO"
        Print #1, ""
    
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Close
    
    Exit Sub
     
Erro_CriaScriptCopiaDados:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165261)
     
    End Select
     
    Close
    
    Exit Sub

End Sub

Private Sub CriaScriptCopiaDadosDic()

Dim lErro As Long, sTabela As String, sColuna As String, sTabelaAnt As String, sColunaAnt As String
Dim sCampos As String, lNumReg As Long, sQualDefDest As String, alComando(1 To 3) As Long, iIndice As Integer
Dim lTransacao As Long

On Error GoTo Erro_CriaScriptCopiaDadosDic

    Open "ScriptCopiaDadosDic.txt" For Output As #1

    sTabela = String(255, 0)
    sColuna = String(255, 0)
    sQualDefDest = "SGEDic_Demo_Novo.dbo."
    
    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 1111
    
    Next
    
    lErro = Comando_Executar(alComando(1), "SELECT INFORMATION_SCHEMA.COLUMNS.TABLE_NAME, COLUMN_NAME FROM INFORMATION_SCHEMA.TABLES, INFORMATION_SCHEMA.COLUMNS WHERE TABLE_TYPE = 'BASE TABLE' AND INFORMATION_SCHEMA.TABLES.TABLE_NAME = INFORMATION_SCHEMA.COLUMNS.TABLE_NAME ORDER BY INFORMATION_SCHEMA.COLUMNS.TABLE_NAME, COLUMN_NAME", sTabela, sColuna)
    If lErro <> AD_SQL_SUCESSO Then gError 1111
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1111
    
    If lErro = AD_SQL_SUCESSO Then
    
        sTabelaAnt = sTabela
    
        Do While lErro = AD_SQL_SUCESSO
        
            'se trocou de tabela
            If sTabelaAnt <> sTabela Then
            
                Print #1, "DELETE FROM " & sQualDefDest & sTabelaAnt
                Print #1, "GO"
                Print #1, ""
                Print #1, "INSERT INTO " & sQualDefDest & sTabelaAnt & "(" & sCampos & ") (SELECT " & sCampos & " FROM " & sTabelaAnt & ")"
                Print #1, "GO"
                Print #1, ""
                
                sTabelaAnt = sTabela
                sCampos = ""
            
            End If
        
            If sCampos = "" Then
                sCampos = sColuna
            Else
                sCampos = sCampos & "," & sColuna
            End If
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1111
            
        Loop
    
        Print #1, "DELETE FROM " & sQualDefDest & sTabelaAnt
        Print #1, "GO"
        Print #1, ""
        Print #1, "INSERT INTO " & sQualDefDest & sTabelaAnt & "(" & sCampos & ") (SELECT " & sCampos & " FROM " & sTabelaAnt & ")"
        Print #1, "GO"
        Print #1, ""
    
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Close
    
    Exit Sub
     
Erro_CriaScriptCopiaDadosDic:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165262)
     
    End Select
     
    Close
    
    Exit Sub

End Sub

Private Sub SetKeyState(ByVal Key As Long, ByVal State As Boolean)

   Dim Keys(0 To 255) As Byte
   Call GetKeyState(Keys(0))
   Keys(Key) = Abs(CInt(State))
   Call SetKeyboardState(Keys(0))

End Sub

Private Property Get NumLock() As Boolean

   NumLock = GetKeyState(KeyCodeConstants.vbKeyNumlock) = 1

End Property

Private Property Let NumLock(ByVal Value As Boolean)

   Call SetKeyState(KeyCodeConstants.vbKeyNumlock, Value)

End Property

Sub TestKeys()

    If (GetKeyState(VK_NUMLOCK) And 1) = 0 Then 'if NUMLOCK Off then set On
        keybd_event VK_NUMLOCK, 0, KEYEVENTF_EXTENDEDKEY, 0 'toggles NumLock
        keybd_event VK_NUMLOCK, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        DoEvents
    End If
    
    If (GetKeyState(VK_NUMLOCK) And 1) = 1 Then 'if NUMLOCK On then set Off
        keybd_event VK_NUMLOCK, 0, KEYEVENTF_EXTENDEDKEY, 0 'toggles NumLock
        keybd_event VK_NUMLOCK, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        DoEvents
    End If

End Sub

Sub Testa_MenuItens()

Dim lErro As Long, lComando As Long
Dim iIdentificador As Integer, sTitulo As String, sNomeTela As String, sNomeControle As String, iIndiceControle As Integer
Dim sCaptionCtl As String, objObjeto As Object

On Error GoTo Erro_Testa_MenuItens

    Open "MenuItensTeste.txt" For Output As #1
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 1234
    
    sTitulo = String(255, 0)
    sNomeTela = String(255, 0)
    sNomeControle = String(255, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Identificador, Titulo, NomeTela, NomeControle, IndiceControle FROM MenuItens WHERE Separador = 0 and nometela is not null ORDER BY Identificador", iIdentificador, sTitulo, sNomeTela, sNomeControle, iIndiceControle)
    If lErro <> AD_SQL_SUCESSO Then gError 1234
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1234
    
    Do While lErro = AD_SQL_SUCESSO
    
        sCaptionCtl = "ERRO"
        Set objObjeto = Nothing
        
        Set objObjeto = Me.Controls(sNomeControle)
        
        If Not (objObjeto Is Nothing) Then
        
            If iIndiceControle = 0 Then
                sCaptionCtl = objObjeto.Caption
            Else
                sCaptionCtl = objObjeto(iIndiceControle).Caption
            End If
        End If
        
        Print #1, CStr(iIdentificador) & " - " & sTitulo & " - " & sNomeTela & " - " & sNomeControle & " - " & CStr(iIndiceControle) & " : " & sCaptionCtl
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1234
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    Close
    
    Exit Sub
     
Erro_Testa_MenuItens:

    Select Case gErr
          
        
        Case Else
            Resume 'Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165263)
     
    End Select
     
    Close
    
    Call Comando_Fechar(lComando)
    
    Exit Sub

End Sub

Sub Cidades_Cadastrar()

Dim sCidade As String, lErro As Long
Dim alComando(1 To 4) As Long, iIndice As Integer
Dim lTransacao As Long

On Error GoTo Erro_Cidades_Cadastrar

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 1111
    
    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 1111
    
    Next
    
    sCidade = String(255, 0)
    lErro = Comando_Executar(alComando(1), "SELECT cidade from enderecos where cidade <> '' group by cidade order by count(*) desc", sCidade)
    If lErro <> AD_SQL_SUCESSO Then gError 1111
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1111
    
    iIndice = 0
    
    Do While lErro = AD_SQL_SUCESSO
    
        iIndice = iIndice + 1
        
        lErro = Comando_Executar(alComando(2), "INSERT INTO Cidades (Codigo,Descricao) VALUES (?,?)", iIndice, sCidade)
        If lErro <> AD_SQL_SUCESSO Then gError 1111
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1111
    
    Loop
    
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 1111
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Sub
     
Erro_Cidades_Cadastrar:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165264)
     
    End Select
     
    Exit Sub

End Sub

Sub Testa_MenuItensPai()

Dim lErro As Long, lComando As Long
Dim sNomeControle As String, iIndiceControle As Integer
Dim sCaptionCtl As String, objObjeto As Object

On Error GoTo Erro_Testa_MenuItensPai

    Open "MenuItensPaiTeste.txt" For Output As #1
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 1234
    
    sNomeControle = String(255, 0)
    
    lErro = Comando_Executar(lComando, "SELECT distinct NomeControlePai, IndiceControlePai FROM MenuItens WHERE nomecontrolepai is not null ORDER BY NomeControlePai, IndiceControlePai", sNomeControle, iIndiceControle)
    If lErro <> AD_SQL_SUCESSO Then gError 1234
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1234
    
    Do While lErro = AD_SQL_SUCESSO
    
        sCaptionCtl = "ERRO"
        Set objObjeto = Nothing
        
        Set objObjeto = Me.Controls(sNomeControle)
        
        If Not (objObjeto Is Nothing) Then
        
            If iIndiceControle = 0 Then
                sCaptionCtl = objObjeto.Caption
            Else
                sCaptionCtl = objObjeto(iIndiceControle).Caption
            End If
        End If
        
        Print #1, sNomeControle & " - " & CStr(iIndiceControle) & " : " & sCaptionCtl
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1234
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    Close
    
    Exit Sub
     
Erro_Testa_MenuItensPai:

    Select Case gErr
          
        
        Case Else
            Resume Next 'Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165265)
     
    End Select
     
    Close
    
    Call Comando_Fechar(lComando)
    
    Exit Sub

End Sub

Private Sub Habilita_Itens_Menu2(colUsuarioItensMenu As Collection)

Dim objUsuItensMenu As New ClassUsuarioItensMenu, iIndice As Integer, bRemover As Boolean

    'pesquisa a visibilidade do item de menu
    For Each objUsuItensMenu In colUsuarioItensMenu

        iIndice = iIndice + 1
        bRemover = False
        
        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL_BACKOFFICE Then

            If UCase(objUsuItensMenu.sNomeControle) = "MNULJMOV" Then

                If UCase(objUsuItensMenu.sNomeTela) = "BORDEROCHEQUE" Then
    
                    bRemover = True
    
                End If

            ElseIf UCase(objUsuItensMenu.sNomeControle) = "MNULJROT" Then

                If UCase(objUsuItensMenu.sNomeTela) = "GERACAOARQBACK" Then

                    bRemover = True

                End If

            End If

        ElseIf giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

            If UCase(objUsuItensMenu.sNomeControle) = "MNULJCAD" Then

                If (UCase(objUsuItensMenu.sNomeTela) = "PRODUTO" Or UCase(objUsuItensMenu.sNomeTela) = "VENDEDORES") Then

                    bRemover = True

                End If

            ElseIf UCase(objUsuItensMenu.sNomeControle) = "MNULJCADTA" Then
            
                If UCase(objUsuItensMenu.sNomeTela) = "TABELAPRECOITEM" Then

                    bRemover = True

                End If

            ElseIf UCase(objUsuItensMenu.sNomeControle) = "MNULJCONFIG" Then

                If UCase(objUsuItensMenu.sNomeTela) = "VENDEDORFILIAL" Then
                
                    bRemover = True

                End If

            ElseIf UCase(objUsuItensMenu.sNomeControle) = "MNULJROT" Then

                If UCase(objUsuItensMenu.sNomeTela) = "GERACAOARQBACK" Then

                    bRemover = True

                End If
                
            End If

        ElseIf giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then

            If UCase(objUsuItensMenu.sNomeControle) = "MNULJCAD" Then

                If UCase(objUsuItensMenu.sNomeTela) = "CAIXA" Or UCase(objUsuItensMenu.sNomeTela) = "OPERADOR" Or UCase(objUsuItensMenu.sNomeTela) = "ECF" Then

                    bRemover = True

                End If

            ElseIf UCase(objUsuItensMenu.sNomeControle) = "MNULJCADTA" Then

                If UCase(objUsuItensMenu.sNomeTela) = "PRODUTODESCONTO" Or UCase(objUsuItensMenu.sNomeTela) = "REDE" Or UCase(objUsuItensMenu.sNomeTela) = "TECLADO" Or UCase(objUsuItensMenu.sNomeTela) = "IMPRESSORAECF" Or UCase(objUsuItensMenu.sNomeTela) = "ADMMEIOPAGTO" Then

                    bRemover = True

                End If

            ElseIf UCase(objUsuItensMenu.sNomeControle) = "MNULJROT" Then

                If UCase(objUsuItensMenu.sNomeTela) = "GERACAOARQCC" Then

                    bRemover = True

                End If
                
            ElseIf UCase(objUsuItensMenu.sNomeControle) = "MNULJCONFIG" Then

                If UCase(objUsuItensMenu.sNomeTela) = "LOJACONFIG" Or UCase(objUsuItensMenu.sNomeTela) = "TECLADOPRODUTO" Then
                
                    bRemover = True

                End If
                
            ElseIf UCase(objUsuItensMenu.sNomeControle) = "MNULJMOV" Then

                If UCase(objUsuItensMenu.sNomeTela) = "DEPOSITOBANCARIO" Or UCase(objUsuItensMenu.sNomeTela) = "DEPOSITOCAIXA" Or UCase(objUsuItensMenu.sNomeTela) = "SAQUECAIXA" Or UCase(objUsuItensMenu.sNomeTela) = "TRANSFCENTRAL" Or UCase(objUsuItensMenu.sNomeTela) = "RECEBIMENTOCARNE" Then

                    bRemover = True

                End If
                
            End If

        End If
    
        If bRemover Then
            colUsuarioItensMenu.Remove (iIndice)
            iIndice = iIndice - 1
        End If
        
    Next

End Sub

'#######################################################
'Inserido por Wagner
Private Sub BotaoBrowseCria_Click()
    
    Call Chama_Tela("BrowseCria")

End Sub
'#######################################################

'#######################################################
'Inserido por Wagner 13/07/2006
Private Function Monta_Botoes_Filiais() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colFilialEmpresa As New Collection
Dim objUsuarioEmpresa As ClassUsuarioEmpresa
Dim sCodUsuario As String
Dim lCodEmpresa As Long
Dim sConteudo As String
Dim sSigla As String

On Error GoTo Erro_Monta_Botoes_Filiais

    lErro = CF("CRFATConfig_Le", EXIBE_BOTOES_FILIAIS, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO And lErro <> 61455 Then gError 181290

    If StrParaInt(sConteudo) = MARCADO Then

        lCodEmpresa = glEmpresa
        sCodUsuario = gsUsuario
    
        'Carregar todas as filiais da empresa selecionada para os quais o usuário está autorizado a acessar
        lErro = FiliaisEmpresa_Le_Usuario(sCodUsuario, lCodEmpresa, colFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 50172 Then gError 181189
    
        'Se não houverem filiais para empresa/usuário em questão ==> erro
        If lErro = 50172 Then gError 181190
        
        For iIndice = 1 To BotaoFilial.Count - 1
            Unload BotaoFilial(iIndice)
        Next
    
        BotaoFilial(0).Visible = True
    
        For iIndice = 1 To colFilialEmpresa.Count - 1
        
            'Inclui na tela um novo Controle para essa coluna
            Load BotaoFilial(iIndice)
            
            'Traz o controle recem desenhado para a frente
            BotaoFilial(iIndice).ZOrder
            
            'Torna o controle visível
            BotaoFilial(iIndice).Visible = True
            
        Next
        
        iIndice = -1
        For Each objUsuarioEmpresa In colFilialEmpresa
            
            iIndice = iIndice + 1
            
            BotaoFilial(iIndice).Tag = objUsuarioEmpresa.sNomeFilial
            
            If objUsuarioEmpresa.iCodFilial = 0 Then objUsuarioEmpresa.sNomeFilial = "Empresa Toda"
        
            Call CF("Menu_Obtem_Sigla_FilialEmpresa", objUsuarioEmpresa.iCodFilial, sSigla)
        
            BotaoFilial(iIndice).Caption = sSigla ' objUsuarioEmpresa.iCodFilial
            BotaoFilial(iIndice).ToolTipText = objUsuarioEmpresa.iCodFilial & SEPARADOR & objUsuarioEmpresa.sNomeFilial
            
            If giFilialEmpresa = objUsuarioEmpresa.iCodFilial Then
            
                BotaoFilial(iIndice).BackColor = BOTAO_FILIAL_ATIVA_COR
                BotaoFilial(iIndice).FontBold = True
                
            Else
            
                BotaoFilial(iIndice).BackColor = BOTAO_FILIAL_NAO_ATIVA_COR
                BotaoFilial(iIndice).FontBold = False
                
            End If
            
            If iIndice <> 0 Then
                BotaoFilial(iIndice).left = BotaoFilial(iIndice - 1).left + BotaoFilial(iIndice - 1).Width + 25
            End If
        
        Next
        
    Else
    
        For iIndice = BotaoFilial.LBound To BotaoFilial.UBound
            BotaoFilial(iIndice).Visible = False
        Next
        
    End If
    
    Monta_Botoes_Filiais = SUCESSO
    
    Exit Function

Erro_Monta_Botoes_Filiais:

    Monta_Botoes_Filiais = gErr

    Select Case gErr

        Case 181189, 181290

        Case 181190
            Call Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_SEM_FILIAIS", gErr, sCodUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 181191)

    End Select

    Exit Function
    
End Function

Private Sub BotaoFilial_Click(Index As Integer)

Dim lErro As Long
Dim colMenuItens As New Collection
Dim iFilEmpAnterior As Integer
Dim objFilialEmpresa As New ClassFilialEmpresa
Dim objForm As Form
Dim objMenu As Object
Dim objAux As Object

On Error GoTo Erro_BotaoFilial_Click

    For Each objForm In Forms
        If Not (objForm Is Me) And Not (objForm Is gobjEstInicial) Then gError 181194
    Next
    
    iFilEmpAnterior = giFilialEmpresa
   
    objFilialEmpresa.sNomeEmpresa = gsNomeEmpresa
    objFilialEmpresa.lCodEmpresa = glEmpresa
    objFilialEmpresa.sNomeFilial = BotaoFilial(Index).Tag
    objFilialEmpresa.iCodFilial = Codigo_Extrai(BotaoFilial(Index).ToolTipText)
    
    'se nao mudou empresa nem filial nao precisa fazer nada
    If objFilialEmpresa.lCodEmpresa <> glEmpresa Or objFilialEmpresa.iCodFilial <> giFilialEmpresa Then
    
        lErro = Sistema_Reseta_Modulos
        If lErro <> SUCESSO Then gError 181195
        
        'Configura Empresa e Filial inclusive conexão
        lErro = Empresa_Filial_Configura(objFilialEmpresa)
        If lErro <> SUCESSO Then gError 181196
        
        PrincipalNovo.Caption = TITULO_TELA_PRINCIPAL & " - " & gsNomeEmpresa & " - " & gsNomeFilialEmpresa
    
        Set gcolModulo = New AdmColModulo
        
        'Carrega em gcolModulo módulos indicando atividade p/ FilialEmpresa
        lErro = CF("Modulos_Le_Empresa_Filial", glEmpresa, giFilialEmpresa, gcolModulo)
        If lErro <> SUCESSO Then gError 181197
            
        'Carrega combo com módulos ativos p/ essa Filial e com permissão (alguma tela ou rotina) p/ Usuário
        lErro = Carrega_ComboModulo()
        If lErro <> SUCESSO Then gError 181198

        'se trocou de filial p/EMPRESA_TODA ou vice-versa
'        If (iFilEmpAnterior = EMPRESA_TODA And giFilialEmpresa <> EMPRESA_TODA) Or (iFilEmpAnterior <> EMPRESA_TODA And giFilialEmpresa = EMPRESA_TODA) Then
        If ((iFilEmpAnterior = EMPRESA_TODA Or iFilEmpAnterior = Abs(giFilialAuxiliar)) And (giFilialEmpresa <> EMPRESA_TODA And giFilialEmpresa <> Abs(giFilialAuxiliar))) Or ((iFilEmpAnterior <> EMPRESA_TODA And iFilEmpAnterior <> Abs(giFilialAuxiliar)) And (giFilialEmpresa = EMPRESA_TODA Or giFilialEmpresa = Abs(giFilialAuxiliar))) Then
        
            'vou ter que acertar o menu
            
            Me.Visible = False
            FormAguarde.Show
            DoEvents
            
            'torna todos os itens de menu visiveis novamente
            For Each objMenu In Me.Controls()
            
                If TypeName(objMenu) = "Menu" Then
                    If objMenu.Index = 0 Then
                    
                        objMenu.Enabled = True
                        objMenu.Visible = True
                        
                    Else
                    
                        Set objAux = Me.Controls(objMenu.Name)
                        objAux(objMenu.Index).Enabled = True
                        objAux(objMenu.Index).Visible = True
                        
                    End If
                End If
                
            Next
            
            'esconde os itens de menus à que o usuario nao tem acesso
            lErro = Habilita_Itens_Menu(colMenuItens)
            If lErro <> SUCESSO Then gError 181199
        
            lErro = Habilita_Separadores(colMenuItens)
            If lErro <> SUCESSO Then gError 181200
            
            lErro = Mostra_Menu(NENHUM_MODULO)
            If lErro <> SUCESSO Then gError 181201
    
            Unload FormAguarde
            Me.Visible = True
            
        End If
        
        Call Monta_Botoes_Filiais
        
    End If
    
    Exit Sub
    
Erro_BotaoFilial_Click:

    Select Case gErr

        Case 343
            Resume Next
                
        Case 181194
            Call Rotina_Erro(vbOKOnly, "ERRO_FECHAR_JANELAS_FILHAS", gErr)
            
        Case 181195 To 181201
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181202)

    End Select

    Unload FormAguarde
    Me.Visible = True
    
    Exit Sub
    
End Sub
'#######################################################
Private Sub Impressoras_DesabilitaSuporteBiDirecional()
Dim iDesabilitar As Integer, sAspas As String, sAux As String
Dim prtPrinter As Printer

    iDesabilitar = GetPrivateProfileInt("Geral", "ImpressorasDesabBidi", 0, NOME_ARQUIVO_ADM)
    If iDesabilitar <> 0 Then
    
        sAspas = Chr$(34)
        For Each prtPrinter In Printers
            If InStr(1, prtPrinter.DeviceName, "TOSHIBA e-STUDIO163") <> 0 Or InStr(1, prtPrinter.DeviceName, "hp officejet 5500 series") <> 0 Then
                sAux = "printui.dll,PrintUIEntry /q /Xs /n " & sAspas & prtPrinter.DeviceName
                sAux = sAux & sAspas & " attributes -EnableBidi"
                ShellExecute Me.hWnd, "Open", "rundll32", sAux, vbNullString, vbHide
            End If
        Next
    
    End If
    
End Sub

Private Sub mnuTRVFATMov_Click(Index As Integer)

    Select Case Index

        Case MENU_TRVFAT_MOV_APORTE
            Call Chama_Tela("TRVAporte")

        Case MENU_TRVFAT_MOV_FAT
            Call Chama_Tela("TRVFaturamento")
 
        Case MENU_TRVFAT_MOV_GER_NF
            Call Chama_Tela("TRVGeracaoNF")
 
        Case MENU_TRVFAT_MOV_LIB_APORTE
            Call Chama_Tela("TRVLiberaAporte")
 
        Case MENU_TRVFAT_MOV_LIB_OCR
            Call Chama_Tela("TRVLiberaOcr")
 
        Case MENU_TRVFAT_MOV_OCR
            Call Chama_Tela("TRVOcorrencias")
 
        Case MENU_TRVFAT_MOV_ACORDOS
            Call Chama_Tela("TRVAcordos")
            
        Case MENU_TRVFAT_MOV_CANCFAT
            Call Chama_Tela("TRVCancelarFatura")
 
        Case MENU_TRVFAT_MOV_VOUCOMI
            Call Chama_Tela("TRVVouComi")
 
        Case MENU_TRVFAT_MOV_FAT_CARTAO
            Call Chama_Tela("TRVFatCartao")
            
        Case MENU_TRVFAT_MOV_OCRCASO
            Call Chama_Tela("TRVOcrCasos")
 
        Case MENU_TRVFAT_MOV_OCRCASO_LIBCOBER
            Call Chama_Tela("TRVLibCoberOcrCasos")
 
        Case MENU_TRVFAT_MOV_OCRCASO_LIBJUR
            Call Chama_Tela("TRVLibJurOcrCasos")
            
        Case MENU_TRVFAT_MOV_REEMBOLSO
            Call Chama_Tela("TRVOcrCasosReemb")
 
    End Select

End Sub

Private Sub mnuTRVFATCon_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index

        Case MENU_TRVFAT_CON_APORTE
            Call Chama_Tela("TRVAportesLista")

        Case MENU_TRVFAT_CON_NVL
            
            colSelecao.Add 1 'INATIVACAO_AUTOMATICA_CODIGO
        
            Call Chama_Tela("TRVOcorrenciaLista", colSelecao, Nothing, Nothing, "Origem = ?")

        Case MENU_TRVFAT_CON_OCR_BLOQ
        
            colSelecao.Add 2 'STATUS_TRV_OCR_BLOQUEADO
        
            Call Chama_Tela("TRVOcorrenciaLista", colSelecao, Nothing, Nothing, "Status = ?")

        Case MENU_TRVFAT_CON_OCR_LIB
            
            colSelecao.Add 1 'STATUS_TRV_OCR_LIBERADO
            
            Call Chama_Tela("TRVOcorrenciaLista", colSelecao, Nothing, Nothing, "Status = ?")

        Case MENU_TRVFAT_CON_OCR_TODAS
            Call Chama_Tela("TRVOcorrenciaLista")

        Case MENU_TRVFAT_CON_PAGTO_NAO_FAT
            Call Chama_Tela("AportesPagtosLista", colSelecao, Nothing, Nothing, "StatusCod = 1 AND TipoDocDestino = 0 AND NumIntDocDestino = 0 AND FormaPagto = 1 AND TipoCod IN (1,2) ")

        Case MENU_TRVFAT_CON_PAGTO_TODOS
            Call Chama_Tela("AportesPagtosLista")

        Case MENU_TRVFAT_CON_APORTE_SF
            Call Chama_Tela("TRVAPortesPagtoFatLista")

        Case MENU_TRVFAT_CON_HIST_APORTE_SF
            Call Chama_Tela("TRVAPortesPagtoFatHistLista")

        Case MENU_TRVFAT_CON_VOU_NAO_FAT
            
            colSelecao.Add "VOU"

            Call Chama_Tela("DocsParaFatLista", colSelecao, Nothing, Nothing, "TipoDoc = ?")

        Case MENU_TRVFAT_CON_VOU_PREV_REEMB
            Call Chama_Tela("VouchersCancNaoReembolsadosLista")

        Case MENU_TRVFAT_CON_VOU_TODOS
        
            Call Chama_Tela("VoucherRapidoLista", colSelecao, Nothing, Nothing)
            
        Case MENU_TRVFAT_CON_TITREC_SEM_NF
            Call Chama_Tela("TitulosSemNotaLista")
            
        Case MENU_TRVFAT_CON_DOCS_A_FAT
            Call Chama_Tela("DocsParaFatLista")

        Case MENU_TRVFAT_CON_DOCS_FAT
            Call Chama_Tela("DocFaturadosLista")
            
        Case MENU_TRVFAT_CON_VOU_REP
            Call Chama_Tela("VouchersPorReprLista")
            
        Case MENU_TRVFAT_CON_VOU_COR
            Call Chama_Tela("VouchersPorCorrLista")
            
        Case MENU_TRVFAT_CON_VOU_EMI
            Call Chama_Tela("VouchersPorEmisLista")
            
        Case MENU_TRVFAT_CON_VOU_TELA
            Call Chama_Tela("TRVVoucher")

        Case MENU_TRVFAT_CON_FAT_CANC
            Call Chama_Tela("TRVFaturasCancLista")

        Case MENU_TRVFAT_CON_FAT
            Call Chama_Tela("TRVConsultaFatura")
            
        Case MENU_TRVFAT_CON_UTIL_CRED
            Call Chama_Tela("TRVCreditosUtilizadosLista")

        Case MENU_TRVFAT_CON_CRED_DISP
            Call Chama_Tela("TRVAporteCredCliLista")

        Case MENU_TRVFAT_CON_VEND_CALLCENTER
            Call Chama_Tela("TRVVendasCallCenterLista")

        Case MENU_TRVFAT_CON_POS_GERAL
            Call Chama_Tela("TRVVouComplLista")

        Case MENU_TRVFAT_CON_EST_VENDA
            Call Chama_Tela("TRVEstVenda")

        Case MENU_TRVFAT_CON_FAT_CRT
            Call Chama_Tela("TRVTitulosCartaoLista")
            
        Case MENU_TRVFAT_CON_VOU_BAIXAS
            Call Chama_Tela("TRVVouBaixaLista")
            
    End Select

End Sub

Private Sub mnuTRVFATCon1_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
    
        Case MENU_TRVFAT_CON1_OCR_TODAS_ENV
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "")

        Case MENU_TRVFAT_CON1_OCR_TODAS_ABERTAS
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "DataAbertura <> {d '1822-09-07'}")
               
        Case MENU_TRVFAT_CON1_OCR_AGUARD_DOCS
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "DataAbertura <> {d '1822-09-07'} AND DataDocsRec = {d '1822-09-07'}")
        
        Case MENU_TRVFAT_CON1_OCR_TODAS_AUTORIZADAS
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "CGAutorizadoPor <> 0")
        
        Case MENU_TRVFAT_CON1_OCR_AUTO_NAO_FAT
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "CGAutorizadoPor <> 0 AND NumIntDocTitPagCobertura = 0")
        
        Case MENU_TRVFAT_CON1_OCR_TODAS_FAT_COBR
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "NumIntDocTitPagCobertura <> 0")
        
        Case MENU_TRVFAT_CON1_OCR_FAT_COBR_NAO_PAGAS
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "NumIntDocTitPagCobertura <> 0 AND DataPagtoPax = {d '1822-09-07'}")
        
        Case MENU_TRVFAT_CON1_OCR_TODAS_PROCESSO
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "Judicial = 'SIM'")
        
        Case MENU_TRVFAT_CON1_OCR_PROCESSO_ABERTO
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "Judicial = 'SIM' AND DataFimProcesso = {d '1822-09-07'}")
        
        Case MENU_TRVFAT_CON1_OCR_TODAS_CONDENADAS
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "Condenado = 'SIM'")
        
        Case MENU_TRVFAT_CON1_OCR_CONDENADAS_NAO_FAT
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "Condenado = 'SIM' AND NumIntDocTitPagProcesso = 0")
        
        Case MENU_TRVFAT_CON1_OCR_TODAS_FAT_PROC
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "NumIntDocTitPagProcesso <> 0")
        
        Case MENU_TRVFAT_CON1_OCR_FAT_PROC_NAO_PAGAS
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "NumIntDocTitPagProcesso <> 0 AND DataPagtoCond = {d '1822-09-07'}")
        
        Case MENU_TRVFAT_CON1_OCR_TODAS_REEMB
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "ValorPreReceber > 0")
        
        Case MENU_TRVFAT_CON1_OCR_REEMB_NAO_REC
            Call Chama_Tela("TRVAssistenciaLista", colSelecao, Nothing, Nothing, "ValorPreReceber > 0 AND NumIntDocTitRecReembolso = 0")
        
        Case MENU_TRVFAT_CON1_TP_COBERTURA
            Call Chama_Tela("TitPagTodosTFLista", colSelecao, Nothing, Nothing, "SiglaDocumento = 'OCRC'")
        
        Case MENU_TRVFAT_CON1_TP_JUDICIAL
            Call Chama_Tela("TitPagTodosTFLista", colSelecao, Nothing, Nothing, "SiglaDocumento = 'OCRJ'")
        
        Case MENU_TRVFAT_CON1_TR_REEMBOLSO
            Call Chama_Tela("TitRecTodosTFLista", colSelecao, Nothing, Nothing, "SiglaDocumento = 'OCRR'")
        
        Case MENU_TRVFAT_CON1_VLR_AUTO
            Call Chama_Tela("TRVAssistVlrAutoLista", colSelecao, Nothing, Nothing)
        
    End Select
End Sub

Private Sub mnuTRVFATCon2_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index

        Case MENU_TRVFAT_CON2_ACORDOS
            Call Chama_Tela("TRVAcordosLista")

        Case MENU_TRVFAT_CON2_CMC
            
            colSelecao.Add "CMC"
        
            Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ?")

        Case MENU_TRVFAT_CON2_CMC_BLOQ
            
            colSelecao.Add "CMC"
            colSelecao.Add 2
        
            Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND StatusCod = ? ")

        Case MENU_TRVFAT_CON2_CMC_LIB
            
            colSelecao.Add "CMC"
            colSelecao.Add 1
            colSelecao.Add 0
        
            Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND StatusCod = ? AND NumIntDocDestino = ?")

        Case MENU_TRVFAT_CON2_CMA
            
            colSelecao.Add "CMA"
        
            Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ?")

        Case MENU_TRVFAT_CON2_CMCC
            
            colSelecao.Add "CMCC"
        
            Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ?")

        Case MENU_TRVFAT_CON2_CMCC_BLOQ
            
            colSelecao.Add "CMCC"
            colSelecao.Add 2
        
            Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND StatusCod = ? ")

        Case MENU_TRVFAT_CON2_CMCC_LIB
            
            colSelecao.Add "CMCC"
            colSelecao.Add 1
            colSelecao.Add 0
        
            Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND StatusCod = ? AND NumIntDocDestino = ?")

        Case MENU_TRVFAT_CON2_CMR
        
            colSelecao.Add "CMR"
        
            Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ?")

        Case MENU_TRVFAT_CON2_CMR_BLOQ
        
            colSelecao.Add "CMR"
            colSelecao.Add 2
        
            Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND StatusCod = ?")

        Case MENU_TRVFAT_CON2_CMR_LIB
        
            colSelecao.Add "CMR"
            colSelecao.Add 1
            colSelecao.Add 0
        
            Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND StatusCod = ? AND NumIntDocDestino = ?")

        Case MENU_TRVFAT_CON2_OVER
        
            colSelecao.Add "OVER"
        
            Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ?")

        Case MENU_TRVFAT_CON2_OVER_BLOQ
        
            colSelecao.Add "OVER"
            colSelecao.Add 2
        
            Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND StatusCod = ?")

        Case MENU_TRVFAT_CON2_OVER_LIB
        
            colSelecao.Add "OVER"
            colSelecao.Add 1
            colSelecao.Add 0
        
            Call Chama_Tela("VoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND StatusCod = ? AND NumIntDocDestino = ?")

        Case MENU_TRVFAT_CON2_COMI_CC
            Call Chama_Tela("NFPagEmpTodaLista", colSelecao, Nothing, Nothing, "NumIntTitPag = 0")

        Case MENU_TRVFAT_CON2_OVER_FAT
            Call Chama_Tela("TRVOverLista", colSelecao, Nothing, Nothing, "NumTitulo <> 0")
        
        Case MENU_TRVFAT_CON2_OVER_IH
            colSelecao.Add gdtDataAtual
            Call Chama_Tela("TRVOverIHLista", colSelecao, Nothing, Nothing, "DataEmissao = ?")

    End Select

End Sub

Private Sub mnuTRVFATConfig_Click(Index As Integer)

    Select Case Index

        Case MENU_TRVFAT_CONFIG_CONFIG
            Call Chama_Tela("TRVConfig")

    End Select

End Sub

Private Sub mnuTRVFATRot_Click(Index As Integer)

    Select Case Index

        Case MENU_TRVFAT_ROT_EXTRAIDADOS
            Call Chama_Tela("TRVExtraiDadosSigav")
            
        Case MENU_TRVFAT_ROT_REG_FAT
            Call Chama_Tela("TRVRegerarFaturas")
            
        Case MENU_TRVFAT_ROT_GER_COMI_INT
            Call Chama_Tela("TRVGerComiInt")
    
        Case MENU_TRVFAT_ROT_GER_RELS_EXCEL
            Call Chama_Tela("TRVRelsEmExcel")
    
    End Select

End Sub

Private Sub MnuFATCadTRV_Click(Index As Integer)

    Select Case Index
        
        Case MENU_TRVFAT_CAD_TIPOOCR
            Call Chama_Tela("TRVTiposOcorrencia")

    End Select


End Sub

Private Sub mnuFATConRPS_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index

        Case MENU_FAT_CON_RPS_ARQ
        
            Call Chama_Tela("RPSCabLista", colSelecao, Nothing, Nothing)

        Case MENU_FAT_CON_RPS_HIST
            
            Call Chama_Tela("RPSEnviadosLista", colSelecao, Nothing, Nothing)


        Case MENU_FAT_CON_RPS_NAOENV

            colSelecao.Add 0
            
            Call Chama_Tela("RPSLista", colSelecao, Nothing, Nothing, "Enviado = ?")


        Case MENU_FAT_CON_RPS_NAONFE
            
            Call Chama_Tela("RPSLista", colSelecao, Nothing, Nothing, "NOT EXISTS (SELECT N.NumIntDoc FROM NFe AS N WHERE N.SerieRPS = RPS.Serie AND N.DataEmissaoRPS = RPS.DataEmissao AND N.NumeroRPS = RPS.Numero AND N.FilialEmpresa = RPS.FilialEmpresa )")


        Case MENU_FAT_CON_RPS_TODOS
            
            Call Chama_Tela("RPSLista", colSelecao, Nothing, Nothing)


    End Select

End Sub

Private Sub mnuFATConNFE_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index

        Case MENU_FAT_CON_NFE_TODAS
            
            Call Chama_Tela("NFeLista", colSelecao, Nothing, Nothing)

        Case MENU_FAT_CON_NFE_HIST
            
            Call Chama_Tela("NFeRecebidasLista", colSelecao, Nothing, Nothing)

    End Select

End Sub

Private Sub mnuSRVConCad_Click(Index As Integer)
               
Dim colSelecao As New Collection

    Select Case Index
        
        Case MENU_SRV_CON_CAD_PRODUTOS
            Call Chama_Tela("ProdutoLista_Consulta", colSelecao)

        Case MENU_SRV_CON_CAD_GARANTIA
            Call Chama_Tela("GarantiaLista", colSelecao)
        
        Case MENU_SRV_CON_CAD_TIPOGARANTIA
            Call Chama_Tela("TipoGarantiaLista", colSelecao)
        
'        Case MENU_SRV_CON_CAD_CONTRATOSRV
'            Call Chama_Tela("ContratoSrvLista")
'
'        Case MENU_SRV_CON_CAD_ITENSCONTRATOSRV
'            Call Chama_Tela("ItensContratoSrvLista")

        Case MENU_SRV_CON_CAD_CONTRATOSRV
            Call Chama_Tela("ItensDeContratoSrvLista")

        Case MENU_SRV_CON_CAD_MAODEOBRA
            Call Chama_Tela("MaoDeObraLista", colSelecao)

        Case MENU_SRV_CON_CAD_MAQUINAS
            Call Chama_Tela("MaquinasLista", colSelecao)

        Case MENU_SRV_CON_CAD_COMPETENCIA
            Call Chama_Tela("CompetenciasLista", colSelecao)

        Case MENU_SRV_CON_CAD_CT
            Call Chama_Tela("CentrodeTrabalhoLista", colSelecao)

    End Select


End Sub

Private Sub mnuSRVConSolic_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
    
        Case MENU_SRV_CON_SOLIC_ABERTA
            Call Chama_Tela("SolicSRVAbertaLista", colSelecao)
    
        Case MENU_SRV_CON_SOLIC_BAIXADA
            Call Chama_Tela("SolicSRVBaixadaLista", colSelecao)
    
        Case MENU_SRV_CON_SOLIC_TODAS
            Call Chama_Tela("SolicitacaoSRVLista", colSelecao)

        Case MENU_SRV_CON_CRM
            Call Chama_Tela("RelacCliSolSRVLista")

        Case MENU_SRV_CON_BI_SS
            Call Chama_Tela("SSRVLista")

    End Select
    
End Sub

Private Sub mnuSRVCon_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
    
        Case MENU_SRV_CON_ORCAMENTOS
            Call Chama_Tela("OrcamentoSRV1Lista", colSelecao)
    
        Case MENU_SRV_CON_ITENSPEDIDOS
            Call Chama_Tela("ItensPedidoSRV1Lista", colSelecao)
            
        Case MENU_SRV_CON_ITEMOS
            Call Chama_Tela("ItemOSLista", colSelecao)
            
        Case MENU_SRV_CON_NFSRV
            Call Chama_Tela("NFiscalSRV1Lista", colSelecao)
            
        Case MENU_SRV_CON_ITEMNFSRV
            Call Chama_Tela("ItemNFSRVLista", colSelecao)
            
    End Select

End Sub

Private Sub mnuSRVMov_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
    
        Case MENU_SRV_MOV_SOLICSRV
            Call Chama_Tela("SolicitacaoSRV")
    
        Case MENU_SRV_MOV_ORCAMENTOSRV
            Call Chama_Tela("OrcamentoSRV")
    
        Case MENU_SRV_MOV_PEDIDOSRV
            Call Chama_Tela("PedidoServico")
            
        Case MENU_SRV_MOV_OS
            Call Chama_Tela("OrdemServico")
    
        Case MENU_SRV_MOV_NFSRV
            Call Chama_Tela("NFiscalSRV")
    
        Case MENU_SRV_MOV_NFFATGARANTIASRV
            Call Chama_Tela("NFiscalFatGarSRV")
    
        Case MENU_SRV_MOV_NFFATPEDIDOSRV
            Call Chama_Tela("NFiscalFatPedSRV")
    
        Case MENU_SRV_MOV_NFFATSRV
            Call Chama_Tela("NFiscalFatSRV")
    
        Case MENU_SRV_MOV_NFPEDIDOSRV
            Call Chama_Tela("NFiscalPedSRV")
    
'        Case MENU_SRV_MOV_ACOMPANHAMENTOSRV
'            Call Chama_Tela("AcompanhamentoSRV")
            
        Case MENU_SRV_MOV_LIBORCSRV
            Call Chama_Tela_Nova_Instancia("LiberaBloqueioGen", 2)
            
        Case MENU_SRV_MOV_LIBPEDSRV
            Call Chama_Tela_Nova_Instancia("LiberaBloqueioGen", 1)
            
        Case MENU_SRV_MOV_OSAPONT
            Call Chama_Tela("OSApontamento")
            
        Case MENU_SRV_MOV_MOVEST
            Call Chama_Tela("MovEstoqueSRV")
            
        Case MENU_SRV_MOV_BAIXAPEDSRV
            Call Chama_Tela("BaixaPedidoSRV")
    End Select

End Sub

Private Sub mnuSRVCad_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
    
        Case MENU_SRV_CAD_PRODUTO
            Call Chama_Tela("Produto")
    
        Case MENU_SRV_CAD_CLIENTE
            Call Chama_Tela("Clientes")
    
        Case MENU_SRV_CAD_GARANTIA
            Call Chama_Tela("Garantia")
    
        Case MENU_SRV_CAD_TIPOGARANTIA
            Call Chama_Tela("TipoGarantia")
            
        Case MENU_SRV_CAD_CONTRATO
            Call Chama_Tela("ContratoCadastro")
            
        Case MENU_SRV_CAD_CONTRATOSRV
            Call Chama_Tela("ContratoSRV")
    
        Case MENU_SRV_CAD_MAODEOBRA
            Call Chama_Tela("MaoDeObra")
    
        Case MENU_SRV_CAD_MAQUINA
            Call Chama_Tela("Maquinas")
    
        Case MENU_SRV_CAD_TIPOMAODEOBRA
            Call Chama_Tela("TiposDeMaoDeObra")
            
        Case MENU_SRV_CAD_CENTROTRABALHO
            Call Chama_Tela("CentrodeTrabalho")
    
        Case MENU_SRV_CAD_COMPETENCIA
            Call Chama_Tela("Competencias")
    
        Case MENU_SRV_CAD_ROTEIRO
            Call Chama_Tela("RoteiroSRV")
            
    End Select

End Sub

Private Sub mnuSRVConOS_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
    
        Case MENU_SRV_CON_OS_ABERTA
            Call Chama_Tela("OSAbertaLista", colSelecao)
    
        Case MENU_SRV_CON_OS_BAIXADA
            Call Chama_Tela("OSBaixadaLista", colSelecao)
    
        Case MENU_SRV_CON_OS_TODAS
            Call Chama_Tela("OSLista", colSelecao)

    End Select

End Sub

Private Sub mnuSRVConPed_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
    
        Case MENU_SRV_CON_PED_ABERTOS
            Call Chama_Tela("PedSRVAbertosLista", colSelecao)
    
        Case MENU_SRV_CON_PED_BAIXADOS
            Call Chama_Tela("PedSRVBaixadosLista", colSelecao)
    
        Case MENU_SRV_CON_PED_TODOS
            Call Chama_Tela("PedidoServico_Lista", colSelecao)

        Case MENU_SRV_CON_BI_PS
            Call Chama_Tela("PSRVLista", colSelecao)

    End Select

End Sub

Private Sub mnuESTConTRV_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
            
        Case MENU_EST_CON_TRV_MOVEST
            Call Chama_Tela("MovEstTRVLista", colSelecao)

        Case MENU_EST_CON_TRV_MOVEST_TRASNFER
            Call Chama_Tela("MovEstTransfTRVLista", colSelecao)

    End Select

End Sub

Private Sub mnuFATMovAporte_Click(Index As Integer)

    Select Case Index
        
        Case MENU_FAT_MOV_APORTE_EMISSAO
            Call Chama_Tela("TRPAporte")
        
        Case MENU_FAT_MOV_APORTE_LIBERACAO
            Call Chama_Tela("TRPLiberaAporte")
        
    End Select
    
End Sub

Private Sub mnuFATMovVou_Click(Index As Integer)

    Select Case Index
        
        Case MENU_FAT_MOV_VOU_EMISSAO
            Call Chama_Tela("TRPVouEmi")
        
        Case MENU_FAT_MOV_VOU_MANUTENCAO
            Call Chama_Tela("TRPVouManu")
            
        Case MENU_FAT_MOV_VOU_COMISSAO
            Call Chama_Tela("TRPVouComi")
        
        
    End Select
    
End Sub

Private Sub mnuFATMovOcr_Click(Index As Integer)

    Select Case Index
        
        Case MENU_FAT_MOV_OCR_EMISSAO
            Call Chama_Tela("TRPOcorrencias")
        
        Case MENU_FAT_MOV_OCR_LIBERACAO
            Call Chama_Tela("TRPLiberaOcr")
        
    End Select
    
End Sub

Private Sub mnuFATMovFat_Click(Index As Integer)

    Select Case Index
        
        Case MENU_FAT_MOV_FAT_FATURAMENTO
            Call Chama_Tela("TRPFaturamento")
        
        Case MENU_FAT_MOV_FAT_CANCELAMENTO
            Call Chama_Tela("TRPCancelarFatura")
    
        Case MENU_FAT_MOV_FAT_FATCARTAO
            Call Chama_Tela("TRPFatCartao")
    
    End Select
    
End Sub

Private Sub mnuFATRotComis_Click(Index As Integer)

    Select Case Index
        
        Case MENU_FAT_ROT_COMIS_EXT
            Call Chama_Tela("TRPGerComiExt")
        
        Case MENU_FAT_ROT_COMIS_INT
            Call Chama_Tela("TRPGerComiInt")
        
    End Select
    
End Sub

Private Sub mnuFATConAporte_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
        
        Case MENU_FAT_CON_APORTE_TODOS
            Call Chama_Tela("TRPAportesPagtosLista")
        
        Case MENU_FAT_CON_APORTE_NAOFAT
            Call Chama_Tela("TRPAportesPagtosLista", colSelecao, Nothing, Nothing, "StatusCod = 1 AND TipoDocDestino = 0 AND NumIntDocDestino = 0 AND FormaPagto = 1 AND TipoCod IN (1,2) ")
        
        Case MENU_FAT_CON_APORTE_SF
            Call Chama_Tela("TRPAPortesPagtoFatLista")
            
        Case MENU_FAT_CON_APORTE_CREDITO
            Call Chama_Tela("TRPAporteCredCliLista")

        Case MENU_FAT_CON_APORTE_HIST_SF
            Call Chama_Tela("TRPAPortesPagtoFatHistLista")
            
        Case MENU_FAT_CON_APORTE_HIST_CRED
            Call Chama_Tela("TRPCreditosUtilizadosLista")
            
    End Select
    
End Sub


Private Sub mnuFATConVou_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
        
        Case MENU_FAT_CON_VOU_TELA
            Call Chama_Tela("TRPVoucher")
        
        Case MENU_FAT_CON_VOU_TODOS
            Call Chama_Tela("TRPVoucherRapidoLista", colSelecao, Nothing, Nothing)
        
        Case MENU_FAT_CON_VOU_NAOFAT
            colSelecao.Add 0
            Call Chama_Tela("TRPVoucherRapidoLista", colSelecao, Nothing, Nothing, "Fatura = ? AND Cancelado = 'Não'")
        
        Case MENU_FAT_CON_VOU_CANCPREV
            Call Chama_Tela("TRPVouchersCancNaoReembolsadosLista")
            
        Case MENU_FAT_CON_VOU_PAX
            Call Chama_Tela("TRPPassageirosLista")

        Case MENU_FAT_CON_VOU_PG
            Call Chama_Tela("TRPVouCompletoLista")

        Case MENU_FAT_CON_VOU_CCSEMAUTO
            Call Chama_Tela("TRPVouCartaoSemAutoLista")
            
    End Select
    
End Sub

Private Sub mnuFATConComis_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
        
        Case MENU_FAT_CON_COMIS_CMRTODAS
            colSelecao.Add "CMR"
            Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ?")
        Case MENU_FAT_CON_COMIS_CMRBLOQ
            colSelecao.Add "CMR"
            colSelecao.Add 2
            Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND Status = ?")
        Case MENU_FAT_CON_COMIS_CMRLIB
            colSelecao.Add "CMR"
            colSelecao.Add 1
            colSelecao.Add 0
            Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND Status = ? AND NumIntDocDestino = ?")
        Case MENU_FAT_CON_COMIS_CMCTODAS
            colSelecao.Add "CMC"
            Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ?")
        Case MENU_FAT_CON_COMIS_CMCBLOQ
            colSelecao.Add "CMC"
            colSelecao.Add 2
            Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND Status = ? ")
        Case MENU_FAT_CON_COMIS_CMCLIB
            colSelecao.Add "CMC"
            colSelecao.Add 1
            colSelecao.Add 0
            Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND Status = ? AND NumIntDocDestino = ?")
        Case MENU_FAT_CON_COMIS_CMCCTODAS
            colSelecao.Add "CMCC"
            Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ?")
        Case MENU_FAT_CON_COMIS_CMCCBLOQ
            colSelecao.Add "CMCC"
            colSelecao.Add 2
            Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND Status = ? ")
        Case MENU_FAT_CON_COMIS_CMCCLIB
            colSelecao.Add "CMCC"
            colSelecao.Add 1
            colSelecao.Add 0
            Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND Status = ? AND NumIntDocDestino = ?")
        Case MENU_FAT_CON_COMIS_CMETODAS
            colSelecao.Add "CME"
            Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ?")
        Case MENU_FAT_CON_COMIS_CMEBLOQ
            colSelecao.Add "CME"
            colSelecao.Add 2
            Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND Status = ?")
        Case MENU_FAT_CON_COMIS_CMELIB
            colSelecao.Add "CME"
            colSelecao.Add 1
            colSelecao.Add 0
            Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ? AND Status = ? AND NumIntDocDestino = ?")
        Case MENU_FAT_CON_COMIS_CMA
            colSelecao.Add "CMA"
            Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc = ?")
        Case MENU_FAT_CON_COMIS_COMISSEMNF
            Call Chama_Tela("TRPNFPagEmpTodaLista", colSelecao, Nothing, Nothing, "NumIntTitPag = 0")
        Case MENU_FAT_CON_COMIS_TODAS
            colSelecao.Add "BRUTO"
            colSelecao.Add "OCR"
            colSelecao.Add "NVL"
            Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "TipoDoc NOT IN (?,?,?)")
    End Select
        
    
End Sub

Private Sub mnuFATConOCR_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
        
        Case MENU_FAT_CON_OCR_TODAS
            Call Chama_Tela("TRPOcorrenciaLista")
        
        Case MENU_FAT_CON_OCR_BLOQUEADAS
            colSelecao.Add 2 'STATUS_TRV_OCR_BLOQUEADO
            Call Chama_Tela("TRPOcorrenciaLista", colSelecao, Nothing, Nothing, "Status = ?")

        Case MENU_FAT_CON_OCR_LIBERADAS
            colSelecao.Add 1 'STATUS_TRV_OCR_LIBERADO
            Call Chama_Tela("TRPOcorrenciaLista", colSelecao, Nothing, Nothing, "Status = ?")
        
        Case MENU_FAT_CON_OCR_INATIVACAO
            colSelecao.Add 1 'INATIVACAO_AUTOMATICA_CODIGO
            Call Chama_Tela("TRPOcorrenciaLista", colSelecao, Nothing, Nothing, "Origem = ?")
        
    End Select
    
End Sub

Private Sub mnuFATConFat_Click(Index As Integer)

Dim colSelecao As New Collection

    Select Case Index
        
        Case MENU_FAT_CON_FAT_CONSULTAFATURA
            Call Chama_Tela("TRPConsultaFatura")
        
        Case MENU_FAT_CON_FAT_CANCELAFATURA
            Call Chama_Tela("TRPFaturasCancLista")
        
        Case MENU_FAT_CON_FAT_TITRECSEMNF
            Call Chama_Tela("TRPTitulosSemNotaLista")
            
        Case MENU_FAT_CON_FAT_DOCPARAFAT
            Call Chama_Tela("TRPDocsParaFatLista")

        Case MENU_FAT_CON_FAT_DOCFATURADO
            Call Chama_Tela("TRPDocFaturadosLista")
        
    End Select
    
End Sub

Private Sub mnuFATConNFeFed_Click(Index As Integer)
    
    Select Case Index
    
        Case MENU_FAT_CON_NFEFED_LOTE
            Call Chama_Tela("NFeFedLoteViewLista")
            
        Case MENU_FAT_CON_NFEFED_LOTELOG
            Call Chama_Tela("NFeFedLoteLogViewLista")
            
        Case MENU_FAT_CON_NFEFED_RETENVI
            Call Chama_Tela("NFeFedRetEnviLista")
            
        Case MENU_FAT_CON_NFEFED_STATUSNF
            Call Chama_Tela("NFeFedProtNFeViewLista")
            
        Case MENU_FAT_CON_NFEFED_RETCONSLOTE
            Call Chama_Tela("NFeFedRetConsReciLista")
            
        Case MENU_FAT_CON_NFEFED_RETCANC
            Chama_Tela ("NFeFedRetCancNFeViewLista")
            
        Case MENU_FAT_CON_NFEFED_RETINUTFAIXA
            Chama_Tela ("NFeFedRetInutNFeLista")
            
    End Select

End Sub

Private Sub mnuFATRotNFe_Click(Index As Integer)

Dim objRelatorio As New AdmRelatorio

    Select Case Index
    
        Case MENU_FAT_ROT_NFEFED_GERALOTEENVIO
            Call Chama_Tela("NFe")
        
        Case MENU_FAT_ROT_NFEFED_CONSULTALOTE
            Call Chama_Tela("ConsultaLoteNFe")
            
        Case MENU_FAT_ROT_NFEFED_EMAIL
            Call Chama_Tela("NFePorEmail")
            
        Case MENU_FAT_ROT_NFEFED_INUTFAIXA
            Call Chama_Tela("NFeInutFaixa")
            
        Case MENU_FAT_ROT_NFEFED_CONSULTANFE
            Call Chama_Tela("ConsultaNFe")
            
        Case MENU_FAT_ROT_NFEFED_NFESCAN
            Call Chama_Tela("NFeFedScan")

        Case MENU_FAT_ROT_NFEFED_EXPORTXMLNFE
            Call objRelatorio.Rel_Menu_Executar("Exportar Xml NFe")
            
        Case MENU_FAT_ROT_NFEFED_CARTACORRECAO
            Call Chama_Tela("CartaCorrecaoNFe")
    
    End Select
        

End Sub

Private Sub mnuFATRotNFSE_Click(Index As Integer)

    Select Case Index
    
        Case MENU_FAT_ROT_NFSE_GERALOTEENVIO
            Call Chama_Tela("NFSE")
        
        Case MENU_FAT_ROT_NFSE_CONSULTALOTE
            Call Chama_Tela("ConsultaLoteNFSE")
            
    End Select

End Sub

Private Sub mnuFATConNFSE_Click(Index As Integer)
    
    Select Case Index
    
        Case MENU_FAT_CON_NFSE_LOTE
            Call Chama_Tela("RPSWEBLoteViewLista")
            
        Case MENU_FAT_CON_NFSE_LOTELOG
            Call Chama_Tela("RPSWEBLoteLogViewLista")
            
        Case MENU_FAT_CON_NFSE_RETENVI
            Call Chama_Tela("RPSWEBRetEnviLista")
            
        Case MENU_FAT_CON_NFSE_AUTORIZADAS
            Call Chama_Tela("RPSWEBProtViewLista")
            
        Case MENU_FAT_CON_NFSE_RETCONSLOTE
            Call Chama_Tela("RPSWEBConsLoteLista")
            
        Case MENU_FAT_CON_NFSE_RETCANC
            Chama_Tela ("RPSWEBRetCancViewLista")
            
        Case MENU_FAT_CON_NFSE_SITLOTE
            Chama_Tela ("RPSWEBConsSitLoteLista")
            
    End Select

End Sub

Private Sub mnuFATConBenef_Click(Index As Integer)

    Select Case Index
    
        Case MENU_FAT_CON_BENEF_SALDO
            Call Chama_Tela("NFRemBenefSaldoLista")
            
        Case MENU_FAT_CON_BENEF_DEV
            Call Chama_Tela("NFRemBenefDevLista")
            
    End Select
    
End Sub

Private Function Habilita_Itens_Menu_VPN() As Long

Dim lErro As Long, lComando As Long, objObjeto As Object
Dim sNomeControle As String, iIndiceControle As Integer

On Error GoTo Erro_Habilita_Itens_Menu_VPN

    lComando = Comando_AbrirExt(GL_lConexaoDicBrowse)
    If lComando = 0 Then gError 201478
    
    sNomeControle = String(255, 0)
    lErro = Comando_Executar(lComando, "SELECT NomeControle, IndiceControle FROM VPNItensMenuEsconder ORDER BY NomeControle, IndiceControle", sNomeControle, iIndiceControle)
    If lErro <> AD_SQL_SUCESSO Then gError 201479
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201480
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objObjeto = Me.Controls(sNomeControle)
        If iIndiceControle <> 0 Then
            objObjeto(iIndiceControle).Visible = False
        Else
            objObjeto.Visible = False
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201481
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    Habilita_Itens_Menu_VPN = SUCESSO
    
    Exit Function
    
Erro_Habilita_Itens_Menu_VPN:

    Habilita_Itens_Menu_VPN = gErr

    Select Case gErr

        Case 387
            Resume Next

        Case 201478

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201477)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Obter_Pgm_Padrao_Email() As String

Dim lngResult As Long
Dim lngKeyHandle As Long
Dim lngValueType As Long
Dim lngBufferSize As Long
Dim strBuffer As String
Dim sNomePgm As String
Dim sDirPgm As String

On Error GoTo Erro_Obter_Pgm_Padrao_Email:
    
    'Obtem o nome do programa padrão de email do usuário
    lngResult = RegOpenKey(HKEY_CURRENT_USER, "SOFTWARE\Clients\Mail", lngKeyHandle)
    lngResult = RegQueryValueEx(lngKeyHandle, "", 0, lngValueType, ByVal 0, lngBufferSize)
    
    If lngValueType = REG_SZ Then
        strBuffer = String(lngBufferSize, " ")
        lngResult = RegQueryValueEx(lngKeyHandle, "", 0, 0, ByVal strBuffer, lngBufferSize)
        If lngResult = ERROR_SUCCESS Then
            If (InStr(strBuffer, Chr$(0))) > 0 Then
                sNomePgm = left(strBuffer, (InStr(strBuffer, Chr$(0))) - 1)
            Else
                sNomePgm = strBuffer
            End If
        End If
    End If
    
    lngResult = RegCloseKey(lngKeyHandle)
    
    'Se o usuário não tem um padrão tenta obter o padrão da máquina
    If Len(Trim(sNomePgm)) = 0 Then
    
        lngResult = 0
        lngValueType = 0
        lngKeyHandle = 0
        lngBufferSize = 0
        strBuffer = ""
        
        lngResult = RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Clients\Mail", lngKeyHandle)
        lngResult = RegQueryValueEx(lngKeyHandle, "", 0, lngValueType, ByVal 0, lngBufferSize)
        
        If lngValueType = REG_SZ Then
            strBuffer = String(lngBufferSize, " ")
            lngResult = RegQueryValueEx(lngKeyHandle, "", 0, 0, ByVal strBuffer, lngBufferSize)
            If lngResult = ERROR_SUCCESS Then
                If (InStr(strBuffer, Chr$(0))) > 0 Then
                    sNomePgm = left(strBuffer, (InStr(strBuffer, Chr$(0))) - 1)
                Else
                    sNomePgm = strBuffer
                End If
            End If
        End If
        
        lngResult = RegCloseKey(lngKeyHandle)
        
    End If
    
    'Com o nome do programa padrão pega a localização
    If Len(Trim(sNomePgm)) > 0 Then
    
        lngResult = 0
        lngValueType = 0
        lngKeyHandle = 0
        lngBufferSize = 0
        strBuffer = ""
        
        lngResult = RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Clients\Mail\" & sNomePgm & "\shell\open\command", lngKeyHandle)
        lngResult = RegQueryValueEx(lngKeyHandle, "", 0, lngValueType, ByVal 0, lngBufferSize)
        
        If lngResult = ERROR_SUCCESS Then
            strBuffer = String(lngBufferSize, " ")
            lngResult = RegQueryValueEx(lngKeyHandle, "", 0, 0, ByVal strBuffer, lngBufferSize)
            If lngResult = ERROR_SUCCESS Then
                If (InStr(strBuffer, Chr$(0))) > 0 Then
                    sDirPgm = left(strBuffer, (InStr(strBuffer, Chr$(0))) - 1)
                Else
                    sDirPgm = strBuffer
                End If
            End If
        End If
        
        lngResult = RegCloseKey(lngKeyHandle)
        
    End If
    
    'Se não achou das outras forma pega um padrão
    If Len(Trim(sDirPgm)) = 0 Then
        
        lngResult = 0
        lngValueType = 0
        lngKeyHandle = 0
        lngBufferSize = 0
        strBuffer = ""
        
        lngResult = RegOpenKey(HKEY_CLASSES_ROOT, "mailto\shell\open\command", lngKeyHandle)
        lngResult = RegQueryValueEx(lngKeyHandle, "", 0, lngValueType, ByVal 0, lngBufferSize)
        
        If lngValueType = REG_SZ Then
            strBuffer = String(lngBufferSize, " ")
            lngResult = RegQueryValueEx(lngKeyHandle, "", 0, 0, ByVal strBuffer, lngBufferSize)
            If lngResult = ERROR_SUCCESS Then
                If (InStr(strBuffer, Chr$(0))) > 0 Then
                    sNomePgm = left(strBuffer, (InStr(strBuffer, Chr$(0))) - 1)
                Else
                    sNomePgm = strBuffer
                End If
            End If
        End If
        
        lngResult = RegCloseKey(lngKeyHandle)
    
    End If
    
    'Se não achou de jeito nenhum abre o IE
    If Len(Trim(sDirPgm)) = 0 Then
        sDirPgm = "http://www.corporator.com.br"
    End If
       
    Obter_Pgm_Padrao_Email = sDirPgm
    
    Exit Function
    
Erro_Obter_Pgm_Padrao_Email:

    Obter_Pgm_Padrao_Email = ""

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201477)

    End Select
    
    Exit Function
    
End Function
