VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain2 
   Caption         =   "Teste LINKC"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18420
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   18420
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraLog 
      Caption         =   " Log "
      Height          =   5535
      Left            =   5520
      TabIndex        =   19
      Top             =   120
      Width           =   12855
      Begin RichTextLib.RichTextBox txtLog 
         Height          =   5175
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   9128
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmMain2.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraSimulacaoCarros 
      Caption         =   " Simulação Carros "
      Height          =   855
      Left            =   240
      TabIndex        =   14
      Top             =   1440
      Width           =   5175
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   2760
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdRunCarros 
         Caption         =   "Run"
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtQtCarros 
         Height          =   375
         Left            =   720
         TabIndex        =   15
         Text            =   "1"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Carros"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   450
      End
   End
   Begin VB.Frame fraResumo 
      Caption         =   " Informações "
      Height          =   3255
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   5175
      Begin MSFlexGridLib.MSFlexGrid grdResumo 
         Height          =   1215
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   2143
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.TextBox QtdeVeicEstaciona 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Text            =   "0"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox QtdeVeicFilaSaida 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         TabIndex        =   11
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox QtdeVeicFilaEntrada 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         TabIndex        =   10
         Text            =   "0"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade de veículos dentro do estacionamento"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   3600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade de veículos que passaram por cada entrada e saída"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   4590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade de veículos na fila de saída"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   2850
      End
      Begin VB.Label lblQtVeicFilaEntrada 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade de veículos na fila de entrada"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   2985
      End
   End
   Begin VB.Frame fraEntrSai 
      Caption         =   " Entradas / Saídas "
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtQtSaidas 
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Text            =   "2"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtQtEntradas 
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Text            =   "2"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblSaidas 
         AutoSize        =   -1  'True
         Caption         =   "Saídas"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblEntradas 
         AutoSize        =   -1  'True
         Caption         =   "Entradas"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmMain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents Timer As cSelfTimer
Attribute Timer.VB_VarHelpID = -1

Private Sub cmdStop_Click()
    Set Timer = Nothing
End Sub

Private Sub Command1_Click()
    
    
End Sub

Private Sub Timer_Timer(ByVal Seconds As Currency)
Dim dNow As Date
Dim sKeyFila As String

    Me.Caption = "Teste LINKC - " & Format$(Seconds, "0.000") & " segundos"
    
    Dim bTodosProcessados As Boolean
    bTodosProcessados = True
    For Each c In ListaCarro
        If c.StatusProcessamento <> StatusProcessamento.ProcSaidaFinalizado Then
            bTodosProcessados = False
        End If
    Next c
    
    If bTodosProcessados Then
        MsgBox "Simulação concluída.", vbOKOnly + vbInformation, "Mensagem"
        Set Timer = Nothing
        Exit Sub
    End If
    
    'Dim dNow As Date
    dNow = CDate(Now)
    
    sKeyFila = vbNullString
    For Each c In ListaCarro
        'Debug.Print "Agora: " & dNow & " TempoInicioEntrada: " & c.TempoInicioEntrada
        
        'OK Aguardando Inicio do Processamento de Entrada
        If c.StatusProcessamento = StatusProcessamento.AguardandoIniProcEntrada Then
            InsereLog txtLog, "Processando:" & c.ID & " Agora: " & dNow & " TempoInicioEntrada: " & c.TempoInicioEntrada
            If CDate(dNow) >= CDate(c.TempoInicioEntrada) Then
                'Debug.Print "... inicio do processo de entrada"
                InsereLog txtLog, "ID:" & c.ID & "... inicio do processo de entrada"
                'txtLog.Text = Text4.Text & vbCrLf & txtLog.Text
                c.StatusProcessamento = StatusProcessamento.ProcEntradaEmAndamento
                InsereLog txtLog, "ID:" & c.ID & " Status:" & DescribeStatus(c.StatusProcessamento)
                'caso não tenha
            Else
                'Debug.Print "... aguardando inicio do processamento"
                InsereLog txtLog, "ID:" & c.ID & "... aguardando inicio do processamento entrada | " & " Status:" & DescribeStatus(c.StatusProcessamento)
            End If
        End If
        
        'Ok Processamento de Entrada
        If c.StatusProcessamento = StatusProcessamento.ProcEntradaEmAndamento Then
            InsereLog txtLog, "Processando:" & c.ID & " Agora: " & dNow & " TempoFimProcessamentoEntrada: " & c.TempoFimProcessamentoEntrada
            If CDate(dNow) >= CDate(c.TempoFimProcessamentoEntrada) Then
                '####
                'Aqui verifica se tem fila de entrada disponível.
                'criar metodo no objeto lista de fila de entrada.
                'caso tenha fila disponivel loca a fila e retorna a mesma
                'realiza movimento de entrada de veiculo com a fila retornada
                'unloca a fila
                sKeyFila = oListaInOut.ESDisponivel("E")
                
                'Se tem fila disponivel Finaliza Processo de Entrada
                If sKeyFila <> vbNullString Then
                    
                    Call oListaInOut.OcupaES(ListaInOut, sKeyFila)
                    InsereLog txtLog, "Fila:" & sKeyFila & " - locada"
                    Call PrintGridResumo
                    
                    c.KeyEntrada = sKeyFila
                    c.StatusProcessamento = StatusProcessamento.ProcEntradaFinalizado
                    QtdeVeicEstaciona.Text = CStr(oEstacionamento.Soma)
                    
                    'se o carro estiver na fila de entrada tira da fila de entrada
                    
                    InsereLog txtLog, "ID:" & c.ID & "... processo de entrada finalizado"
                    c.FilaEntrada = 0
                    c.StatusProcessamento = StatusProcessamento.AguardandoIniProcSaida
                    
                    Call oListaInOut.LiberaES(ListaInOut, sKeyFila)
                    InsereLog txtLog, "Fila:" & sKeyFila & " - liberada"
                    Call PrintGridResumo
                    
                    InsereLog txtLog, "ID:" & c.ID & " Status:" & DescribeStatus(c.StatusProcessamento)
                Else
                    'coloca carro na fila de entrada
                    c.FilaEntrada = 1
                    
                End If
                
                QtdeVeicFilaEntrada.Text = CStr(oListaCarro.QtdeFilaEntrada())
                Call PrintGridResumo
                
                'Caso não haja fila de entrada disponível, segue o jogo
                'Debug.Print "... processo de entrada finalizado"
                
                
                'Debug.Print
                'Set Timer = Nothing
            Else
                'Debug.Print "... em processamento de entrada"
                InsereLog txtLog, "ID:" & c.ID & "... em processamento de entrada | " & " Status:" & DescribeStatus(c.StatusProcessamento)
            End If
        End If
        
        'Aguardando Inicio do Processamento de Saida
        If c.StatusProcessamento = StatusProcessamento.AguardandoIniProcSaida Then
            InsereLog txtLog, "Processando:" & c.ID & " Agora: " & dNow & " Tempo Inicio Saida: " & c.TempoInicioSaida
            If CDate(dNow) >= CDate(c.TempoInicioSaida) Then
                'Debug.Print "... inicio do processo de entrada"
                InsereLog txtLog, "ID:" & c.ID & "... inicio do processo de saída"
                'txtLog.Text = Text4.Text & vbCrLf & txtLog.Text
                c.StatusProcessamento = StatusProcessamento.ProcSaidaEmAndamento
                InsereLog txtLog, "ID:" & c.ID & " Status:" & DescribeStatus(c.StatusProcessamento)
            Else
                'Debug.Print "... aguardando inicio do processamento"
                InsereLog txtLog, "ID:" & c.ID & "... aguardando inicio do processamento saida | " & " Status:" & DescribeStatus(c.StatusProcessamento)
            End If
        End If
        
        'Processamento de Saida
        If c.StatusProcessamento = StatusProcessamento.ProcSaidaEmAndamento Then
            InsereLog txtLog, "Processando:" & c.ID & " Agora: " & dNow & " Tempo Fim Processamento Saida: " & c.TempoFimProcessamentoSaida
            If CDate(dNow) >= CDate(c.TempoFimProcessamentoSaida) Then
                
'                InsereLog txtLog, "ID:" & c.ID & "... processo de saída finalizado"
'                c.StatusProcessamento = StatusProcessamento.ProcSaidaFinalizado
'                QtdeVeicEstaciona.Text = CStr(oEstacionamento.Subtrai)
'
'                InsereLog txtLog, "ID:" & c.ID & " Status:" & DescribeStatus(c.StatusProcessamento)
                
                '####
                'Aqui verifica se tem fila de entrada disponível.
                'criar metodo no objeto lista de fila de entrada.
                'caso tenha fila disponivel loca a fila e retorna a mesma
                'realiza movimento de entrada de veiculo com a fila retornada
                'unloca a fila
                sKeyFila = oListaInOut.ESDisponivel("S")
                
                'Se tem fila disponivel Finaliza Processo de Entrada
                If sKeyFila <> vbNullString Then
                    Call oListaInOut.OcupaES(ListaInOut, sKeyFila)
                    InsereLog txtLog, "Fila:" & sKeyFila & " - locada"
                    Call PrintGridResumo


                    c.KeySaida = sKeyFila
                    c.StatusProcessamento = StatusProcessamento.ProcSaidaFinalizado
                    QtdeVeicEstaciona.Text = CStr(oEstacionamento.Subtrai)
                    InsereLog txtLog, "ID:" & c.ID & "... processo de saída finalizado"
                    c.FilaSaida = 0
                    'c.StatusProcessamento = StatusProcessamento.AguardandoIniProcSaida
                    
                    Call oListaInOut.LiberaES(ListaInOut, sKeyFila)
                    InsereLog txtLog, "Fila:" & sKeyFila & " - liberada"
                    Call PrintGridResumo
                    
                    InsereLog txtLog, "ID:" & c.ID & " Status:" & DescribeStatus(c.StatusProcessamento)
                Else
                    'continua na fila aguardando uma saida disponível.
                    'QtdeVeicFilaSaida.Text = CStr(oEstacionamento.SomaFilaSaida)
                    c.FilaSaida = 1
                End If
                QtdeVeicFilaSaida.Text = CStr(oListaCarro.QtdeFilaSaida())
                Call PrintGridResumo
                'Caso não haja fila de entrada disponível, segue o jogo
            Else
                'Debug.Print "... em processamento de entrada"
                InsereLog txtLog, "ID:" & c.ID & "... em processamento de saída | " & " Status:" & DescribeStatus(c.StatusProcessamento)
            End If
        End If

        Me.Refresh
    Next c
    
End Sub

Private Sub cmdRunCarros_Click()
    Set ListaInOut = New Collection
    Set ListaCarro = New Collection
    Set oListaInOut = New cListaInOut
    Set oListaCarro = New cListaCarro
    Set oEstacionamento = New cEstacionamento
    
    Set Timer = New cSelfTimer
    
    'tempo de intervalo do timer em milisgundos
    Timer.Interval = 500
    
    txtLog.Text = vbNullString

    If Not IsNumeric(txtQtEntradas.Text) Or Not IsNumeric(txtQtSaidas.Text) Or Not IsNumeric(txtQtCarros.Text) Then
        MsgBox "Entradas, saídas e carros devem ser numéricos.", vbOKOnly + vbCritical, "Mensagem"
        Exit Sub
    End If


    Set ListaCarro = New Collection
    Call oListaInOut.InitListaInOut(CInt(txtQtEntradas.Text), CInt(txtQtSaidas.Text))
    
    Call oListaInOut.OcupaES(ListaInOut, "E1")
    'Call oListaInOut.OcupaES("S2")

    
    Call PrintGridResumo
    
    Call oListaCarro.AddListaCarro(CInt(txtQtCarros.Text))
    For Each kk In ListaCarro
        InsereLog txtLog, "Id:" & kk.ID & " Tempo: " & kk.Tempo & _
            " Tempo Inicio Entrada: " & kk.TempoInicioEntrada & _
            " Tempo Fim Processamento Entrada: " & kk.TempoFimProcessamentoEntrada & _
            " Tempo Inicio Saída: " & kk.TempoInicioSaida & _
            " Tempo Fim Processamento Saída: " & kk.TempoFimProcessamentoSaida & _
            " Status: " & DescribeStatus(kk.StatusProcessamento)
    Next
    
End Sub

Public Function InCollection(col As Collection, sKey As String) As Boolean

Dim bTest As Boolean

    On Error Resume Next
    
    bTest = IsObject(col(sKey))
    If (Err = 0) Then
        InCollection = True
    Else
        Err.Clear
    End If

End Function

Public Function InCollectionXXX(col As Collection, sKey As String) As Object

Dim bTest As Boolean

    On Error Resume Next
    
    bTest = IsObject(col(sKey))
    If (Err = 0) Then
        'InCollection = True
        Set InCollectionXXX = col(sKey)
    Else
        Err.Clear
    End If

End Function

Private Sub PrintGridResumo()
    grdResumo.Clear
    grdResumo.Rows = 0
    grdResumo.Cols = 4
    grdResumo.AddItem ("ID" & vbTab & "Tipo" & vbTab & "Status" & vbTab & "Qtde")
    'grdResumo.Rows = 2
    'grdResumo.FixedRows = 1
    grdResumo.FixedRows = 0
    grdResumo.FixedCols = 0
    'grdResumo.col = 0
    'grdResumo.Text = "Col1"
    grdResumo.Enabled = False
    For Each kk In ListaInOut
        'Debug.Print kk.ID & kk.Status
        grdResumo.AddItem (kk.ID & vbTab & IIf(kk.Tipo = "E", "Entrada", "Saída") & vbTab & IIf(kk.Status = 1, "Liberada", "Ocupada") & vbTab & kk.Qtde)
        'grdResumo.col = 1
        If kk.Status = 1 Then
            
            grdResumo.col = 0
            grdResumo.Row = grdResumo.Rows - 1
            grdResumo.CellBackColor = vbGreen
        Else
            grdResumo.col = 0
            grdResumo.Row = grdResumo.Rows - 1
            grdResumo.CellBackColor = vbRed
        End If
    Next

    If grdResumo.Rows >= 2 Then
        grdResumo.FixedRows = 1
    Else
        grdResumo.AddItem ("" & vbTab & "" & vbTab & "" & vbTab & "")
        grdResumo.FixedRows = 1
    End If
    grdResumo.Enabled = True

End Sub

