VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FrmVerifica 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificação do titulo"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   Icon            =   "FrmVerifica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser webv 
      Height          =   765
      Left            =   6975
      TabIndex        =   18
      Top             =   3075
      Width           =   1140
      ExtentX         =   2011
      ExtentY         =   1349
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.ComboBox combodia 
      Height          =   315
      Left            =   525
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1275
      Width           =   690
   End
   Begin VB.ComboBox comboano 
      Height          =   315
      Left            =   3150
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1275
      Width           =   1065
   End
   Begin VB.ComboBox combomes 
      Height          =   315
      ItemData        =   "FrmVerifica.frx":08CA
      Left            =   1425
      List            =   "FrmVerifica.frx":08F5
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1275
      Width           =   1515
   End
   Begin VB.TextBox txttitulo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   540
      TabIndex        =   12
      Top             =   2865
      Width           =   1350
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   525
      TabIndex        =   16
      Top             =   5550
      Width           =   1875
   End
   Begin VB.CheckBox naoconsta 
      BackColor       =   &H00FFFFFF&
      Caption         =   "não consta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3000
      TabIndex        =   9
      Top             =   1800
      Width           =   1365
   End
   Begin VB.TextBox nomemae 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   525
      TabIndex        =   10
      Top             =   2100
      Width           =   3750
   End
   Begin VB.TextBox txtnome 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   555
      TabIndex        =   1
      Top             =   495
      Width           =   3750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continuar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2925
      TabIndex        =   17
      Top             =   5475
      Width           =   1725
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   870
      Left            =   6975
      TabIndex        =   15
      Top             =   3900
      Width           =   1125
      ExtentX         =   1984
      ExtentY         =   1535
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin SHDocVwCtl.WebBrowser WEB2 
      Height          =   1545
      Left            =   375
      TabIndex        =   13
      Top             =   3300
      Width           =   3900
      ExtentX         =   6879
      ExtentY         =   2725
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Se não conseguir acessar, verifique se os dados estão realmente iguais aos do seu cadastro no TSE."
      Height          =   465
      Left            =   675
      TabIndex        =   20
      Top             =   6975
      Width           =   4290
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   75
      Picture         =   "FrmVerifica.frx":095E
      Top             =   6255
      Width           =   480
   End
   Begin VB.Label label500 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo de eleitor:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   525
      TabIndex        =   11
      Top             =   2550
      Width           =   2010
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Digite o código acima:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   525
      TabIndex        =   14
      Top             =   5250
      Width           =   1860
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome da mãe:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   525
      TabIndex        =   8
      Top             =   1800
      Width           =   1260
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Data de nascimento:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   525
      TabIndex        =   2
      Top             =   975
      Width           =   2040
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3000
      TabIndex        =   7
      Top             =   1350
      Width           =   315
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1275
      TabIndex        =   6
      Top             =   1350
      Width           =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome completo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   525
      TabIndex        =   0
      Top             =   150
      Width           =   5010
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmVerifica.frx":1228
      Height          =   615
      Left            =   675
      TabIndex        =   19
      Top             =   6225
      Width           =   4215
   End
End
Attribute VB_Name = "FrmVerifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias _
                         "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As _
                                                                        Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10
                
Dim fonte As String                              ' codigo fonte

Private Sub combodia_Click()
    If CInt(combodia.Text) > 29 And combomes.Text = "FEVEREIRO" Then
        combodia.ListIndex = 28
    End If
End Sub

Private Sub combomes_Click()
    If CInt(combodia.Text) > 29 And combomes.Text = "FEVEREIRO" Then
        combodia.ListIndex = 28
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    If txtcodigo.Text = "" Or txtnome.Text = "" Or nomemae.Text = "" Then
        MsgBox "Algum campo não foi preenchido", vbCritical, "Erro"
        Exit Sub
    End If

    mes = "0" & combomes.ItemData(combomes.ListIndex)
    mes = Right(mes, 2)
    dia = "0" & combodia.Text
    dia = Right(dia, 2)


    web.Document.consultaLocalVotacaoNomeForm.nomeEleitor.Value = UCase(txtnome.Text)
    web.Document.consultaLocalVotacaoNomeForm.dataNascimento.Value = dia & "/" & mes & "/" & comboano.Text
    web.Document.consultaLocalVotacaoNomeForm.nomemae.Value = UCase(nomemae.Text)
    web.Document.All.naoConstaMae.Checked = naoconsta.Value
    web.Document.consultaLocalVotacaoNomeForm.codigoCaptcha.Value = txtcodigo.Text
    web.Document.consultaLocalVotacaoNomeForm.consultar.Click
End Sub

Private Sub Form_Load()
    'Frame1.Enabled = False
    web.Silent = True
    web.Navigate "http://www.tse.jus.br/certidaoquitacao/consultaLocalVotacaoNome.do?dispatcher=exibirConsultaLocalVotacaoNome&validate=false"
    For i = 1996 To 1900 Step -1
        comboano.AddItem (i)
    Next i

    For i = 1 To 31
        combodia.AddItem (i)
    Next i

    combodia.ListIndex = 0
    combomes.ListIndex = 0
    comboano.ListIndex = 0

End Sub

Private Sub naoconsta_Click()
    If naoconsta.Value = Checked Then
        nomemae.Enabled = False
        nomemae.Text = "NÃO CONSTA"
    Else
        nomemae.Enabled = True
        nomemae.Text = ""
    End If
End Sub

Private Sub txtano_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKey0 To vbKey9
    Case vbKeyBack, vbKeyClear, vbKeyDelete
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
    Case Else
        KeyAscii = 0
        Beep
    End Select
End Sub

Private Sub txtdia_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKey0 To vbKey9
    Case vbKeyBack, vbKeyClear, vbKeyDelete
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
    Case Else
        KeyAscii = 0
        Beep
    End Select
End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKey0 To vbKey9
    Case vbKeyBack, vbKeyClear, vbKeyDelete
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
    Case Else
        KeyAscii = 0
        Beep
    End Select
End Sub

Private Sub web_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next


    'web.Document.body.bgcolor = ""
    a = web.Document.getelementsbytagname("img")(0).src

    WEB2.Navigate a
    fonte = web.Document.body.innerhtml


    If InStr(1, fonte, "O código informado não corresponde. Por ") Then
        MsgBox "O Código de segurança informado esta incorreto!", vbCritical, "Erro"
        web.Navigate "http://www.tse.jus.br/certidaoquitacao/consultaLocalVotacaoNome.do?dispatcher=exibirConsultaLocalVotacaoNome&validate=false"
    End If


    If InStr(1, fonte, "PIRANGI") And InStr(1, fonte, txttitulo.Text) Then
        Titulo_eleitor = txttitulo.Text
        Nome_eleitor = txtnome.Text
        webv.Navigate "http://www.alankoba.com.br/urna/verificar.php?titulo=" & Titulo_eleitor
    Else
    
        If InStr(1, fonte, "Os dados informados (nome, data de nascimento e/ou filiação)") Then
            MsgBox "Alguma informação não confere com os dados cadastrados no TSE.", vbCritical, "Erro na verificação"
            web.Navigate "http://www.tse.jus.br/certidaoquitacao/consultaLocalVotacaoNome.do?dispatcher=exibirConsultaLocalVotacaoNome&validate=false"
            txtcodigo.Text = ""
        End If
                
        If InStr(1, fonte, "DD/MM/AAAA") = False Then
            MsgBox "Digitou alguma informação errada, tente novamente.", vbCritical, "Erro"
            web.Navigate "http://www.tse.jus.br/certidaoquitacao/consultaLocalVotacaoNome.do?dispatcher=exibirConsultaLocalVotacaoNome&validate=false"
            txtcodigo.Text = ""
        End If
    End If
End Sub

Private Sub WEB2_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    WEB2.Document.body.Style.border = "none"
    WEB2.Document.body.Scroll = "no"
End Sub

Private Sub webv_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error GoTo hell
    If InStr(1, webv.Document.body.innerhtml, "PODEIR") Then
        FrmInicio.Show
        Unload Me
    End If
    If InStr(1, webv.Document.body.innerhtml, "JAVOTOU") Then
        MsgBox "Este titulo já votou", vbCritical, "Erro"
        End
    End If
    Exit Sub
    hell:
End Sub