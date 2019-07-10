VERSION 5.00
Begin VB.Form FrmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MousePointer    =   11  'Hourglass
   Picture         =   "FrmSplash.frx":0000
   ScaleHeight     =   5010
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6750
      Top             =   4425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde carregando..."
      Height          =   240
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   2490
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const FLAG_ICC_FORCE_CONNECTION = &H1
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

Private Sub Form_Load()
Me.Width = 488 * 15
Me.Height = 332 * 15



End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
    If InternetCheckConnection("http://www.alankoba.com.br/", FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
        MsgBox "Falha ao conectar-se no servidor de totalização." + vbNewLine + "Tente novamente mais tarde, ou verifique as configurações de sua conexão.", vbCritical, "Erro na conexão"
        End
    Else
        FrmVerifica.Show
        Unload Me
    End If
    
End Sub
