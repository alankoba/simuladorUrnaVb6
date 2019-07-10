VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FrmInicio 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   Icon            =   "FrmInicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmInicio.frx":08CA
   ScaleHeight     =   5505
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer sairt 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3750
      Top             =   75
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2115
      Left            =   2100
      TabIndex        =   47
      Top             =   5550
      Width           =   5640
      ExtentX         =   9948
      ExtentY         =   3731
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
   Begin VB.Timer piscador 
      Interval        =   805
      Left            =   9000
      Top             =   4500
   End
   Begin VB.Frame frameFIM 
      BorderStyle     =   0  'None
      Height          =   2940
      Left            =   375
      TabIndex        =   34
      Top             =   7725
      Width           =   4665
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "FIM"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   0
         TabIndex        =   35
         Top             =   975
         Width           =   4590
      End
   End
   Begin VB.ListBox listav 
      Height          =   1815
      Left            =   4800
      TabIndex        =   23
      Top             =   5700
      Width           =   4515
   End
   Begin VB.ListBox listap 
      Height          =   1815
      Left            =   300
      TabIndex        =   22
      Top             =   5700
      Width           =   4440
   End
   Begin VB.Timer pisca 
      Interval        =   100
      Left            =   9000
      Top             =   4950
   End
   Begin VB.Frame tela 
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   675
      TabIndex        =   13
      Top             =   1875
      Width           =   4740
      Begin VB.Frame framefotos 
         BorderStyle     =   0  'None
         Height          =   2865
         Left            =   2925
         TabIndex        =   24
         Top             =   75
         Visible         =   0   'False
         Width           =   1815
         Begin VB.Image fotov 
            Appearance      =   0  'Flat
            Height          =   1140
            Left            =   750
            Stretch         =   -1  'True
            Top             =   1650
            Width           =   1065
         End
         Begin VB.Image fotop 
            Appearance      =   0  'Flat
            Height          =   1590
            Left            =   300
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1515
         End
      End
      Begin VB.Frame framemsg 
         BorderStyle     =   0  'None
         Height          =   840
         Left            =   75
         TabIndex        =   25
         Top             =   2125
         Visible         =   0   'False
         Width           =   4665
         Begin VB.Line Line1 
            X1              =   0
            X2              =   4725
            Y1              =   150
            Y2              =   150
         End
         Begin VB.Label votolegenda 
            Caption         =   "(voto de legenda)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3225
            TabIndex        =   33
            Top             =   600
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label Label3 
            Caption         =   "LARANJA para REINICIAR este voto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   75
            TabIndex        =   28
            Top             =   675
            Width           =   3315
         End
         Begin VB.Label Label2 
            Caption         =   "VERDE para confirmar este voto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   75
            TabIndex        =   27
            Top             =   450
            Width           =   3165
         End
         Begin VB.Label Label1 
            Caption         =   "Aperte a tecla:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   75
            TabIndex        =   26
            Top             =   225
            Width           =   1140
         End
      End
      Begin VB.Label legendagrande 
         Caption         =   "VOTO DE LEGENDA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   46
         Top             =   1800
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label inexistente 
         Caption         =   "CANDIDATO INEXISTENTE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   45
         Top             =   1275
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.Label titulo 
         Caption         =   "ELEIÇÕES 2012 - PIRANGI-SP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   44
         Top             =   75
         Width           =   2565
      End
      Begin VB.Label voteinulo 
         Caption         =   "VOTO NULO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1275
         TabIndex        =   43
         Top             =   1800
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label nerrado 
         Caption         =   "NÚMERO ERRADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   41
         Top             =   1275
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.Label lblvice2 
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
         Left            =   1050
         TabIndex        =   40
         Top             =   1950
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Label lblvice 
         Caption         =   "Vice-prefeito:"
         Height          =   240
         Left            =   0
         TabIndex        =   39
         Top             =   1950
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label votobranco 
         Caption         =   "VOTO EM BRANCO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   38
         Top             =   825
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.Label lblpartido 
         Caption         =   "Partido:"
         Height          =   240
         Left            =   0
         TabIndex        =   31
         Top             =   1725
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblnome2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   525
         TabIndex        =   30
         Top             =   1275
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label lblnome 
         Caption         =   "Nome:"
         Height          =   240
         Left            =   0
         TabIndex        =   29
         Top             =   1275
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblnumero 
         Caption         =   "Número:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   0
         TabIndex        =   21
         Top             =   1030
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblcargo 
         Caption         =   "CANDIDATO(A)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   375
         TabIndex        =   20
         Top             =   525
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "SEU VOTO PARA"
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
         Left            =   225
         TabIndex        =   19
         Top             =   300
         Width           =   2490
      End
      Begin VB.Label nvoto 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   4
         Left            =   2250
         TabIndex        =   18
         Top             =   835
         Width           =   270
      End
      Begin VB.Label nvoto 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   3
         Left            =   1875
         TabIndex        =   17
         Top             =   835
         Width           =   270
      End
      Begin VB.Label nvoto 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   2
         Left            =   1500
         TabIndex        =   16
         Top             =   835
         Width           =   270
      End
      Begin VB.Label nvoto 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   1
         Left            =   1125
         TabIndex        =   15
         Top             =   835
         Width           =   270
      End
      Begin VB.Label nvoto 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Index           =   0
         Left            =   750
         TabIndex        =   14
         Top             =   835
         Width           =   270
      End
      Begin VB.Shape quadro 
         BackColor       =   &H8000000F&
         BorderColor     =   &H00808080&
         Height          =   405
         Index           =   4
         Left            =   2175
         Top             =   810
         Width           =   405
      End
      Begin VB.Shape quadro 
         BackColor       =   &H8000000F&
         BorderColor     =   &H00808080&
         Height          =   405
         Index           =   3
         Left            =   1800
         Top             =   810
         Width           =   405
      End
      Begin VB.Shape quadro 
         BackColor       =   &H8000000F&
         BorderColor     =   &H00808080&
         Height          =   405
         Index           =   2
         Left            =   1425
         Top             =   810
         Width           =   405
      End
      Begin VB.Shape quadro 
         BackColor       =   &H8000000F&
         BorderColor     =   &H00808080&
         Height          =   405
         Index           =   1
         Left            =   1050
         Top             =   810
         Width           =   405
      End
      Begin VB.Shape quadro 
         BackColor       =   &H8000000F&
         BorderColor     =   &H00808080&
         Height          =   405
         Index           =   0
         Left            =   675
         Top             =   810
         Width           =   405
      End
      Begin VB.Label lblpartido2 
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
         Left            =   600
         TabIndex        =   32
         Top             =   1725
         Visible         =   0   'False
         Width           =   2265
      End
   End
   Begin VB.Label Label9 
      Caption         =   "NÚMERO ERRADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1425
      TabIndex        =   42
      Top             =   3750
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   375
      TabIndex        =   37
      Top             =   375
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   375
      TabIndex        =   36
      Top             =   75
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label numero 
      BackStyle       =   0  'Transparent
      Height          =   465
      Index           =   300
      Left            =   7440
      TabIndex        =   12
      Top             =   4290
      Width           =   675
   End
   Begin VB.Label numero 
      BackStyle       =   0  'Transparent
      Height          =   405
      Index           =   200
      Left            =   6705
      TabIndex        =   11
      Top             =   4350
      Width           =   615
   End
   Begin VB.Label numero 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   100
      Left            =   6000
      TabIndex        =   10
      Top             =   4350
      Width           =   585
   End
   Begin VB.Label numero 
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   0
      Left            =   6795
      TabIndex        =   9
      Top             =   3885
      Width           =   360
   End
   Begin VB.Label numero 
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   9
      Left            =   7380
      TabIndex        =   8
      Top             =   3420
      Width           =   360
   End
   Begin VB.Label numero 
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   8
      Left            =   6765
      TabIndex        =   7
      Top             =   3420
      Width           =   360
   End
   Begin VB.Label numero 
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   7
      Left            =   6180
      TabIndex        =   6
      Top             =   3420
      Width           =   360
   End
   Begin VB.Label numero 
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   6
      Left            =   7350
      TabIndex        =   5
      Top             =   2955
      Width           =   360
   End
   Begin VB.Label numero 
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   5
      Left            =   6735
      TabIndex        =   4
      Top             =   2955
      Width           =   360
   End
   Begin VB.Label numero 
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   4
      Left            =   6150
      TabIndex        =   3
      Top             =   2955
      Width           =   360
   End
   Begin VB.Label numero 
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   3
      Left            =   7365
      TabIndex        =   2
      Top             =   2475
      Width           =   360
   End
   Begin VB.Label numero 
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   2
      Left            =   6750
      TabIndex        =   1
      Top             =   2475
      Width           =   360
   End
   Begin VB.Label numero 
      BackStyle       =   0  'Transparent
      Height          =   285
      Index           =   1
      Left            =   6165
      TabIndex        =   0
      Top             =   2475
      Width           =   360
   End
End
Attribute VB_Name = "FrmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////
' Busca no Listbox
Private Const CB_FINDSTRING = &H14C
Private Const LB_FINDSTRING = &H18F
Private Declare Function SendMessage Lib _
                         "user32" Alias "SendMessageA" (ByVal _
                                                        hWnd As Long, ByVal wMsg As Long, _
                                                        ByVal wParam As Long, lParam As Any) _
                                                        As Long
' //////////////////////////////////////////////////////////////////////////////////
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias _
                         "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As _
                                                                        Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const IDC_HAND As Long = &H7F89          '32649
  
' pisca o quadro'
Dim cont As Integer                              ' 1 até xxx

Private Declare Sub ReleaseCapture Lib "user32" ()

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
      
'Variaveis do voto
Dim voto As String
Dim cargo As String
' totais
Dim totalp As Integer                            ' total de prefeitos
Dim totalv As Integer                            ' total de vereadores
Dim nulo As Boolean
Dim branco As Boolean
Dim legenda As Boolean
''''''''''''''''''''
' sempre no topo
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Function AlwaysOnTop(FrmID As Form, ByVal OnTop As Boolean) As Boolean
    On Error Resume Next
    Rem --- Deixa o form sempre no topo das demais janelas abertas no Windows ---
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2

    If OnTop = True Then
        AlwaysOnTop = SetWindowPos(FrmID.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        AlwaysOnTop = SetWindowPos(FrmID.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If

End Function

' codigo daqui pra baixo
Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()

    'Titulo_eleitor = "5415454"
    'Nome_eleitor = "ALAN KOBAYASHI"



    On Error Resume Next
    'MsgBox Command
    'MsgBox App.Path
    cargo = "VEREADOR"
    lblcargo.Caption = cargo & "(A)"
    ActiveTransparency Me, True, True, 255, &HFF00FF
    ' Branco = Index 100
    ' Corrige = Index 200
    ' Confirma = Index 300
    totalp = CInt(ReadINI(App.Path & "/Candidatos/prefeito.ini", "URNA", "total"))
    totalv = CInt(ReadINI(App.Path & "/Candidatos/vereador.ini", "URNA", "total"))

    For i = 1 To totalp
        listap.AddItem CStr((ReadINI(App.Path & "/Candidatos/prefeito.ini", i, "numero"))) & ";" & _
                                                                                           (ReadINI(App.Path & "/Candidatos/prefeito.ini", i, "p")) & ";" & (ReadINI(App.Path & "/Candidatos/prefeito.ini", i, "v"))

    Next i

    For i = 1 To totalv
        listav.AddItem CStr((ReadINI(App.Path & "/Candidatos/vereador.ini", i, "numero"))) & ";" & _
                                                                                           (ReadINI(App.Path & "/Candidatos/vereador.ini", i, "nome"))

    Next i

    'AlwaysOnTop Me, True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    Dim lngReturnValue As Long

    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, _
                                     HTCAPTION, 0&)
    End If
End Sub

Private Sub listap_Click()
    If listap.ListIndex <> -1 And Len(voto) = 2 Then ' se não existir
        On Error Resume Next
        If cargo = "PREFEITO" Then               ' evita erros
            Dim temp() As String
            temp() = Split(listap.List(listap.ListIndex), ";")
            lblnome2.Caption = temp(1)
            lblvice2.Caption = temp(2)
            fotop.Picture = LoadPicture(App.Path & "/Candidatos/fotos/" & temp(0) & ".jpg")
            fotov.Picture = LoadPicture(App.Path & "/Candidatos/fotos/" & temp(0) & "v.jpg")
            framefotos.Visible = True
        End If

        If cargo = "PREFEITO" And Len(voto) = 2 Then
            lblnome.Visible = True
            lblnome2.Visible = True
            lblvice.Visible = True
            lblvice2.Visible = True
        End If
    
    ElseIf listap.ListIndex = -1 Then
        lblnome.Visible = False
        'lblnome2.Caption = "NÚMERO ERRADO"
        nulo = True
    End If

    ' ARRUMA BUG DOS QUADROS
    quadro(0).Visible = True
    quadro(1).Visible = True

End Sub

Private Sub listav_Click()




    If listav.ListIndex <> -1 And Len(voto) = 5 Then
        On Error Resume Next
        If cargo = "VEREADOR" Then               ' evita erros
            Dim temp() As String
            temp() = Split(listav.List(listav.ListIndex), ";")
            lblnome2.Caption = temp(1)
            fotop.Picture = LoadPicture(App.Path & "/Candidatos/fotos/" & temp(0) & ".jpg")
            framefotos.Visible = True
        End If

        If cargo = "VEREADOR" And Len(voto) = 5 Then
            lblnome.Visible = True
            lblnome2.Visible = True
        End If
        ' arruma bug de não mostrar os quadros

        quadro(0).Visible = True
        quadro(1).Visible = True
        quadro(2).Visible = True
        quadro(3).Visible = True
        quadro(4).Visible = True
        
    ElseIf listav.ListIndex = 1 Then
        nulo = True
    End If
End Sub

Private Sub numero_Click(Index As Integer)
    Dim votou As String
    Select Case Index
    Case 0 To 200
        If Index < 100 Then
            If Len(voto) < 5 And cargo = "VEREADOR" Then
                votou = Index
                voto = voto & votou
                nvoto(Len(voto) - 1).Caption = votou
            ElseIf Len(voto) < 2 And cargo = "PREFEITO" Then
                votou = Index
                voto = voto & votou
                nvoto(Len(voto) - 1).Caption = votou
            End If
                
            If Len(voto) = 5 And cargo = "VEREADOR" Then
                FindLB listav, voto              ' procura
                Exit Sub
            ElseIf Len(voto) = 2 And cargo = "PREFEITO" Then
                FindLB listap, voto              ' procura
                Exit Sub
            End If
        End If
        
        If Index = 100 And Len(voto) = 0 Then    ' branco
            votolegenda.Visible = False
            voto = "BRANCO"
        
            If cargo = "PREFEITO" Then
                Nome_Prefeito = "BRANCO"
            Else
                Nome_Vereador = "BRANCO"
            End If
        
            branco = True
            SoundName$ = App.Path & "/Candidatos/botao.wav"
            wFlags% = SND_ASYNC Or SND_NODEFAULT
            X% = sndPlaySound(SoundName$, wFlags%)
        End If
        
        If Index = 200 Then                      ' corrige
            Call Limpar
        End If
        
        ' tocar o som do botão
        If Index <> 100 Then
            SoundName$ = App.Path & "/Candidatos/botao.wav"
            wFlags% = SND_ASYNC Or SND_NODEFAULT
            X% = sndPlaySound(SoundName$, wFlags%)
        End If
        
    Case 300                                     ' confirmação
        If Len(voto) < 2 Then
            Exit Sub
        End If

        
        If cargo = "PREFEITO" Then
            SoundName$ = App.Path & "/Candidatos/confirmau.wav"
            wFlags% = SND_ASYNC Or SND_NODEFAULT
            X% = sndPlaySound(SoundName$, wFlags%)
        Else
            SoundName$ = App.Path & "/Candidatos/confirma.wav"
            wFlags% = SND_ASYNC Or SND_NODEFAULT
            X% = sndPlaySound(SoundName$, wFlags%)
        End If
        
        
        
        If cargo = "VEREADOR" Then
            cargo = "PREFEITO"
            quadro(4).Visible = False
            quadro(3).Visible = False
            quadro(2).Visible = False
            lblcargo.Caption = cargo & "(A)"
            If nulo = True Then
                Voto_Vereador = "NULO-" & voto
                Nome_Vereador = "NULO"
            Else
                Voto_Vereador = voto             ' salva o voto
                Nome_Vereador = lblnome2.Caption
            End If
                
            If legenda = True Then
                Voto_Vereador = "LEG-" & voto    ' salva o voto
                Nome_Vereador = "LEGENDA"
            End If
                
        Else
            Call DesativarTudo
            frameFIM.Top = tela.Top
            frameFIM.Left = tela.Left
            If nulo = True Then
                Voto_Prefeito = "NULO-" & voto
                Nome_Prefeito = "NULO"
            Else
                Voto_Prefeito = voto             ' salva o voto
                Nome_Prefeito = lblnome2.Caption
            End If
            ' Call FazWeb
       
            If Voto_Prefeito = "BRANCO" Then
                Nome_Prefeito = "BRANCO"
            End If
            If Voto_Vereador = "BRANCO" Then
                Nome_Vereador = "BRANCO"
            End If
       
            '        web.Navigate "http://www.alankoba.com.br/urna/votar.php?nomev=" & Nome_Vereador & _
            '        "&nomep=" & Nome_Prefeito & _
            '        "&eleitor=" & Nome_eleitor & _
            '        "&titulo=" & Titulo_eleitor & _
            '        "&numerop=" & Voto_Prefeito & _
            '        "&numerov=" & Voto_Vereador

            sairt.Enabled = True

       
        End If
        
        Call Limpar
        
    End Select
End Sub

Function FazWeb()
    'web.Document.Formulario.Nome_Vereador.Value = Nome_Vereador
    'web.Document.Formulario.Nome_Prefeito.Value = Nome_Prefeito
    'web.Document.Formulario.numerop.Value = Voto_Prefeito
    'web.Document.Formulario.numerov.Value = Voto_Vereador
    'web.Document.Formulario.titulo.Value = titulo_eleitor
    'web.Document.Formulario.eleitor.Value = Nome_eleitor
    'web.Document.Formulario.submit
End Function

Private Sub numero_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Function Limpar()
    lblvice2.Caption = ""
    lblnome2.Caption = ""
    lblpartido2.Caption = ""
    legenda = False
    voto = ""
    votou = ""
    nvoto(0).Caption = ""
    nvoto(1).Caption = ""
    nvoto(2).Caption = ""
    nvoto(3).Caption = ""
    nvoto(4).Caption = ""
    framefotos.Visible = False
    lblnome.Visible = False
    lblnome2.Visible = False
    nulo = False
    branco = False
    lblvice.Visible = False
    lblvice2.Visible = False
    votobranco.Visible = False
    If cargo = "VEREADOR" Then
        quadro(0).Visible = True
        quadro(1).Visible = True
        quadro(2).Visible = True
        quadro(3).Visible = True
        quadro(4).Visible = True
    Else
        quadro(0).Visible = True
        quadro(1).Visible = True
    End If
        
End Function

Private Sub pisca_Timer()



    On Error Resume Next
    quadro(Len(voto) - 1).Visible = True
    If cargo = "PREFEITO" And Len(voto) = 2 Then
        ' ARRUMA BUG DOS QUADROS
        quadro(0).Visible = True
        quadro(1).Visible = True
        
        If lblpartido2.Caption <> "" And lblnome2.Caption = "" Then
            lblpartido.Visible = False
            lblpartido2.Visible = False
            lblvice.Visible = False
            lblvice2.Visible = False
            nulo = True
        End If
        
    End If



    If framefotos.Visible = True Then
        legenda = False
    End If


    If Len(voto) = 5 Then
        If lblnome2.Caption = "" Then
            inexistente.Visible = True
            If nulo = False Then
                legendagrande.Visible = True
                votolegenda.Visible = False
                legenda = True
            End If
        End If
    Else
        inexistente.Visible = False
        legendagrande.Visible = False

        If lblpartido2.Caption <> "" Then
            legenda = True
        End If

    End If


    If nulo = True Then
        nerrado.Visible = True
        voteinulo.Visible = True
    Else
        nerrado.Visible = False
        voteinulo.Visible = False
    End If




    '
    Label6.Caption = Voto_Vereador
    Label7.Caption = Voto_Prefeito
    '
    If branco = True Then
        votobranco.Visible = True
        quadro(0).Visible = False
        quadro(1).Visible = False
        quadro(2).Visible = False
        quadro(3).Visible = False
        quadro(4).Visible = False
        framemsg.Visible = True
        Exit Sub
    End If

    'On Error GoTo err:

    If Len(voto) >= 2 Then

        If cargo = "PREFEITO" Then
            'Exit Sub
        End If

        If Len(voto) >= 2 And Len(voto) < 5 Then
            lblpartido2.Caption = PegaPartido(Left(voto, 2))
            If PegaPartido(Left(voto, 2)) <> "" Then
                votolegenda.Visible = True
            Else
                votolegenda.Visible = False
            End If
        End If

        framemsg.Visible = True
        lblnumero.Visible = True
        If cargo = "PREFEITO" And nulo = False And branco = False Then
            lblnome.Visible = True
            lblnome2.Visible = True
            lblvice.Visible = True
            lblvice2.Visible = True
        End If
    

        If nulo = False And branco = False Then
            lblpartido.Visible = True
            lblpartido2.Visible = True
        End If
    Else
        framemsg.Visible = False
        lblnumero.Visible = False
        'lblnome.Visible = False
        'lblnome2.Visible = False
        lblpartido.Visible = False
        lblpartido2.Visible = False
    End If

    Exit Sub
    err:

End Sub

Private Sub FindLB(obj As Object, TextToFind As String)
    obj.ListIndex = SendMessage( _
                    obj.hWnd, LB_FINDSTRING, -1, ByVal _
                                                TextToFind)
End Sub

Function PegaPartido(voto As String) As String
    Select Case voto
    Case "45"
        PegaPartido = "PSDB"
    Case "55"
        PegaPartido = "PSD"
    Case "13"
        PegaPartido = "PT"
    Case "14"
        PegaPartido = "PTB"
    Case "43"
        PegaPartido = "PV"
    Case "15"
        PegaPartido = "PMDB"
    Case "40"
        PegaPartido = "PSB"
    Case "25"
        PegaPartido = "DEM"
    Case "11"
        PegaPartido = "PP"
    Case "12"
        PegaPartido = "PDT"
    Case Else
        nulo = True
    End Select
End Function

Function DesativarTudo()
    On Error Resume Next
    For i = 0 To 300
        numero(i).Enabled = False
    Next i
End Function

Private Sub sairt_Timer()
    End
End Sub

Private Sub piscador_Timer()
    On Error Resume Next
    If cargo <> "PREFEITO" Then
        quadro(Len(voto)).Visible = Not quadro(Len(voto)).Visible
        quadro(Len(voto) - 1).Visible = True
    Else
        If Len(voto) < 2 Then
            quadro(Len(voto)).Visible = Not quadro(Len(voto)).Visible
            quadro(Len(voto) - 1).Visible = True
        End If
    End If
End Sub

Private Sub web_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    'On Error GoTo err:
    'Dim fonte As String
    'fonte = web.Document.body.innerhtml
    '
    'If Left(web.Document.URL, 10) = "http://www" Then
    '
    '
    '    If InStr(1, fonte, "DEUCERTODEUCERTO") Then
    '    MsgBox "Seu voto foi computado com sucesso, obrigado por participar", vbInformation, "Pronto"
    '    'End
    '    Else
    '    MsgBox "Seu voto não pode ser computado, ou você já votou ou ocorreu um erro na conexão com o servidor, tente novamente mais tarde", vbCritical, "Erro"
    '   ' End
    '    End If
    '
    '
    'End If
    '
    'Exit Sub
    'err:
    'Exit Sub
End Sub

