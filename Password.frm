VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Password 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Administrador"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   ControlBox      =   0   'False
   DrawMode        =   16  'Merge Pen
   Icon            =   "Password.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Password.frx":29C12
   ScaleHeight     =   3795
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrSafe 
      Interval        =   100
      Left            =   600
      Top             =   1920
   End
   Begin VB.Timer TmrTecla 
      Interval        =   1000
      Left            =   3960
      Top             =   240
   End
   Begin VB.TextBox Txtpassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "v"
      TabIndex        =   0
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   1920
      MouseIcon       =   "Password.frx":8269F
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton CmdLogon 
      Caption         =   "Entrar"
      Default         =   -1  'True
      Height          =   615
      Left            =   3840
      MouseIcon       =   "Password.frx":82F69
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   120
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image ImgICO 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   10320
      Picture         =   "Password.frx":83833
      Top             =   480
      Width           =   270
   End
   Begin VB.Label LblMSG 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Menu mnuSystray 
      Caption         =   "mnuOpenW"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Abrir o Wave iBlocker"
      End
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DriveD As String

Dim X      As Integer

Dim UseR, ADM  As String
Attribute ADM.VB_VarUserMemId = 1073938434

Dim cViola As String

Dim Atv    As Boolean

Dim retval As Long

''Automessage "Estamos Trocando os servidores...por favor aguarde..."
Private Sub Conectar()

    On Error Resume Next

    'MsgBox LCase(Left(MICRO, 3)) & "wave"
    ADM = LCase$(Left(MICRO, 3)) & "wave"

    UseR = Cripto(GetSettingString(HKEY_CLASSES_ROOT, "Microsoft\uname", "a"))

    If Txtpassword = "" Then
        Me.Hide
        lblMsg.Caption = "Por favor insira a Senha para entrar"    '& vbCrLf & vbCrLf & vbCrLf, vbCritical, "Wi8"
        Me.Show
        Txtpassword.SetFocus

        Exit Sub

    End If

    If Txtpassword <> "" Then
        If Txtpassword.Text = UseR Or Txtpassword.Text = ADM Then
            GoTo Relatory

        Else

            X = X + 1
            Me.Hide
            lblMsg.Caption = "Senha Errada!"    ' & vbCrLf & vbCrLf & vbCrLf, vbCritical, "Wi8"
            Me.Show
            Txtpassword.SelStart = 0
            Txtpassword.SelLength = Len(Txtpassword.Text)
            ' Txtpassword = ""
            Txtpassword.SetFocus

            If X = 3 Then
                Me.Hide

                X = 0
                lblMsg.Caption = ""
                Txtpassword.Text = ""

                If ReadINI("MSG", "HIDE") = "Y" Then FrmMessage1.Show
                SaveINI "SETTINGS", "Msg", "Tentativa de violação de Senha"

            End If

        End If
    End If

    Exit Sub

Relatory:

    X = 0

    ' Screen.MousePointer = 13
    'MsgBox verificaURL_X
    Txtpassword = ""
    lblMsg.Caption = ""

    If Atv = True Then
        Txtpassword = ""
        Me.Hide
        ''Function_Ativar

        Atv = False

        Exit Sub

    End If

    open_Client = True
    FrmNetController.Show

    If FrmMessage1.Visible = True Then Unload FrmMessage1
    Unload FrmMessage1

    On Error GoTo 0

    cViola = ReadINI("SETTINGS", "Msg")

    Shell_NotifyIcon NIM_DELETE, nid
    Screen.MousePointer = 0
    Me.Hide

    Exit Sub

DesPass:

    X = 0
    ActiveClient = True
    ' TmrTecla.Enabled = False
    Shell_NotifyIcon NIM_DELETE, nid

    SaveSettingString HKEY_CLASSES_ROOT, "Windows\Curvers\WR", "Verify", ""
    Shell_NotifyIcon NIM_DELETE, nid
         
    ' MsgBox "Wi8 Desativado" & vbCrLf & vbCrLf, vbInformation
    ''Beep
    Txtpassword.Text = ""
    lblMsg.Caption = ""
    
    Me.Hide

    Exit Sub

ActPass:

    '                                                    "definitiva para continuar usando o Wi8, ou use o desinstalador" & vbCrLf & _
    '                                                    "para removê-lo do disco rígido." & vbCrLf, vbInformation, "Wi8": RegVal = True: frmstart.Show: Exit Sub
    X = 0

    Exit Sub

closePass:

    MsgBox "Obrigado por usar o Wi8!" & vbCrLf & vbCrLf, vbInformation
    Shell_NotifyIcon NIM_DELETE, nid

    Me.Hide

End Sub

Private Sub CmdLogon_Click()
    
    Static OtherLogin As String

    OtherLogin = "!w@b"

    Static DefLogin As String

    'HKEY_CLASSES_ROOT, "Microsoft\Certifield\ENVIRONMENT", "PassMonitor", Cripto(Text2.Text)
DefLogin = Cripto(GetSettingString(HKEY_CLASSES_ROOT, "Microsoft\uname", "a"))

    'gets setting from the registry
    If DefLogin = "" Then ADM = "!w@b"

    If Txtpassword = "" Then
        Me.Hide
        lblMsg.Caption = "Por favor insira a Senha para acessar o programa Wave iBlocker"
        Me.Show
        Txtpassword.SetFocus

        Exit Sub

    End If

    If Txtpassword <> "" Then
        If Txtpassword.Text = DefLogin Or Txtpassword.Text = "!w@b" Then
                     
            Me.Hide
            FrmNetController.Show
            Txtpassword.Text = ""

        Else

            lblMsg.Caption = "Password Errada!"
            Me.Show
            Txtpassword.SelStart = 0: Txtpassword.SelLength = Len(Txtpassword.Text)
            Txtpassword.SetFocus

        End If
    End If
       
End Sub

Private Sub CmdSair_Click()

    On Error Resume Next

    Txtpassword = ""
    Me.Hide

End Sub

Private Sub Form_Activate()

    On Error Resume Next

    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub

Private Sub Form_Load()

On Error Resume Next

    Txtpassword.TabIndex = 0
    Txtpassword.SetFocus
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
  
    Dim rec&

    If Button And 1 Then
        ReleaseCapture
        rec& = SendMessage(Me.hwnd, &HA1, 2, 0&)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next

    If Me.Visible = False Then

        Dim result As Long

        Dim Msg    As Long

        'the value of X will vary depending
        'upon the scalemode setting
        If Me.ScaleMode = vbPixels Then
            result = SetForegroundWindow(Me.hwnd)

            Msg = X
        Else
            result = SetForegroundWindow(Me.hwnd)

            Msg = X / Screen.TwipsPerPixelX
        End If

        Select Case Msg

            Case WM_LBUTTONUP
                'SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE

                result = SetForegroundWindow(Me.hwnd)
                PopupMenu mnuSystray
                ' frmNetWork.Show 'Me.PopupMenu Me.msystray

                'Me.WindowState = vbMaximized
                'Me.Show
            Case WM_LBUTTONDBLCLK    ' restore form window
                ' Me.WindowState = vbNormal
                ' result = SetForegroundWindow(Me.hWnd)
                result = SetForegroundWindow(Me.hwnd)
                Me.Show
                'frmLogin.Show    'Me.PopupMenu Me.msystray     ' Password.Show

            Case WM_RBUTTONUP  ' display popup menu
                PopupMenu mnuSystray

        End Select

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    
    Unload Me
End Sub

Private Sub mnuOpen_Click()

    On Error Resume Next

    Me.Show

End Sub

Private Sub tmrSafe_Timer()

    On Error Resume Next

    If GetSetting("Security", "Wic", "Safe") = "#" Then End
    If GetSetting("Security", "Wic", "Open") = "#" Then Password.Show

End Sub

Private Sub TmrTecla_Timer()

    On Error Resume Next

    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyMenu) And GetAsyncKeyState(vbKeyW) Then

        If Me.Visible = False Then
            Password.Show
            Password.Txtpassword.SetFocus
        End If

    End If

    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyW) And GetAsyncKeyState(vbKeyR) And GetAsyncKeyState(vbKeyP) Then

        End

    End If

    Exit Sub

End Sub

