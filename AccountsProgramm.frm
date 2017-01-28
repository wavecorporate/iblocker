VERSION 5.00
Begin VB.Form AccountsProgramm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Senha para acessar ao Programa"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AccountsProgramm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "AccountsProgramm.frx":030A
   MousePointer    =   99  'Custom
   Picture         =   "AccountsProgramm.frx":0BD4
   ScaleHeight     =   4155
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   12
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "w"
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2880
      MaxLength       =   12
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "w"
      TabIndex        =   0
      Top             =   1350
      Width           =   2535
   End
   Begin VB.CommandButton btnAlterar 
      Caption         =   "Alterar"
      Height          =   285
      Left            =   6720
      TabIndex        =   13
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton btnSaveMaster 
      Caption         =   "Salvar"
      Height          =   285
      Left            =   6720
      TabIndex        =   12
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton btnNovaSenhaMaster 
      Caption         =   "..."
      Height          =   285
      Left            =   10200
      TabIndex        =   11
      Top             =   6960
      Width           =   615
   End
   Begin VB.CommandButton Cmdsalvar 
      Caption         =   "&Salvar"
      Default         =   -1  'True
      Height          =   645
      Left            =   4440
      TabIndex        =   10
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton CmdClose 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   405
      Left            =   2640
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdalterar 
      Caption         =   "Alterar"
      Height          =   405
      Left            =   840
      Picture         =   "AccountsProgramm.frx":1C0886
      TabIndex        =   8
      Top             =   3120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3600
      ScaleHeight     =   1185
      ScaleWidth      =   4665
      TabIndex        =   2
      Top             =   6000
      Width           =   4695
      Begin VB.TextBox TxtDefault 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alterar senha Master do programa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   3165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2700
      End
   End
   Begin VB.Label LblView 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   450
      Left            =   720
      TabIndex        =   17
      Top             =   2760
      Width           =   4380
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha Administrador"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   480
      TabIndex        =   16
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirme a Senha:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   480
      TabIndex        =   15
      Top             =   1920
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atenção, não perca a Senha. Sem ela não poderá acessar ao Programa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   14
      Top             =   2400
      Width           =   5175
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha Master - "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8760
      TabIndex        =   7
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"AccountsProgramm.frx":380538
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   3720
      TabIndex        =   6
      Top             =   5040
      Width           =   4620
   End
End
Attribute VB_Name = "AccountsProgramm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Module    : AccountsProgramm
' Author    : Administrador Vagner Costa
' Empresa   : Wave Company Software
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Dim X           As Integer

Dim Modoadd     As Byte

Dim Modoedit    As Boolean

Dim strEditRec  As String

Dim Getpassword As String

Dim howUser     As Boolean

Dim controle    As Boolean

Private Sub cmdNav_Click(Index As Integer)

    On Error Resume Next

    Unload Me
End Sub

Private Sub btnAlterar_Click()
    TxtDefault.Locked = False

    TxtDefault.SelStart = 0
    TxtDefault.SelLength = Len(TxtDefault.Text)
    TxtDefault.SetFocus
    controle = True
End Sub

Private Sub btnNovaSenhaMaster_Click()

    If AccountsProgramm.Height = 5985 Then Me.Height = 3600 Else: AccountsProgramm.Height = 5985

End Sub

Private Sub btnSaveMaster_Click()

    If controle Then
        If TxtDefault.Text <> "" Then
            SaveSettingString HKEY_CLASSES_ROOT, "Microsoft\Certifield\ENVIRONMENT", "Passwd", Cripto(TxtDefault.Text)
            TxtDefault.Locked = True
            MsgBox "A Password: " & TxtDefault.Text & " foi salva com sucesso!" & vbCrLf & vbCrLf & vbCrLf, vbInformation
            TxtDefault.Locked = True
            controle = False
        Else
            MsgBox "Por favor, digite a nova Password master do programa!" & vbCrLf & vbCrLf & vbCrLf, vbExclamation

        End If

    End If

End Sub

Private Sub cmdalterar_Click()

    On Error Resume Next

    Modoadd = 1
    UnLockEm
    'Text1.Enabled = False
    strEditRec = Text2
    SetButtons False

    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)

    If Getpassword <> "" Then Text2.SetFocus
End Sub

Private Sub CmdClose_Click()

    On Error Resume Next

    If Modoadd = 1 Then MsgBox "Por favor salve antes de sair!" & vbCrLf & vbCrLf & vbCrLf & vbCrLf, vbExclamation: Exit Sub

    Unload Me
End Sub

Private Sub Cmdsalvar_Click()

    On Error GoTo Handler

    If ValidData Then
        If ValidX Then
            If v_Senha Then
                aC

                If Modoadd = 2 Then MsgBox "Password criada com sucesso!" & vbCrLf & vbCrLf & vbCrLf & vbCrLf, vbInformation, "Wave iBlocker": GoTo w1
                If Modoadd = 1 Then MsgBox "Alterado com sucesso!" & vbCrLf & vbCrLf & vbCrLf & vbCrLf, vbInformation, "Wave iBlocker": Unload Me
                Modoadd = 4

                SetButtons True

                'Text1.SetFocus
                LockEm
                'ClEar
                cmdalterar.Enabled = True

            End If
        End If
    End If

    Exit Sub

w1:

    Unload Me
    
    mINI.Load_BDNET

    FrmNetController.Show
    FrmNetController.ImgStatus.Picture = FrmNetController.pic_ProtectOFF.Picture

    Exit Sub

Handler:
    Call ErroS
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys ("{tab}")
        Call Cmdsalvar_Click
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()

    On Error Resume Next

    Screen.MousePointer = vbDefault

    SetButtons True
    Modoadd = False

    
    Getpassword = Cripto(GetSettingString(HKEY_CLASSES_ROOT, "Microsoft\uname", "a"))
    If Getpassword = "" Then

        'SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "FolderURL", fso.GetFolder(APP.path)
        '                 HKEY_CURRENT_USER, "Shell.Application\                                    CurVer\ENVIRONMENT", "PassMonitor", cripto(Text1.text)
        MsgBox "Importante: Crie agora uma senha para o Wave iBlocker" & vbCrLf & vbCrLf & vbCrLf & "Pressione OK para continuar." & vbCrLf & vbCrLf & vbCrLf & vbCrLf, vbExclamation, "Crie uma Password"
        Call cmdalterar_Click: Modoadd = 2:  CmdClose.Enabled = False
    Else
        LockEm

        exibeX
    End If

End Sub

Private Sub exibeX()

    On Error Resume Next

    Text2.Text = Getpassword
    Text3.Text = Text2.Text
    LblView.Caption = "A sua Password é: " & Getpassword
End Sub

Private Sub LockEm()
    Text2.Locked = True
    Text3.Locked = True
End Sub

Private Sub UnLockEm()
    Text2.Locked = False
    Text3.Locked = False
End Sub

Private Sub aC()

    On Error Resume Next


     
    SaveSettingString HKEY_CLASSES_ROOT, "Microsoft\uname", "a", Cripto(Text2.Text)
    


    '-------user
End Sub

Private Sub ClEar()

    On Error Resume Next

    Text2.Text = ""
    Text3 = ""
End Sub

Private Function ValidX() As Boolean

    '-------compare--------------
    If Text3.Text <> Text2.Text Then
        MsgBox "A Confirmação de Password está errada!" & vbCrLf & vbCrLf & vbCrLf & vbCrLf, vbCritical, "Erro de Password"
        Text3.SelStart = 0
        Text3.SelLength = Len(Text3.Text)
        Text3.SetFocus

    Else
        ValidX = True

    End If

End Function

Private Function ValidData() As Boolean

    On Error Resume Next

    If Text2 = "" Then
        Text2.SetFocus
        MsgBox "Digite a Password." & vbCrLf & vbCrLf & vbCrLf & vbCrLf, vbCritical, "Wave iBlocker"
    ElseIf Text3 = "" Then
        Text3.SetFocus
        MsgBox "Digite a confirmação da Password." & vbCrLf & vbCrLf & vbCrLf & vbCrLf, vbCritical, "Wave iBlocker"
    ElseIf Text3 = "" Then

    Else
        ValidData = True
    End If

End Function

Public Sub SetButtons(bVal As Boolean)

    On Error Resume Next

    cmdalterar.Enabled = bVal
    Cmdsalvar.Enabled = Not bVal

    CmdClose.Enabled = bVal

End Sub

Private Sub Image3_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Text2_Click()

    If Modoadd = 1 Then Text3 = ""

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    If Modoadd = 1 Then Text3 = ""

End Sub

Private Function v_Senha() As Boolean

    On Error Resume Next

    Dim i As Integer

    i = Len(Text2.Text)

    If i < 4 Then
        MsgBox "Por favor, o campo Password, deve ter no mínimo 4 caracteres" & vbCrLf & vbCrLf & vbCrLf & vbCrLf, vbCritical, "Erro de Password"
        Text2.SelStart = 0
        Text2.SelLength = Len(Text2.Text)
        Text2.SetFocus
    Else
        v_Senha = True
    End If

End Function
