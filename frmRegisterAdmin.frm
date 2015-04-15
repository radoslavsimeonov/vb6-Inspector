VERSION 5.00
Begin VB.Form frmRegisterUser 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Отказ"
      Height          =   360
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2160
      Width           =   1110
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Запис"
      Default         =   -1  'True
      Height          =   360
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2160
      Width           =   1110
   End
   Begin VB.CheckBox chkNAVY 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "NAVY"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin Inspector.TextBoxEx txtPhone 
      Height          =   510
      Left            =   3240
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   900
      Caption         =   "Телефон"
   End
   Begin Inspector.TextBoxEx txtBuilding 
      Height          =   510
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   900
      Caption         =   "Сграда"
   End
   Begin Inspector.TextBoxEx txtFunction 
      Height          =   510
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   900
      Caption         =   "Длъжност"
      Tips            =   "длъжност във в.ф./дирекция/звено"
   End
   Begin Inspector.TextBoxEx txtFirstName 
      Height          =   510
      Left            =   2880
      TabIndex        =   4
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   900
      Caption         =   "Име"
   End
   Begin Inspector.TextBoxEx txtMiddleName 
      Height          =   510
      Left            =   5040
      TabIndex        =   5
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   900
      Caption         =   "Презиме"
   End
   Begin Inspector.TextBoxEx txtSurName 
      Height          =   510
      Left            =   7200
      TabIndex        =   6
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   900
      Caption         =   "Фамилия"
   End
   Begin Inspector.TextBoxEx txtRoom 
      Height          =   510
      Left            =   2280
      TabIndex        =   7
      Top             =   2040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   900
      Caption         =   "Стая"
   End
   Begin Inspector.TextBoxEx cmbRanks 
      Height          =   510
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   900
      Caption         =   "Звание"
      MaxRowsToDisplay=   21
      ControlType     =   1
      ComboStyle      =   1
   End
   Begin Inspector.TextBoxEx txtUserName 
      Height          =   510
      Left            =   7080
      TabIndex        =   9
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   900
      Caption         =   "Потребителско име"
   End
   Begin Inspector.TextBoxEx txtDepartment 
      Height          =   510
      Left            =   4320
      TabIndex        =   10
      Top             =   1200
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   900
      Caption         =   "Дирекция / отдел / звено"
   End
End
Attribute VB_Name = "frmRegisterUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RankLabels() As Elements

Private Sub chkNAVY_Click()
    If chkNAVY.Value = 0 Then
        chkNAVY.ForeColor = &HC0C0C0
        SwitchForcesRanks 0, cmbRanks
    Else
        chkNAVY.ForeColor = &HFF00&
        SwitchForcesRanks 1, cmbRanks
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    RankLabels = Fill_Rank(cmbRanks)
    cmbRanks.AutoResizeColumns
End Sub

Private Sub txtBuilding_KeyPress(KeyAscii As Integer)
    If Len(txtBuilding.Text) = 0 Or txtBuilding.SelLength = Len(txtBuilding.Text) Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
End Sub

Private Sub txtDepartment_KeyPress(KeyAscii As Integer)
    If Len(txtDepartment.Text) = 0 Or txtDepartment.SelLength = Len(txtDepartment.Text) Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
    If Len(txtFirstName.Text) = 0 Or txtFirstName.SelLength = Len(txtFirstName.Text) Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
End Sub

Private Sub txtFunction_KeyPress(KeyAscii As Integer)
    If Len(txtFunction.Text) = 0 Or txtFunction.SelLength = Len(txtFunction.Text) Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
End Sub

Private Sub txtMiddleName_KeyPress(KeyAscii As Integer)
    If Len(txtMiddleName.Text) = 0 Or txtMiddleName.SelLength = Len(txtMiddleName.Text) Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
End Sub

Private Sub txtSocketSKS_LostFocus()
    If NotEmpty(txtSocketSKS.Text) Then _
        REG_WORKSTATION.SocketSKS = txtSocketSKS.Text
End Sub

Private Sub txtSurName_KeyPress(KeyAscii As Integer)
    If Len(txtSurName.Text) = 0 Or txtSurName.SelLength = Len(txtSurName.Text) Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
End Sub

