VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Register details"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5430
   FillColor       =   &H000080FF&
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
   ScaleHeight     =   5910
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Отказ"
      Height          =   360
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   1350
   End
   Begin VB.Frame frameRegistry 
      Appearance      =   0  'Flat
      Caption         =   "Registry administrator signature"
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   5175
      Begin Inspector.TextBoxEx txtRegistryDate 
         Height          =   510
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   900
         BackColor       =   -2147483633
         ForeColor       =   0
         Caption         =   "Seal date"
         CaptionColor    =   0
         UnderlineColor  =   0
         Tips            =   "dd/mm/yyyy"
      End
      Begin Inspector.TextBoxEx txtRegistryAdmin 
         Height          =   510
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   900
         BackColor       =   -2147483633
         ForeColor       =   0
         Caption         =   ""
         CaptionColor    =   8388608
         UnderlineColor  =   0
         Tips            =   "Title and Surname"
      End
      Begin Inspector.TextBoxEx txtStickerRegistrty 
         Height          =   510
         Left            =   3600
         TabIndex        =   8
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   900
         BackColor       =   -2147483633
         ForeColor       =   0
         Caption         =   "Sticker number"
         CaptionColor    =   0
         UnderlineColor  =   0
         Required        =   -1  'True
      End
   End
   Begin VB.Frame frameAdmin 
      Appearance      =   0  'Flat
      Caption         =   "Network administrator signature"
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   5175
      Begin Inspector.TextBoxEx txtSecurityAdminName 
         Height          =   510
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   900
         BackColor       =   -2147483633
         ForeColor       =   0
         Caption         =   ""
         CaptionColor    =   0
         UnderlineColor  =   0
         Tips            =   "Title and Surname"
      End
      Begin Inspector.TextBoxEx txtStickerAdmin 
         Height          =   510
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   900
         BackColor       =   -2147483633
         ForeColor       =   0
         Caption         =   "Sticker number"
         CaptionColor    =   0
         UnderlineColor  =   0
         Required        =   -1  'True
      End
      Begin Inspector.TextBoxEx txtAdminDate 
         Height          =   510
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   900
         BackColor       =   -2147483633
         ForeColor       =   0
         Caption         =   "Seal date"
         CaptionColor    =   0
         UnderlineColor  =   0
         Tips            =   "dd/mm/yyyy"
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Запис"
      Default         =   -1  'True
      Height          =   360
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   1350
   End
   Begin Inspector.TextBoxEx txtInvNo 
      Height          =   510
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   900
      BackColor       =   -2147483633
      ForeColor       =   0
      Caption         =   "Inventary number"
      CaptionColor    =   0
      UnderlineColor  =   0
      Tips            =   "x-xxx"
      Required        =   -1  'True
   End
   Begin Inspector.TextBoxEx txtInvDate 
      Height          =   510
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   900
      BackColor       =   -2147483633
      ForeColor       =   0
      Caption         =   "Date"
      CaptionColor    =   0
      UnderlineColor  =   0
      Required        =   -1  'True
   End
   Begin Inspector.TextBoxEx txtSerialNumber 
      Height          =   510
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   840
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   900
      BackColor       =   -2147483633
      ForeColor       =   0
      Caption         =   "Hard disk drive serial number"
      CaptionColor    =   0
      Locked          =   -1  'True
      UnderlineColor  =   0
   End
   Begin Inspector.TextBoxEx txtBookSerialNumber 
      Height          =   510
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   900
      BackColor       =   -2147483633
      ForeColor       =   0
      Caption         =   "Hard disk drive serial number (if differs from the one above)"
      CaptionColor    =   0
      UnderlineColor  =   0
   End
   Begin Inspector.TextBoxEx txtRegistryNo 
      Height          =   510
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   900
      BackColor       =   -2147483633
      ForeColor       =   0
      Caption         =   "Registry number"
      CaptionColor    =   0
      UnderlineColor  =   0
      Tips            =   "RBxxxxxxxxxxxxxx"
      Required        =   -1  'True
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    
    If Not ValidateRegForm Then Exit Sub
        
    With HW_HARDDISKS(RegisterHDDIndex)
    
        .Registry.RegistryNo = txtRegistryNo.Text
        .Registry.InventaryNo = txtInvNo.Text
        .Registry.InventaryDate = txtInvDate.Text
        .Registry.InventarySerialNum = txtBookSerialNumber.Text
        
        .Registry.AdminName = txtSecurityAdminName.Text
        .Registry.AdminSticker = txtStickerAdmin.Text
        .Registry.AdminDate = txtAdminDate.Text
        
        .Registry.RegistryName = txtRegistryAdmin.Text
        .Registry.RegistrySticker = txtStickerRegistrty.Text
        .Registry.RegistryDate = txtRegistryDate.Text
        
    End With
    
    Unload Me
End Sub

Private Function ValidateRegForm() As Boolean
    
    Dim bUncls As Boolean
        
    ValidateRegForm = False
    bUncls = REG_WORKSTATION.Classification <> "unclassified"
    
    If Not NotEmpty(txtRegistryNo.Text) And bUncls Then
        MsgBox "Please enter number of Registry for classified information.", _
            vbExclamation, "Error"
        txtRegistryNo.SetFocus
    ElseIf Not NotEmpty(txtInvNo.Text) Then
        MsgBox "Please enter inventory number of hard disk drive.", vbExclamation, "Error"
        txtInvNo.SetFocus
    ElseIf Not NotEmpty(txtInvDate.Text) Then
        MsgBox "Please enter inventory date of hard disk drive.", vbExclamation, "Error"
        txtInvDate.SetFocus
    ElseIf Not NotEmpty(txtSerialNumber.Text) And Not NotEmpty(txtBookSerialNumber.Text) Then
        MsgBox "Please enter hard disk drive seria number", _
            vbExclamation, "Error"
        txtBookSerialNumber.SetFocus
    ElseIf Not NotEmpty(txtStickerAdmin.Text) And bUncls Then
        MsgBox "Please enter sticker number.", _
            vbExclamation, "Error"
        txtStickerAdmin.SetFocus
    ElseIf Not NotEmpty(txtStickerRegistrty.Text) And bUncls Then
        MsgBox "Please enter sticker number.", _
            vbExclamation, "Error"
        txtStickerRegistrty.SetFocus
    Else
            ValidateRegForm = True
    End If
    
End Function

Private Sub Form_Load()
    
    With HW_HARDDISKS(RegisterHDDIndex)

        Me.Caption = .Model & " (" & .Size & " GB)"
        
        txtSerialNumber.Text = .SerialNumber
        
        If .SerialNumber = vbNullString Then
            txtBookSerialNumber.Required = True
        End If
        
        If .Registry.RegistryNo <> vbNullString Then _
            txtRegistryNo.Text = .Registry.RegistryNo
        If .Registry.InventaryNo <> vbNullString Then _
            txtInvNo.Text = .Registry.InventaryNo
        txtInvDate.Text = .Registry.InventaryDate
        txtBookSerialNumber.Text = .Registry.InventarySerialNum
        
        If .Registry.AdminName <> vbNullString Then _
            txtSecurityAdminName.Text = .Registry.AdminName
        txtStickerAdmin.Text = .Registry.AdminSticker
        txtAdminDate.Text = .Registry.AdminDate

        If .Registry.RegistryName <> vbNullString Then _
            txtRegistryAdmin.Text = .Registry.RegistryName
        txtStickerRegistrty.Text = .Registry.RegistrySticker
        txtRegistryDate.Text = .Registry.RegistryDate

    End With
    
    FormPrepare
End Sub

Private Sub txtBookSerialNumber_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtInvNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtRegistryNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub FormPrepare()
    Select Case REG_WORKSTATION.Classification
    
        Case "unclassified"
            txtRegistryNo.Enabled = False
            txtStickerAdmin.Enabled = False
            txtStickerRegistrty.Enabled = False
            txtAdminDate.Enabled = False
            txtRegistryDate.Enabled = False
            
            frameAdmin.Enabled = False
            frameRegistry.Enabled = False
        Case Else
            txtRegistryNo.Enabled = True
            txtStickerAdmin.Enabled = True
            txtStickerRegistrty.Enabled = True
            txtAdminDate.Enabled = True
            txtRegistryDate.Enabled = True
            
            frameAdmin.Enabled = True
            frameRegistry.Enabled = True
    End Select
End Sub
