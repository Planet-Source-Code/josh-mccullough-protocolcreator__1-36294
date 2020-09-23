VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ProtocolCreator"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRun 
      Caption         =   "&Run registry file after creation"
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Height          =   285
      Left            =   5760
      TabIndex        =   7
      ToolTipText     =   "Create the REG file with the selected options."
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   285
      Left            =   5760
      TabIndex        =   6
      ToolTipText     =   "Find the executable."
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1080
      MaxLength       =   255
      TabIndex        =   5
      ToolTipText     =   "Path that the protocol will launch. Example: C:\Program Files\myApp"
      Top             =   840
      Width           =   4575
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1080
      MaxLength       =   255
      TabIndex        =   3
      Text            =   "ShareThere Client Launcher"
      ToolTipText     =   "Describe what this protocol will do."
      Top             =   480
      Width           =   5775
   End
   Begin VB.TextBox txtProtocol 
      Height          =   285
      Left            =   1080
      MaxLength       =   255
      TabIndex        =   1
      Text            =   "ShareThere"
      ToolTipText     =   "Example: http, ftp, myApp."
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label lblFooter 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "ProtocolCreator v1.0 (C) JSM Enterprises.  http://jsment.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   -90
      TabIndex        =   8
      Top             =   1680
      Width           =   7155
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   6840
      X2              =   120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "P&ath:"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   885
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Description:"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   525
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Protocol:"
      Height          =   195
      Left            =   345
      TabIndex        =   0
      Top             =   165
      Width           =   630
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    Dialog.ShowOpen
    Dialog.CancelError = False
    txtPath = Dialog.FileName
End Sub

Private Sub cmdGo_Click()
    If Trim(txtProtocol) = "" Then
        MsgBox "Please enter a Protocol.", vbCritical, "Human error"
        txtProtocol.SetFocus
    ElseIf Trim(txtPath) = "" Then
        MsgBox "Please enter a Path.", vbCritical, "Human error"
    Else
        MsgBox GetProtocolCreationError(CreateProtocol(txtProtocol, txtDescription, txtPath, IIf(chkRun = vbChecked, True, False), False, Me))
    End If
End Sub
