VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Khaznadar"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1470
      Left            =   210
      TabIndex        =   4
      Top             =   1275
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "O&k"
         Height          =   450
         Left            =   2490
         TabIndex        =   6
         Top             =   930
         Width           =   1770
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "This is the Enhanced Version of K-ColorPicker so please if you found any bug mail me at: thekingofdeath@hotmail.com"
         Height          =   600
         Left            =   150
         TabIndex        =   5
         Top             =   270
         Width           =   3930
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H0040BDA3&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   4680
      TabIndex        =   1
      Top             =   2925
      Width           =   4680
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "©All Rights Reserved For Majed A.Khaznadar"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   225
         TabIndex        =   2
         Top             =   225
         Width           =   4230
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by: Majed A.Khaznadar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0040BDA3&
      Height          =   240
      Left            =   285
      TabIndex        =   3
      Top             =   930
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "Form1.frx":0000
      Top             =   210
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Khaznadar ColorPicker Ver1.1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A96001&
      Height          =   900
      Left            =   735
      TabIndex        =   0
      Top             =   105
      Width           =   3765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

