VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   10815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16890
   LinkTopic       =   "Form5"
   ScaleHeight     =   10815
   ScaleWidth      =   16890
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "STOCK INFORMATION"
      BeginProperty Font 
         Name            =   "Sitka Display"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   3480
      TabIndex        =   1
      Top             =   3720
      Width           =   13095
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1560
         TabIndex        =   2
         Text            =   "Select Product Id"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Available Stock is"
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   1320
         TabIndex        =   5
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   570
         Left            =   7560
         TabIndex        =   4
         Top             =   2640
         Width           =   3435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   7560
         TabIndex        =   3
         Top             =   1080
         Width           =   3450
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "CURTIN CREAMERY"
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   4920
      TabIndex        =   0
      Top             =   240
      Width           =   9975
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   120
      Picture         =   "stockinfo.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_click()
If Combo1.Text <> "Select Product Id" Then
Module1.retdata ("select * from product_master where pid=" & CInt(Combo1.Text) & "")
If Not rs.EOF Then
Label4.Caption = rs.Fields(1).Value
Label3.Caption = rs.Fields(3).Value
Else
MsgBox ("No Such Product")
End If
Else
MsgBox ("Select product id")
End If

End Sub

Private Sub Form_Load()
Module1.retdata ("select pid from product_master")
While Not rs.EOF
Combo1.AddItem (rs.Fields(0).Value)
rs.MoveNext

Wend

End Sub


