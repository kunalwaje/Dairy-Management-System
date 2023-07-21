VERSION 5.00
Begin VB.Form RECEIVE_OREDR 
   Caption         =   "Form5"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15030
   LinkTopic       =   "Form5"
   ScaleHeight     =   8745
   ScaleWidth      =   15030
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "RECEIVE ORDER"
      BeginProperty Font 
         Name            =   "Sitka Display"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   3360
      TabIndex        =   0
      Top             =   3120
      Width           =   13575
      Begin VB.ComboBox cmdOid 
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1200
         TabIndex        =   3
         Text            =   "Select Order Number"
         Top             =   1440
         Width           =   3975
      End
      Begin VB.TextBox txtqty 
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   7200
         TabIndex        =   2
         Top             =   3000
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Receive"
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   4440
         TabIndex        =   1
         Top             =   5040
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Product ID"
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   6960
         TabIndex        =   6
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label lblpid 
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11160
         TabIndex        =   5
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Order Quantity"
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   1320
         TabIndex        =   4
         Top             =   3120
         Width           =   3495
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "CURTIN CREAMERY"
      BeginProperty Font 
         Name            =   "Sitka Display"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6480
      TabIndex        =   7
      Top             =   360
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   0
      Picture         =   "RECEIVE_OREDR.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
End
Attribute VB_Name = "RECEIVE_OREDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOid_Click()
Module1.retdata ("select qty,pid from PURCHASE_DETAILS where oid=" & CInt(cmdOid.Text) & "")
If Not rs.EOF Then
lblpid.Caption = rs.Fields(1).Value
txtqty.Text = rs.Fields(0).Value
txtqty.Enabled = False

End If

End Sub

Private Sub Command1_Click()
Module1.retdata ("update PURCHASE_DETAILS set ostatus='Delivered' where oid=" & CInt(cmdOid.Text) & "")
Module1.retdata ("update product_master set pqty=pqty+" & CInt(txtqty.Text) & " where pid=" & CInt(lblpid.Caption) & "")

MsgBox ("Order Received")
cmdOid.CLEAR

Module1.retdata ("select oid from PURCHASE_DETAILS where ostatus='Undelivered'")
While Not rs.EOF
cmdOid.AddItem (rs.Fields(0).Value)
rs.MoveNext
Wend


End Sub

Private Sub Form_Load()
Module1.retdata ("select oid from PURCHASE_DETAILS where ostatus='Undelivered'")
While Not rs.EOF
cmdOid.AddItem (rs.Fields(0).Value)
rs.MoveNext
Wend

End Sub


