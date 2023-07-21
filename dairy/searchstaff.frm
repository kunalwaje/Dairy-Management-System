VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H80000009&
   Caption         =   "Form9"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "Sitka Display"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form9"
   ScaleHeight     =   7830
   ScaleWidth      =   14535
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "SEARCH STAFF"
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
      Height          =   8535
      Left            =   2280
      TabIndex        =   0
      Top             =   2400
      Width           =   15615
      Begin VB.TextBox txtid 
         DataField       =   "STAFF ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   9
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtpass 
         DataField       =   "PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   8
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtgen 
         DataField       =   "GENDER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   3480
         TabIndex        =   7
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox txtsal 
         DataField       =   "SALARY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   3480
         TabIndex        =   6
         Top             =   5160
         Width           =   1815
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   4680
         TabIndex        =   5
         Top             =   7320
         Width           =   2055
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   8520
         TabIndex        =   4
         Top             =   7320
         Width           =   2055
      End
      Begin VB.TextBox txtname 
         DataField       =   "NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10440
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtadd 
         DataField       =   "ADDRESS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10440
         TabIndex        =   2
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtmob 
         DataField       =   "CONTACT NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   10440
         TabIndex        =   1
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label staffid 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "STAFF ID"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   840
         TabIndex        =   16
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label password 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   960
         TabIndex        =   15
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label gender 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GENDER"
         BeginProperty Font 
            Name            =   "Sitka Heading"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   960
         TabIndex        =   14
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label salary 
         BackStyle       =   0  'Transparent
         Caption         =   "SALARY"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1200
         TabIndex        =   13
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label name 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   7440
         TabIndex        =   12
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label address 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   7560
         TabIndex        =   11
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label contact 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   7560
         TabIndex        =   10
         Top             =   3960
         Width           =   2295
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "CURTIN CREAMERY"
      BeginProperty Font 
         Name            =   "Sitka Display"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   6840
      TabIndex        =   17
      Top             =   480
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   0
      Picture         =   "searchstaff.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSearch_Click()

If txtid.Text <> "" Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from STAFF where staffid=" & CInt(txtid.Text) & "", con, adOpenDynamic, adLockOptimistic
    If Not rs.EOF And Not rs.BOF Then
        Call loaddata
        
    Else
        MsgBox "This is new Staff"
        Module1.retdata ("select max(staffid)+1 from staff")
        txtid.Text = rs.Fields(0).Value
        
        
        txtname.Text = ""
        txtmob.Text = ""
        txtadd.Text = ""
        txtgen.Text = ""
        txtpass.Text = ""
        txtsal.Text = ""
        
        
        End If
 Else
 MsgBox ("Enter id number to search")
 txtid.SetFocus
 
 End If

End Sub

Private Sub loaddata()

txtid.Text = rs.Fields(0).Value
txtname.Text = rs.Fields(1).Value
txtpass.Text = rs.Fields(2).Value
txtadd.Text = rs.Fields(3).Value
txtgen.Text = rs.Fields(4).Value
txtmob.Text = rs.Fields(5).Value
txtsal.Text = rs.Fields(6).Value
End Sub

Private Sub CmdUpdate_Click()

If txtid.Text <> "" And txtname.Text <> "" And txtpass.Text <> "" And txtadd.Text <> "" And txtmob.Text <> "" And txtgen.Text <> "" And txtsal.Text <> "" Then
Module1.retdata ("update staff set staffid='" & txtid.Text & "', name = '" & txtname.Text & "', pass = '" & txtpass.Text & "', address='" & txtadd.Text & "', gender='" & txtgen.Text & "', contact_no='" & txtmob.Text & "' , salary='" & txtsal.Text & "' where staffid = " & CInt(txtid.Text) & "")
MsgBox ("Data Updated")
Module1.retdata ("select * from staff")
Call loaddata
Else
MsgBox ("Fill All The Fields")
End If

End Sub
