VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H8000000B&
   Caption         =   "Form11"
   ClientHeight    =   9630
   ClientLeft      =   -360
   ClientTop       =   465
   ClientWidth     =   15855
   LinkTopic       =   "Form11"
   ScaleHeight     =   9630
   ScaleWidth      =   15855
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "INVOICE"
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
      Height          =   10455
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   15615
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   480
         TabIndex        =   27
         Text            =   "PRODUCTS"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox txtQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         MaxLength       =   10
         TabIndex        =   26
         Top             =   4440
         Width           =   2055
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add to bill"
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   25
         Top             =   6240
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Sitka Display"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   24
         Top             =   7440
         Width           =   3615
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   9375
         Left            =   7680
         TabIndex        =   1
         Top             =   600
         Width           =   6255
         Begin VB.Label lblcname 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-  0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   1800
            TabIndex        =   30
            Top             =   2640
            Width           =   330
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Bill"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2400
            TabIndex        =   23
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblGrand 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   4680
            TabIndex        =   22
            Top             =   8880
            Width           =   1575
         End
         Begin VB.Label lblGST 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   4560
            TabIndex        =   21
            Top             =   8400
            Width           =   1695
         End
         Begin VB.Label lblsubtotal 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   4560
            TabIndex        =   20
            Top             =   7920
            Width           =   1695
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "SIGN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   600
            TabIndex        =   19
            Top             =   8760
            Width           =   1575
         End
         Begin VB.Label lblGT 
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   3000
            TabIndex        =   18
            Top             =   8880
            Width           =   1575
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "GST (5%)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   3000
            TabIndex        =   17
            Top             =   8400
            Width           =   1695
         End
         Begin VB.Label lblTlt 
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   3000
            TabIndex        =   16
            Top             =   7920
            Width           =   1695
         End
         Begin VB.Line Line3 
            BorderWidth     =   2
            X1              =   120
            X2              =   5880
            Y1              =   7800
            Y2              =   7800
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   4560
            TabIndex        =   15
            Top             =   4320
            Width           =   735
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   2880
            TabIndex        =   14
            Top             =   4320
            Width           =   1095
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Price"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   13
            Top             =   4320
            Width           =   735
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Item Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   4320
            Width           =   1695
         End
         Begin VB.Label lblmobile 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- 0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   1800
            TabIndex        =   11
            Top             =   3360
            Width           =   270
         End
         Begin VB.Label lblDate 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   4680
            TabIndex        =   10
            Top             =   3240
            Width           =   1335
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblbno 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   5160
            TabIndex        =   9
            Top             =   2640
            Width           =   975
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   240
            X2              =   6000
            Y1              =   4200
            Y2              =   4200
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Bill No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   3600
            TabIndex        =   6
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   3600
            TabIndex        =   5
            Top             =   3240
            Width           =   735
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   240
            X2              =   6000
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Address : Shop No. 12 , BKC Complex,Prabhat road, Pune-411-004"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   735
            Left            =   240
            TabIndex        =   4
            Top             =   1200
            Width           =   5535
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No. - 020-42042020   |   Mobile No.- 9922033913   |   GSTIN:- INDIA020123"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   735
            Left            =   240
            TabIndex        =   3
            Top             =   1920
            Width           =   6495
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Curtin Creamery"
            BeginProperty Font 
               Name            =   "Sitka Heading"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   735
            Left            =   1560
            TabIndex        =   2
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Item Name"
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
         Height          =   615
         Left            =   480
         TabIndex        =   29
         Top             =   1920
         Width           =   3735
         WordWrap        =   -1  'True
      End
      Begin VB.Label LblQty 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Quantity"
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
         Height          =   735
         Index           =   1
         Left            =   600
         TabIndex        =   28
         Top             =   3600
         Width           =   2895
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   0
      Picture         =   "Bill.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wg As String
Dim qty As Integer
Dim tqty As Integer
Dim ccnt As Integer 'counter for control


Private Sub cmdAdd_Click()
If txtQty.Text <> "" Then
ccnt = ccnt + 1
Module1.retdata ("select * from product_master where pname='" & Combo1.Text & "'")
wg = rs.Fields(4).Value
'Module1.retdata ("insert into bill values(" & CInt(lblbno.Caption) & ",'" & Combo1.Text & "'," & CInt(txtQty.Text) & ")")

If rs.State = 1 Then rs.Close
rs.Open
If qty < txtQty.Text Then
MsgBox ("not enough quantity")
ccnt = ccnt - 1

ElseIf Not rs.EOF Then

Load Label10(ccnt)
With Label10(ccnt)
.Caption = rs.Fields(1).Value
.Visible = True
.Top = Label10(ccnt - 1).Top + Label10(ccnt - 1).Height + 15
End With
Load Label11(ccnt)
With Label11(ccnt)
.Caption = rs.Fields(2)
.Visible = True
.Top = Label11(ccnt - 1).Top + Label11(ccnt - 1).Height + 15
End With
Load Label19(ccnt)
With Label19(ccnt)
.Caption = txtQty.Text
.Visible = True
.Top = Label19(ccnt - 1).Top + Label19(ccnt - 1).Height + 15
End With

Load Label13(ccnt)
With Label13(ccnt)
.Caption = CInt(rs.Fields(2).Value) * CInt(txtQty.Text)
.Visible = True
.Top = Label10(ccnt - 1).Top + Label10(ccnt - 1).Height + 15
End With
lblsubtotal.Caption = CInt(lblsubtotal.Caption) + CInt(Label13(ccnt).Caption)
lblGST.Caption = CInt(lblsubtotal.Caption) * 0.18
lblGrand.Caption = CInt(lblsubtotal.Caption) + CInt(lblGST.Caption)

tqty = qty - CInt(txtQty.Text)
Module1.retdata ("insert into bill values(" & CInt(lblbno.Caption) & ",'" & Combo1.Text & "'," & CInt(txtQty.Text) & ")")
Module1.retdata ("update product_master set pqty==" & tqty & " where pname = '" & Combo1.Text & "'")
MsgBox ("Data Updated")

End If
Else
MsgBox ("Enter Quantity")
End If

End Sub

Private Sub Combo1_click()
Module1.retdata ("select pqty from product_master where pname='" & Combo1.Text & "'")
MsgBox ("Available Quantity is " & rs.Fields(0).Value)
qty = rs.Fields(0).Value
'If qty < txtQty.Text Then
'MsgBox ("not enough quantity")
'End If

End Sub


Private Sub Command1_Click()

Module1.retdata ("insert into mainbill values(" & CInt(lblbno.Caption) & ",'" & lblmobile.Caption & "','" & lblsubtotal.Caption & "','" & lblGST.Caption & "','" & lblDate.Caption & "','" & lblGrand.Caption & "')")


PrintForm

End Sub

Private Sub Form_Load()
lblDate.Caption = Format(Date, "dd/MM/yyyy")

Module1.numquery ("select max(bid)+1 from bill")
lblbno.Caption = nm

Module1.retdata ("select distinct(pname) from product_master")
While Not rs.EOF
Combo1.AddItem (rs.Fields(0).Value)
rs.MoveNext
Wend

End Sub


