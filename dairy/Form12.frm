VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "AdminMDI2"
   ClientHeight    =   10260
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   18270
   LinkTopic       =   "Form12"
   ScaleHeight     =   10260
   ScaleWidth      =   18270
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   12135
      Left            =   0
      Picture         =   "Form12.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
   Begin VB.Menu staff 
      Caption         =   "STAFF"
      Begin VB.Menu addstaff 
         Caption         =   "ADD STAFF"
      End
      Begin VB.Menu srchstaff 
         Caption         =   "SEARCH STAFF"
      End
   End
   Begin VB.Menu stock 
      Caption         =   "STOCKS"
   End
   Begin VB.Menu reports 
      Caption         =   "REPORTS"
      Begin VB.Menu supplierreports 
         Caption         =   "SUPPLIER REPORTS"
      End
      Begin VB.Menu productreport 
         Caption         =   "PRODUCTS REPORTS"
      End
      Begin VB.Menu purchasereports 
         Caption         =   "PURCHASE REPORTS"
      End
      Begin VB.Menu cusreports 
         Caption         =   "CUSTOMER REPORTS"
      End
   End
   Begin VB.Menu exit 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addstaff_Click()
Form3.Show
End Sub

Private Sub cusreports_Click()
DataReport1.Show
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub productreport_Click()
DataReport2.Show
End Sub

Private Sub purchasereports_Click()
DataReport4.Show
End Sub

Private Sub srchstaff_Click()
Form9.Show
End Sub

Private Sub stock_Click()
Form5.Show
End Sub

Private Sub supplierreports_Click()
DataReport3.Show
End Sub
