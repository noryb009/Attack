VERSION 5.00
Begin VB.Form frmSTORE 
   Caption         =   "Store"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBACK 
      Caption         =   "Back"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "frmSTORE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Attack
' Luke Lorimer
' 21 November, 2011
' Defend your castle!

Private Sub cmdBACK_Click()
    frmLEVELSELECT.Show
    Unload frmSTORE
End Sub

Private Sub Form_Load()

End Sub
