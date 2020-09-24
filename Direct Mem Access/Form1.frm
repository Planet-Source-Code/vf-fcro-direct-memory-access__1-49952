VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Direct Memory Access By Vanja Fuckar,EMAIL:inga@vip.hr"
   ClientHeight    =   585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   585
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Direct Memory Access (R/W)"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HMEM As Long

Private Sub Command1_Click()
AccessMemory HMEM, WriteMem
AccessMemory HMEM, ReadMem
End Sub



Private Sub Form_Load()
HMEM = GlobalAlloc(&H40&, 12) 'Allocate 12 bytes from global heap
'Number of allocated bytes must be exactly EQU/LESS with length of MEMORY type
End Sub

Private Sub Form_Unload(Cancel As Integer)
GlobalFree HMEM
End Sub
