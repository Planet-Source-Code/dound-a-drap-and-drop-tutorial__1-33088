VERSION 5.00
Begin VB.Form frmDragDropTutorial 
   BackColor       =   &H8000000A&
   Caption         =   "Drag/Drop Tutorial"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDragDropTutorial.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDrag 
      Caption         =   ":- ) Drag Me Around :-)"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDragDropTutorial.frx":08CA
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label lblBlue 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblWhite 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1380
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblRed 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDragDropTutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tutorial Created By Dound (dound@kbs.recongamer.com)
Dim StartX, StartY 'Variables to hold where on the control you clicked when you started to drag it

Private Sub cmdDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'You could use picture boxes instead of command buttons too (and that is probably
    'what must coders need drag/drop for)
    
    StartX = X 'Store where on the button your mouse was when you started to drag it.
    StartY = Y
    cmdDrag.Drag vbBeginDrag 'Starts drag
End Sub

Private Sub cmdDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdDrag.Drag vbEndDrag 'Ends Drag
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move X - StartX, Y - StartY 'Move the control to where the user dropped it
End Sub

Private Sub lblBlue_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Source.Move lblBlue.Left, lblBlue.Top 'Position the control over this label
    
    cmdDrag.Drag vbEndDrag 'Tells the control it is no longer being dragged
End Sub

Private Sub lblInstructions_DragDrop(Source As Control, X As Single, Y As Single)
    'Move the control to where the user dropped it
    Source.Move X - StartX + lblInstructions.Left, Y - StartY + lblInstructions.Top
End Sub

Private Sub lblRed_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Source.Move lblRed.Left, lblRed.Top 'Position the control over this label
    
    cmdDrag.Drag vbEndDrag 'Tells the control it is no longer being dragged
End Sub

Private Sub lblWhite_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Source.Move lblWhite.Left, lblWhite.Top 'Position the control over this label
    
    cmdDrag.Drag vbEndDrag 'Tells the control it is no longer being dragged
End Sub
