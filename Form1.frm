VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Drag/Drop Selected Text (from regular text box) Demo"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   4200
      TabIndex        =   2
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton MoveTextBox 
      Caption         =   "Move The Text Box"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   1800
      Width           =   5895
   End
   Begin VB.Label DragLbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DragLbl"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6120
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label CounterLbl 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Items In The List Box:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drag Selected Text To The List Box Below:"
      Height          =   195
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   3075
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Text Below... then drag to the list box above"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LocationNum, A$
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ Use this code as you see fit!                                                          ++
'++ I wrote this demo because I wanted to drag selected text from a regular VB text box to ++
'++ a list box.  Using a hidden label was the obvious cheap solution... but placing it was ++
'++ my biggest problem... placing it outside the text box was counter-intuitive for the    ++
'++ drag part of the operation and was the result of my using the wrong event...           ++
'++ the double_click event... the answer was ultimately under my nose... hope this helps!  ++
'++ This is easy enough with the RichTextBox... but a bit tougher (I thought) with the     ++
'++ regular VB text box.                                                                   ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub MoveTextBox_Click()
    'this sub just moves the text box and label to demonstrate that the location of the text box on the form doesn't matter.
    LocationNum = LocationNum + 1
    If LocationNum > 3 Then LocationNum = 1
    Select Case LocationNum
        Case 1
            Text1.Left = 240
            Text1.Top = 1800
        Case 2
            Text1.Left = 2160
            Text1.Top = 1800
        Case 3
            Text1.Left = 1320
            Text1.Top = 3240
    End Select
    'reposition the label
        Label1.Left = Text1.Left
        Label1.Top = Text1.Top - 240
End Sub

Private Sub Form_Load()
    LocationNum = 1 'initialize the location of the text box
    List1.Clear 'clear the list box
    CounterLbl.Caption = List1.ListCount ' show # of items in list
    'create some text to drag
    Msg = "There is some text in this text box for you to select and drag to the list box above." & vbCrLf & "Here is some more text and if that is not enough" & vbCrLf & "You can add your own text!"
    Text1.Text = Msg 'set the text to drag
    'the following attributes can be set at design time instead of wasting code... as I've done here
        DragLbl.Visible = False
        DragLbl.AutoSize = True
End Sub
Private Sub List1_DragDrop(Source As Control, X As Single, Y As Single)
'receive the dropped text
    List1.AddItem (DragLbl.Caption) ' add item to the listbox
    DragLbl.Caption = "" 'after the drop... clear the invisible label
    CounterLbl.Caption = List1.ListCount ' update item count of the listbox control
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This is the event I "should have used!"... didn't need the double click at all... since the mouse_up happens with the double click as well.
    'drag selected text
            If Text1.SelLength > 0 Then
                'remove leading and trailing space from selection... you don't have to do this
                'but leading/trailing spaces in selections can be problemsome.
                A$ = Trim(Text1.SelText)
                    'remove trailing comma if it exists... again, you may not need to do this
                    L = Len(A$)
                    If Right(A$, 1) = "," Then
                        A$ = Mid(A$, 1, L - 1)
                    End If
            Else
                Exit Sub
            End If
        DragLbl.Caption = A$ ' set the dragged text value
        DragLbl.Left = X + Text1.Left ' locate the "hidden label"
        DragLbl.Top = Y + Text1.Top ' locate the "hidden label"
        DragLbl.Drag 'start the drag operation
End Sub
