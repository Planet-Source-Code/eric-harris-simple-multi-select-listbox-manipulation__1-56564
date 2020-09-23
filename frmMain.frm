VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multi-Select Listbox Manipulation"
   ClientHeight    =   6555
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "<< R&emove All"
      Height          =   375
      Left            =   2430
      TabIndex        =   4
      Top             =   2955
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddAll 
      Caption         =   "A&dd All >>"
      Height          =   375
      Left            =   2430
      TabIndex        =   2
      Top             =   1755
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add >"
      Height          =   375
      Left            =   2430
      TabIndex        =   1
      Top             =   1275
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "< &Remove"
      Height          =   375
      Left            =   2430
      TabIndex        =   3
      Top             =   2475
      Width           =   1215
   End
   Begin VB.ListBox lstList2 
      Height          =   4155
      Left            =   3990
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   225
      Width           =   1935
   End
   Begin VB.ListBox lstList1 
      Height          =   4155
      Left            =   150
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   225
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Multi-select listbox usage"
      Height          =   1815
      Left            =   150
      TabIndex        =   6
      Top             =   4560
      Width           =   5775
      Begin VB.Label Label6 
         Caption         =   "3) Use both the Ctrl and Shift keys together to select multiple ranges."
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   5535
      End
      Begin VB.Label Label5 
         Caption         =   $"frmMain.frx":0000
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "1) Hold down the Ctrl key to select multiple items."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   5535
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "&List"
      Begin VB.Menu mnuListClear 
         Caption         =   "&Clear"
         Begin VB.Menu mnuListClearList1 
            Caption         =   "List &1"
         End
         Begin VB.Menu mnuListClearList2 
            Caption         =   "List &2"
         End
         Begin VB.Menu mnuListClearBoth 
            Caption         =   "&Both"
         End
      End
      Begin VB.Menu mnuListReset 
         Caption         =   "&Reset"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

Dim i As Integer ' Just a counter

' If list1 is empty, exit routine.
If lstList1.ListIndex = -1 Then Exit Sub

' Loop through all items in list1
For i = lstList1.ListCount - 1 To 0 Step -1
    ' If a selected item is found add it to list2
    ' then remove it from list1
    If lstList1.Selected(i) = True Then
        lstList2.AddItem lstList1.List(i)
        lstList1.RemoveItem i
        cmdRemove.Enabled = True
        cmdRemoveAll.Enabled = True
    End If
Next i

' If the last item was added, then disable the add button.
If lstList1.ListCount = 0 Then
    cmdAdd.Enabled = False
    cmdAddAll.Enabled = False
End If

End Sub

Private Sub cmdAddAll_Click()

Dim i As Integer ' Just a counter

' If list1 is empty, exit routine.
If lstList1.ListIndex = -1 Then Exit Sub

' Loop through all items in list1 adding each to list2
For i = 0 To lstList1.ListCount - 1
    lstList2.AddItem lstList1.List(i)
Next i

' Clear contents of list1
lstList1.Clear

' There are no more items to add, so disable the add
' buttons and enable the remove buttons.
cmdAdd.Enabled = False
cmdAddAll.Enabled = False
cmdRemove.Enabled = True
cmdRemoveAll.Enabled = True

End Sub

Private Sub cmdRemove_Click()

Dim i As Integer ' Just a counter

' If list2 is empty, exit routine.
If lstList2.ListIndex = -1 Then Exit Sub

' Loop through all items in list2
For i = lstList2.ListCount - 1 To 0 Step -1
    ' If a selected item is found add it to list1
    ' then remove it from list2
    If lstList2.Selected(i) = True Then
        lstList1.AddItem lstList2.List(i)
        lstList2.RemoveItem i
        cmdAdd.Enabled = True
        cmdAddAll.Enabled = True
    End If
Next i

' If the last item was removed, then disable the remove button.
If lstList2.ListCount = 0 Then
    cmdRemove.Enabled = False
    cmdRemoveAll.Enabled = False
End If

End Sub

Private Sub cmdRemoveAll_Click()

Dim i As Integer ' Just a counter

' If list2 is empty, exit routine.
If lstList2.ListIndex = -1 Then Exit Sub

' Loop through all items in list2 adding each to list1
For i = 0 To lstList2.ListCount - 1
    lstList1.AddItem lstList2.List(i)
Next i

' Clear contents of list1
lstList2.Clear

' There are no more items to remove, so disable the remove
' buttons and enable the add buttons.
cmdRemove.Enabled = False
cmdRemoveAll.Enabled = False
cmdAdd.Enabled = True
cmdAddAll.Enabled = True


End Sub

Private Sub Form_Load()

' Call the subroutine to load the default list items
loadDefaultListItems

' These 'If' statements are here in case either list is initially unpopulated.
' If either list is empty, then disable the corresponding buttons.
If lstList1.ListCount = 0 Then
    cmdAdd.Enabled = False
    cmdAddAll.Enabled = False
' This highlights the first item in list 1, if you want to.
Else: lstList1.Selected(0) = True
End If
If lstList2.ListCount = 0 Then
    cmdRemove.Enabled = False
    cmdRemoveAll.Enabled = False
' This highlights the first item in list 2, if you want to.
Else: lstList2.Selected(0) = True
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim Msg As Integer 'This just stores the msgBox answer.

' Asks the user a Yes/No question, when they exit the program.
Msg = MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Close listbox project")

' If they clicked No, then set Cancel to True, so that the program
' will remain open.
If Msg = vbNo Then
Cancel = True
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
' Show the About form
frmAbout.Show
End Sub


Private Sub mnuListClearBoth_Click()
' Clear both lists
lstList1.Clear
lstList2.Clear

' Disable the Add and Remove buttons since both lists are empty
cmdAdd.Enabled = False
cmdAddAll.Enabled = False
cmdRemove.Enabled = False
cmdRemoveAll.Enabled = False
End Sub

Private Sub mnuListClearList1_Click()
' Clear List1
lstList1.Clear

' Disable the Add buttons since List1 is empty
cmdAdd.Enabled = False
cmdAddAll.Enabled = False
End Sub

Private Sub mnuListClearList2_Click()
' Clear List2
lstList2.Clear

' Disable the Remove buttons since List2 is empty
cmdRemove.Enabled = False
cmdRemoveAll.Enabled = False
End Sub

Private Sub loadDefaultListItems()

' Populate list 1 with some items.
With lstList1
    .AddItem "List 1 - Item 1"
    .AddItem "List 1 - Item 2"
    .AddItem "List 1 - Item 3"
    .AddItem "List 1 - Item 4"
    .AddItem "List 1 - Item 5"
    .AddItem "List 1 - Item 6"
    .AddItem "List 1 - Item 7"
    .AddItem "List 1 - Item 8"
    .AddItem "List 1 - Item 9"
    .AddItem "List 1 - Item 10"
End With

' Populate list 2 with some items.
With lstList2
    .AddItem "List 2 - Item 1"
    .AddItem "List 2 - Item 2"
    .AddItem "List 2 - Item 3"
    .AddItem "List 2 - Item 4"
    .AddItem "List 2 - Item 5"
    .AddItem "List 2 - Item 6"
    .AddItem "List 2 - Item 7"
    .AddItem "List 2 - Item 8"
    .AddItem "List 2 - Item 9"
    .AddItem "List 2 - Item 10"
End With

End Sub

Private Sub mnuListReset_Click()

'Let's first clear both lists, otherwise it will keep appending the default items.
lstList1.Clear
lstList2.Clear

' Call the Form_Load to re-load the default list items.
Form_Load

End Sub
