VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Sorting Examples"
   ClientHeight    =   6870
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   34
      Top             =   6720
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   8
      Left            =   2160
      TabIndex        =   31
      Top             =   6360
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   29
      Top             =   6360
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   26
      Top             =   5640
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   24
      Top             =   5280
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   22
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton cmdClrSort 
      Caption         =   "Clear Sorted"
      Height          =   375
      Left            =   2880
      TabIndex        =   19
      Top             =   3120
      Width           =   1095
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   16
      Top             =   4560
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   15
      Top             =   4200
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   255
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox lstSort 
      Height          =   1815
      Left            =   2760
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ListBox lstList 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Odd-Even Transposition Sort"
      Height          =   255
      Left            =   600
      TabIndex        =   33
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label15 
      Caption         =   "Shaker Sort"
      Height          =   255
      Left            =   2520
      TabIndex        =   32
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Radix Sort"
      Height          =   255
      Left            =   600
      TabIndex        =   30
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Other Sorts"
      Height          =   255
      Left            =   1080
      TabIndex        =   28
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Quick Sort"
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Merge Sort"
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Heap Sort"
      Height          =   255
      Left            =   600
      TabIndex        =   23
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "O(n log n) Sorts"
      Height          =   255
      Left            =   1200
      TabIndex        =   21
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "O(n^2) Sorts"
      Height          =   255
      Left            =   1440
      TabIndex        =   20
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Shell Sort"
      Height          =   255
      Left            =   2520
      TabIndex        =   18
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Bubble Sort"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4080
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label5 
      Caption         =   "Selection Sort"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Insertion Sort"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label lblItems 
      Caption         =   "0 items in list"
      Height          =   615
      Left            =   1560
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Sorted List"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Entry"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "List"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Menu mnuPop 
      Caption         =   "File"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Const Max_Items = 50
Dim list(Max_Items) As Integer
Dim temp(Max_Items) As Integer
Dim items As Integer

Private Sub cmdAdd_Click()
    Dim entry As Integer
    
    'get entry
    entry = Val(txtEntry.Text)
    
    'check for overflow
    If items = 50 Then
        MsgBox "Maximum number of items reached!", vbInformation, "Entry Error!"
        Exit Sub
    End If
    
    'everything is good
    items = items + 1
    lstList.AddItem (entry)
    lblItems = items & " items in list"
    
    'return to entry
    txtEntry.SetFocus
End Sub

Private Sub cmdClear_Click()
    'clear everything out
    items = 0
    InitList
    lstList.Clear
    lstSort.Clear
    txtEntry.Text = ""
    lblItems.Caption = items & " items in list"
End Sub

Private Sub cmdClrSort_Click()
    'clear out the sorted list and the sorted listbox
    InitList
    lstSort.Clear
End Sub

Private Sub cmdRemove_Click()
    'check for valid selection
    If lstList.Text = "" Then
        MsgBox "You must select an item to remove!", vbInformation, "Selection Error!"
        Exit Sub
    End If
    
    'remove the entry
    items = items - 1
    lstList.RemoveItem (lstList.ListIndex)
    lblItems = items & " items in list"
End Sub

Private Sub cmdSort_Click()
    'make sure there's enough items
    If items <= 1 Then
        MsgBox "Not enough items to sort!", vbInformation, "Array Formation Error!"
        Exit Sub
    End If
    
    'make the list
    FormList
    
    'sort the list
    If optSort(0).Value Then
        InsertionSort list(), items
    End If
    If optSort(1).Value Then
        SelectionSort list(), items
    End If
    If optSort(2).Value Then
        BubbleSort list(), items
    End If
    If optSort(3).Value Then
        ShellSort list(), items
    End If
    If optSort(4).Value Then
        HeapSort list(), items
    End If
    If optSort(5).Value Then
        MergeSort list(), temp(), items
    End If
    If optSort(6).Value Then
        QuickSort list(), items
    End If
    If optSort(7).Value Then
        RadixSort list(), temp(), items
    End If
    If optSort(8).Value Then
        ShakerSort list(), items
    End If
    If optSort(9).Value Then
        OETSort list(), items
    End If
    
    'display the list
    DisplayList
End Sub

Private Sub FormList()
    Dim i As Integer
    
    'get the value from the list and add it to the array
    For i = 0 To (items - 1)
        list(i) = Val(lstList.list(i))
    Next i
End Sub

Private Sub DisplayList()
    Dim i As Integer
    
    'clear the sorted list
    lstSort.Clear
    
    'add all of the values to the sorted list
    For i = 0 To (items - 1)
        lstSort.AddItem (list(i))
    Next i
End Sub

Private Sub Form_Load()
    'init form
    Me.Show
    mnuPop.Visible = False
    DoEvents
    
    'init variables
    items = 0
    InitList
    
    Debug.Print ""
    Debug.Print "Start of New Run"
End Sub

Private Sub InitList()
    Dim i As Integer
    
    'zero out the array
    For i = 0 To (Max_Items - 1)
        list(i) = 0
    Next i
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'popup the menu
    If Button = 2 Then
        frmMain.PopupMenu mnuPop
    End If
End Sub

Private Sub mnuAbout_Click()
    'show the about box
    MsgBox "This is the an exhibition of several sorting methods. " & vbCrLf & _
    "There are several more sorting methods that will be added eventually. " & vbCrLf & _
    "The Odd-Even Transposition Sort (OETS) although intended for parallel " & vbCrLf & _
    "processing, is only going to run in 1 thread on 1 processor, and has " & vbCrLf & _
    "been rewritten accordingly." & vbCrLf & vbCrLf & "Written by Guillaume Couture-Levesque" & _
    vbCrLf & "July 22nd 2003", vbInformation, "About This Program"
End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)
    'check for enter
    If KeyAscii = 13 Then
        cmdAdd_Click
    End If
End Sub
