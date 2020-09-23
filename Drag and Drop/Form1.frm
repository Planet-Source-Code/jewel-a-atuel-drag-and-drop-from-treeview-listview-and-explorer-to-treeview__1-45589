VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Launch Explorer"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Selected Items Listview"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView1 
      CausesValidation=   0   'False
      Height          =   6375
      Left            =   3720
      TabIndex        =   1
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   11245
      View            =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      Icons           =   "imgListTvw"
      SmallIcons      =   "imgListTvw"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgListTvw 
      Left            =   3840
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "FolderClosed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0554
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0666
            Key             =   "Root"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvProject 
      Height          =   6375
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   11245
      _Version        =   393217
      Style           =   7
      ImageList       =   "imgListTvw"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blnDragging As Boolean
Private tvwDragging As Boolean
Private expDragging As Boolean
Private CurrentNode As Node
Private Hnum As Integer
Dim o_Fso As Scripting.FileSystemObject
Private lstDragging As Boolean


Private Sub Command1_Click()
On Error Resume Next
    Dim X
    X = Shell("Explorer.exe", 1)
End Sub

Private Sub Command2_Click()
    Dim i As Integer
    Dim l As ListItems
    Dim intcount As Integer

    intcount = ListView1.ListItems.Count
    i = 1
    While i <= intcount
        
        If ListView1.ListItems(i).Selected = True Then
             ListView1.ListItems.Remove (i)
             intcount = intcount - 1
        Else
            i = i + 1
        
        End If
    Wend
    
    
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
With tvProject.Nodes
        .Add , , "Root", "Root Item", "Root"
        '// add some child folders
        .Add "Root", tvwChild, "ChildFolder1", "Child Folder 1", "FolderClosed"
        .Add "Root", tvwChild, "ChildFolder2", "Child Folder 2", "FolderClosed"
        .Add "Root", tvwChild, "ChildFolder3", "Child Folder 3", "FolderClosed"
        '// add some children to the folders
        .Add "ChildFolder1", tvwChild, "Child1OfFolder1", "Child 1 Of Folder 1", "Item"
        .Add "ChildFolder1", tvwChild, "Child2OfFolder1", "Child 2 Of Folder 1", "Item"
        .Add "ChildFolder2", tvwChild, "Child1OfFolder2", "Child 1 Of Folder 2", "Item"
    End With
    Hnum = 1
    tvwDragging = False

Set CurrentNode = tvProject.Nodes(1)
End Sub



Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    'ListView1.ListItems(tvProject.SelectedItem.Key).Text = NewString
    tvProject.Nodes(ListView1.SelectedItem.Key).Text = NewString
End Sub

Private Sub tvProject_LostFocus()
    tvwDragging = False
End Sub

Private Sub tvProject_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    Dim nodNode As Node
    '// get the node we are over
    Set nodNode = tvProject.HitTest(X, Y)
    If nodNode Is Nothing Then Exit Sub '// no node
    '// ensure node is actually selected, just incase we start dragging.
    nodNode.Selected = True
    tvwDragging = True
    
    
End Sub

Private Sub tvProject_NodeClick(ByVal Node As MSComctlLib.Node)
Dim sNode As Node

    Set CurrentNode = Node
    
    ListView1.ListItems.Clear
    Set sNode = Node.Child
    
    For i = 1 To Node.Children
        ListView1.ListItems.Add i, sNode.Key, sNode.Text, sNode.Image, sNode.Image
        Set sNode = sNode.Next
        
        
    Next
    
    Set sNode = Nothing
    
End Sub

Private Sub tvProject_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim strSourceKey As String
Dim nodTarget    As Node

    Set o_Fso = New Scripting.FileSystemObject
    
    Set nodTarget = tvProject.HitTest(X, Y)
    
    If expDragging = True Then
        Dim numFiles As Integer
        numFiles = Data.Files.Count
        'Count number of files
        'Add all dropped files into the list
        
        For i = 1 To numFiles
            'File or directory?
            If (GetAttr(Data.Files(i)) And vbDirectory) = vbDirectory Then
                'lstFiles.AddItem "Directory: " & Data.Files(i)
            Else
                Dim s
                s = Data.Files(i)
                s = o_Fso.GetFileName(s)
                tvProject.Nodes.Add nodTarget, tvwChild, "k" & Hnum, o_Fso.GetFileName(Data.Files(i)), "Item"
                Hnum = Hnum + 1
            End If
        Next i
    End If

    If lstDragging = True Then
'    If ListView1.SelectedItem Is Nothing Then Exit Sub
        For i = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(i).Selected = True Then
                tvProject.Nodes.Remove ListView1.ListItems(i).Key
                tvProject.Nodes.Add nodTarget, tvwChild, ListView1.ListItems(i).Key, ListView1.ListItems(i).Text, "Item"
            End If
        Next

    End If



If tvwDragging = True Then
    '// get the carried data
    strSourceKey = Data.GetData(vbCFText)
    '// get the target node
    
    '// if the target node is not a folder or the root item
    '// then get it's parent (that is a folder or the root item)
    If nodTarget.Image <> "FolderClosed" And nodTarget.Key <> "Root" Then
        Set nodTarget = nodTarget.Parent
    End If
    '// move the source node to the target node
    Set tvProject.Nodes(strSourceKey).Parent = nodTarget
    '// NOTE: You will also need to update the key to reflect the changes
    '// if you are using it
    '// we are not dragging from this control any more
    blnDragging = False
    '// cancel effect so that VB doesn't muck up your transfer
    Effect = 0
End If

    nodTarget.Expanded = True
    tvProject_NodeClick nodTarget
    tvProject.SetFocus
    tvProject.SelectedItem = nodTarget
    Set o_Fso = Nothing
End Sub



Private Sub tvProject_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
          
    'dragging from explorer
'    If expDragging = True Then
        If Data.GetFormat(vbCFFiles) Then
            Effect = vbDropEffectCopy
            tvwDragging = False
            expDragging = True
        End If
'    End If
    
    'dragging from treeview
    If tvwDragging = True Then
        Dim nodNode As Node
'       // set the effect
        Effect = vbDropEffectMove
'        // get the node that the object is being dragged over
        Set nodNode = tvProject.HitTest(X, Y)
        If nodNode Is Nothing Or blnDragging = False Then
'            // the dragged object is not over a node, invalid drop target
'            // or the object is not from this control.
        Effect = vbDropEffectNone
        End If
        tvwDragging = True
    End If

    'dragging from listview
    If lstDragging = True Then
        Effect = vbDropEffectMove
    End If
 
        
End Sub

Private Sub tvProject_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    
    'drag from explorer
    If Data.GetFormat(vbCFFiles) = True Then
        tvProject.OLEDragMode = ccOLEDragManual
        tvProject.OLEDropMode = ccOLEDropManual
        '// Set the effect to move
        AllowedEffects = vbDropEffectMove
        expDragging = True
        tvwDragging = False
        lstDragging = False
    End If
    
    'drag from treeview
    If tvwDragging = True Then
        tvProject.OLEDragMode = ccOLEDragAutomatic
        tvProject.OLEDropMode = ccOLEDropManual
        '// Set the effect to move
        AllowedEffects = vbDropEffectMove
        '// Assign the selected item's key to the DataObject
        Data.SetData tvProject.SelectedItem.Key
        blnDragging = True
        lstDragging = False
        expDragging = False
    End If
    
    If lstDragging = True Then
        tvProject.OLEDragMode = ccOLEDragAutomatic
        tvProject.OLEDropMode = ccOLEDropManual
        AllowedEffects = vbDropEffectMove
        tvwDragging = False
        expDragging = False
        
    End If
    
    
End Sub
Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set ListView1.SelectedItem = ListView1.HitTest(X, Y)
    lstDragging = True
   
End Sub
Private Sub ListView1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
End Sub

