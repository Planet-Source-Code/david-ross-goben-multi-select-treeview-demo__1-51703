VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestTv 
   Caption         =   "Test TreeView Modules -- David Goben. Feb 2004"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstFile 
      Height          =   255
      Left            =   6660
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   6720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox lstDir 
      Height          =   255
      Left            =   3300
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   6720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox chkCustomHotExp 
      Caption         =   "Allow Hot Expansions/Collapses"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      ToolTipText     =   "When checked, moving the mouse cursor over a custom folder will auto-expand/collapse it"
      Top             =   6240
      Value           =   1  'Checked
      Width           =   3075
   End
   Begin MSComctlLib.ImageList Images 
      Left            =   2280
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestTv.frx":0000
            Key             =   "drive"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestTv.frx":0452
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestTv.frx":05AC
            Key             =   "close"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestTv.frx":0706
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeViewCustom 
      Height          =   5295
      Left            =   180
      TabIndex        =   10
      Top             =   420
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   9340
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin VB.DriveListBox DriveDir 
      Height          =   315
      Left            =   3420
      TabIndex        =   9
      ToolTipText     =   "Select the drive whose contents you wish to display"
      Top             =   420
      Width           =   3135
   End
   Begin VB.CommandButton cmdFilesOnly 
      Caption         =   "Files Only"
      Height          =   315
      Left            =   8520
      TabIndex        =   8
      ToolTipText     =   "Allow only multiple files to be selected"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdAllowFolders 
      Caption         =   "Allow Folder Select"
      Height          =   315
      Left            =   6720
      TabIndex        =   7
      ToolTipText     =   "Allow multiple folder to be selected which are on the SAME folder level"
      Top             =   5880
      Width           =   1755
   End
   Begin VB.ListBox lstSelected 
      Height          =   5325
      Left            =   6660
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   420
      Width           =   3075
   End
   Begin VB.CommandButton cmdAddCustom 
      Caption         =   "Add to Tree List"
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      ToolTipText     =   "Add a custom Path. ie, Family\Goben\David"
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdClearCustom 
      Caption         =   "Clear"
      Height          =   315
      Left            =   180
      TabIndex        =   3
      ToolTipText     =   "Erase the contents of the custom Tree View"
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefreshDir 
      Caption         =   "Refresh Directory List"
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      ToolTipText     =   "Clear and refresh the directory listing with a fresh re-read of the selected drive"
      Top             =   5880
      Width           =   1935
   End
   Begin MSComctlLib.TreeView TreeViewDir 
      Height          =   4875
      Left            =   3420
      TabIndex        =   11
      Top             =   840
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   8599
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "Images"
      Appearance      =   1
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "Multi-Select using CTRL or SHIFT"
      ForeColor       =   &H80000015&
      Height          =   195
      Left            =   3780
      TabIndex        =   15
      Top             =   6240
      Width           =   2400
   End
   Begin VB.Label lblSel 
      AutoSize        =   -1  'True
      Caption         =   "Selected Items in Directory"
      Height          =   195
      Left            =   6660
      TabIndex        =   6
      Top             =   120
      Width           =   1890
   End
   Begin VB.Label lblCustom 
      AutoSize        =   -1  'True
      Caption         =   "Custom Tree View With Tracking Select"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2835
   End
   Begin VB.Label lblDir 
      AutoSize        =   -1  'True
      Caption         =   "Directory Trere View With Multi-Select"
      Height          =   195
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   2685
   End
End
Attribute VB_Name = "frmTestTv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------------------------------
' Copyright 2004 David Goben. All rights reserved. Feel free to paste these routine
' into your own programs. But don't do what some do and simply paste your name
' into the headings and tout them as your own personal inventions. I'm getting
' tired of helping some desperate soul out, and then find them taking my solution
' out onto the web, and boasting how THEY solved the problem...
'
'---------------------------------
' Notes regarding this sample code
'---------------------------------
' I tend to use Cbool() around integer numeric quantities, even though VB (and C)
' automatically converts these to Booleans during logical testing. I do not much
' care for default actions or forced casting, as it removes from the self-
' documentation from the code (don't get me started on those stupid default
' properties in controls...).
'
' On TreeView nodes with properties for Images and SelectedImage. I like to set
' BOTH to the open or closed state. Not doing so tends to be a nightmare, because
' depending on whether a node has been previously expanded or not, the visual
' effect can sometimes go haywire, an open folder may sometimes display a closed
' icon, and visa versa.
'
' Note that I precheck folder for content (other than "." and "..") when I am
' building a directory list. Rather than using a rather cool recursive subroutine
' to build the entire tree (and make dinner while it is trying to finish on my
' system), I pre-tag folders with a "+" if they have contents, and then, when the
' folder is actually expanded, check to see if I need to actually populate its
' content. The cool part of this is that folders that do not contain subfolders or
' files will not be pre-tagged with a "+", but it still runs like the wind. Reminder:
' a tree view is limited to a maximum of 32,000 entries. hence, on my system, it
' would lock up when this limit is reached, which only happens if I were to
' pre-populate every folder branch.
'
' I tend to prefer Long integers to regular Integers. Longs are processed much faster
' on a 32-bit system than the 16-bit Integers. Integer require internal conversion to
' a Long, anyway. So cut out the intermediate processing and shorten the CPU
' cycles-count.
'
' Multi-selects are only allows on items at the same folder level (same parent).
' When multi-selecting folders in a tree-view, only folders that are children of
' the SAME node can be multi-selected. Selecting additional folders or file
' on a different level will reset multi-select, and select only that last-selected
' item.
'-------------------------------------------------------------------------------
 
Private LastCustomNode As Node      'last custom node selected
Private cTV As cTvMultiSel          'set aside a variable of the multiselect class
Private FSO As FileSystemObject     'File System Object resource

Private Const MinWidth As Long = 10000
Private Const minHeight As Long = 5000

'*******************************************************************************
' Subroutine Name   : Form_Initialize
' Purpose           : Set up XP buttons if an XP system
'*******************************************************************************
Private Sub Form_Initialize()
  Call FormInitialize
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : App Entry. Set some button enablement
'*******************************************************************************
Private Sub Form_Load()
  Set FSO = New FileSystemObject    'create File System Object resource
  Set cTV = New cTvMultiSel         'instantiate the treeview multiselect class
  cTV.Init TreeViewDir              'assign the treeview associated with it to the class
  Me.cmdRefreshDir.Value = True     'refresh the directory listing
  Me.cmdClearCustom.Enabled = False 'disable the clear button for the custom list
  Me.cmdAllowFolders.Enabled = False
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Resizing form
'*******************************************************************************
Private Sub Form_Resize()
  Dim Wdth As Long, BtnTop As Long
  Static Resizing As Boolean
  
  If Me.WindowState = vbMinimized Then Exit Sub
  If Resizing Then Exit Sub
  Resizing = True
  If Me.Width < MinWidth Then Me.Width = MinWidth
  If Me.Height < minHeight Then Me.Height = minHeight
  
  With Me.chkCustomHotExp
    .Top = Me.ScaleHeight - .Height - 30
    Me.lblNote.Top = .Top
    BtnTop = .Top - Me.cmdClearCustom.Height - 60
  End With
  Me.cmdClearCustom.Top = BtnTop
  Me.cmdAddCustom.Top = BtnTop
  Me.cmdRefreshDir.Top = BtnTop
  Me.cmdAllowFolders.Top = BtnTop
  Me.cmdFilesOnly.Top = BtnTop
  
  With Me.TreeViewCustom
    Wdth = CLng(Me.ScaleWidth \ 3 - .Left * 1.3)
    .Width = Wdth
    Me.TreeViewDir.Width = Wdth
    Me.TreeViewDir.Left = .Left * 2 + Wdth
    Me.lstSelected.Width = Wdth
    Me.lstSelected.Left = .Left * 3 + Wdth * 2
    .Height = BtnTop - .Top - 60
    Me.lstSelected.Height = .Height
    Me.TreeViewDir.Height = BtnTop - Me.TreeViewDir.Top - 60
    Me.cmdAddCustom.Left = .Left + .Width - Me.cmdAddCustom.Width
  End With
  With Me.TreeViewDir
    Me.cmdRefreshDir.Left = .Left + .Width - Me.cmdRefreshDir.Width
    Me.DriveDir.Left = .Left
    Me.DriveDir.Width = .Width
    Me.lblDir.Left = .Left
    Me.lblNote.Left = .Left
  End With
  With Me.lstSelected
      Me.cmdAllowFolders.Left = .Left
      Me.cmdFilesOnly.Left = .Left + .Width - Me.cmdFilesOnly.Width
      Me.lblSel.Left = .Left
  End With
  Resizing = False
  DoEvents
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Release used resources
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Set cTV = Nothing                 'release the resources
  Set FSO = Nothing
End Sub

'*******************************************************************************
' Subroutine Name   : cmdAddCustom_Click
' Purpose           : Add a custom path to the custome folder
'*******************************************************************************
Private Sub cmdAddCustom_Click()
  Dim str As String
  Dim i As Long
'
' get a new path
'
  str = Trim$(InputBox("Enter a path to add to the custom list. Example:" & vbCrLf & _
        "MyRoot\MyLevel1\MyLevel2\MyLevel3", "Enter Custom Path", vbNullString))
'
' if something entered...
'
  If CBool(Len(str)) Then
'
' first force any accidental forward-slashes to backslashes
'
    i = InStr(1, str, "/")
    Do While i
      Mid$(str, i, 1) = "\"
      i = InStr(i + 1, str, "/")
    Loop
'
' add the new line to the folder
'
    AddPathToTreeView Me.TreeViewCustom, str
'
' Since the list has changed, allow clearing it
'
    Me.cmdClearCustom.Enabled = True
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cmdClearCustom_Click
' Purpose           : Clear the custom tree view
'*******************************************************************************
Private Sub cmdClearCustom_Click()
  ClearTreeView Me.TreeViewCustom     'clear the treeview very fast
  Set LastCustomNode = Nothing        'disable any previous node item
  Me.cmdClearCustom.Enabled = False   'disable clear button (already clear)
End Sub

'*******************************************************************************
' Subroutine Name   : TreeViewCustom_Expand
' Purpose           : Expanding the Custom list. Save the selected node
'*******************************************************************************
Private Sub TreeViewCustom_Expand(ByVal Node As MSComctlLib.Node)
  Set LastCustomNode = Node
End Sub

'*******************************************************************************
' Subroutine Name   : TreeViewCustom_MouseMove
' Purpose           : Moving the mouse cursor over the custom tree view
'*******************************************************************************
Private Sub TreeViewCustom_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim Nd As Node
  
  If Button = 0 And Shift = 0 Then                  'if no button or control key
    TreeViewTrackingSelect Me.TreeViewCustom, x, y  'select as found
    Set Nd = Me.TreeViewCustom.HitTest(x, y)        'get node there
    If Nd Is Nothing Then                           'if there IS NOT one there...
      Set LastCustomNode = Nothing                  'remove holder node
    Else
      If Not LastCustomNode Is Nd Then              'same as last?
        Set LastCustomNode = Nd                     'no, so update
        If Me.chkCustomHotExp.Value Then            'hot toggle?
          Nd.Expanded = Not Nd.Expanded             'yes, flip it
        End If
      End If
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : DriveDir_Change
' Purpose           : Drive selection changed
'*******************************************************************************
Private Sub DriveDir_Change()
  Dim Drv As Drive
  Dim Item As String
  Dim i As Long
  
  ClearTreeView Me.TreeViewDir                'clean its pipes
  Item = UCase$(Me.DriveDir.Drive)            'get the drivespec
  i = InStr(1, Item, "[")                     'strip human readability stuff
  If i Then Item = Trim$(Left$(Item, i - 1))
  Set Drv = FSO.GetDrive(Item)                'get selected drive object
'
' check drive readiness
'
  If Not Drv.IsReady Then                     'drive not ready?
    Me.cmdRefreshDir.Enabled = False          'no, so disable refresh
    MsgBox "Drive: " & Drv.Path & " is not ready", vbOKOnly Or vbExclamation, "Drive Not Ready"
  Else
    Me.cmdRefreshDir.Enabled = True           'is ok, so enable button
    Me.cmdRefreshDir.Value = True             'and refresh
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cmdRefreshDir_Click
' Purpose           : Refresh the directory treeview
'*******************************************************************************
Private Sub cmdRefreshDir_Click()
  Dim Item As String
  Dim i As Long
  Dim Nd As Node
  
  ClearTreeView Me.TreeViewDir                'clean its pipes
  cTV.Clear                                   'remove multiselect list
  Item = UCase$(Me.DriveDir.Drive)            'get the drivespec
  i = InStr(1, Item, "[")                     'strip human readability stuff
  If i Then Item = Trim$(Left$(Item, i - 1))
'
' create root node for the drive
'
  Item = Item & "\"
  Set Nd = Me.TreeViewDir.Nodes.Add(, tvwFirst, Item, Item, 1, 1)
  Nd.Expanded = True                          'tag expanded
  Call NodeExpand(Nd)                         'now open it up
  Me.cmdRefreshDir.Enabled = False            'disable refresh until something changes
End Sub

'*******************************************************************************
' Subroutine Name   : cmdAllowFolders_Click
' Purpose           : Toggle file/folders options
'*******************************************************************************
Private Sub cmdAllowFolders_Click()
  Me.cmdAllowFolders.Enabled = False
  Me.cmdFilesOnly.Enabled = True
  Me.lstSelected.Clear
  cTV.Clear
End Sub

'*******************************************************************************
' Subroutine Name   : cmdFilesOnly_Click
' Purpose           : Toggle file/folders options
'*******************************************************************************
Private Sub cmdFilesOnly_Click()
  Me.cmdAllowFolders.Enabled = True
  Me.cmdFilesOnly.Enabled = False
  Me.lstSelected.Clear
  cTV.Clear
End Sub

'*******************************************************************************
' Subroutine Name   : TreeViewDir_Collapse
' Purpose           : Collapse a node in the directory treeview
'*******************************************************************************
Private Sub TreeViewDir_Collapse(ByVal Node As MSComctlLib.Node)
  Select Case Node.Image
    Case 2
      Node.Image = 3                'set to closed folder if open folder
      Node.SelectedImage = 3
  End Select
  Me.cmdRefreshDir.Enabled = True             'enable refresh button
End Sub

'*******************************************************************************
' Subroutine Name   : TreeViewDir_Expand
' Purpose           : Expanding a treeview node for the directory
'*******************************************************************************
Private Sub TreeViewDir_Expand(ByVal Node As MSComctlLib.Node)
  Select Case Node.Image
    Case 4
      Exit Sub                        'ignore non-folders
    Case 3
      Node.Image = 2                  'is a folder (but not root), so show open
      Node.SelectedImage = 2
  End Select
  NodeExpand Node                       'expand node if not already populated
'
' little trick to show as much of the expanded node contents as possible, but
' ensuring that the opened node is still visible
'
  Node.Child.LastSibling.EnsureVisible  'this 2-line trick is pretty cool for
  Node.EnsureVisible                    '  displaying as much as possible of new stuff
'
' View has changed, so allow refresh (re-reading)
'
  Me.cmdRefreshDir.Enabled = True
End Sub

'*******************************************************************************
' Subroutine Name   : TreeViewDir_NodeClick
' Purpose           : clear multiselects if the we want to select files only and
'                   : the user clicked on a folder (NOTE: we tagged the keys for
'                   : folders with a trailing "\" to make for E-Z identification
'
' NOTE: OF COURSE, files AND folders are not (and should not) be allowed.
'*******************************************************************************
Private Sub TreeViewDir_NodeClick(ByVal Node As MSComctlLib.Node)
  Dim S As String                       'temp string
  Dim Idx As Integer
  
  Me.lstSelected.Clear
  If Not Me.cmdFilesOnly.Enabled Then   'files only?
    S = Node.Key                        'get path to node if so
    If Right$(S, 1) = "\" Then          'folder?
      cTV.Clear                         'yes, so simply clear the multiselect options
      If Node.Selected Then Me.lstSelected.AddItem Node.Text
      Exit Sub
    End If
  End If
  cTV.NodeClick Node                    'else process node clicks normally
  If cTV.SelCount Then                             'if there is a multiselection
    With cTV
      For Idx = 1 To .SelCount                    'loop through list
        Me.lstSelected.AddItem Me.TreeViewDir.Nodes(.Item(Idx))
      Next Idx                                    'do next
    End With
  Else
    If Node.Selected Then Me.lstSelected.AddItem Node.Text
  End If
  
End Sub

'*******************************************************************************
' Subroutine Name   : NodeExpand
' Purpose           : Expand a selected node
'*******************************************************************************
Private Sub NodeExpand(RootNode As Node)
  Dim Nd As Node
  Dim Item As String, RootPath As String, Tmp As String
  Dim Idx As Long
  Dim ReplaceFirst As Boolean
'
' first check to see if the node has been prepopulated...
'
  With RootNode
    If .Children = 1 Then                         'one child?
      ReplaceFirst = Right$(.Child.Key, 1) = "*"  'yes, dummy tag?
    End If
'
' if more than 1 child and it is not a dummy, then no need to populate it
'
    If .Children > 0 And ReplaceFirst = False Then Exit Sub
    RootPath = .Key & "\"                         'init root path
  End With
'
' read the contents of the current folder
'
  Item = Dir$(RootPath & "*.*", vbDirectory)      'get all files/dirs
  Do While CBool(Len(Item))                       'while something to do
    If Left$(Item, 1) <> "." Then                 'ignore "." and ".."
      Item = RootPath & Item                      'full path
      If GetAttr(Item) And vbDirectory Then       'folder?
        Me.lstDir.AddItem Item                    'yes, add to sorted folder list
      Else
        Me.lstFile.AddItem Item                   'else add to sorted file list
      End If
    End If
    Item = Dir$()                                 'read next directory entry
  Loop
'
' now populate with the sorted subfolder list
'
  With Me.lstDir
    For Idx = 0 To .ListCount - 1                 'check all subfolders
      Item = .List(Idx)                           'get a folder path
      Tmp = Mid$(Item, InStrRev(Item, "\") + 1)   'grab just folder name for display
      If ReplaceFirst Then                        'if we have a dummy node...
        ReplaceFirst = False                      'turn off the flag
        Set Nd = RootNode.Child                   'set local node to the dummy
        With Nd
          .Key = Item & "\"                       'populate it with the current data
          .Text = Tmp                             'folder name for display
          .Image = 3                              'closed folder
          .SelectedImage = 3                      'keep visuals saner by doing manually
        End With
      Else                    'not a dummy, so create and set a fresh node
        Set Nd = Me.TreeViewDir.Nodes.Add(RootNode.Index, tvwChild, _
                 Item & "\", Tmp, 3, 3)
      End If
'
' now see if the sub-folder has content
'
      Tmp = Dir$(Item & "\*.*", vbDirectory)
      Do While CBool(Len(Tmp))
        If Left$(Tmp, 1) <> "." Then              'ignore "." and ".."
          Item = Item & "\*"                      'it does, so create dummy node
                                                  'this will force a "+" on the parent
          Call Me.TreeViewDir.Nodes.Add(Nd.Index, tvwChild, Item, Item)
          Exit Do                                 'no need to check further
        End If
        Tmp = Dir$()
      Loop
    Next Idx
  End With
'
' now populate with the sorted file list
'
  With Me.lstFile
    For Idx = 0 To .ListCount - 1                 'process all files
      Item = .List(Idx)                           'get a file path
      Tmp = Mid$(Item, InStrRev(Item, "\") + 1)   'grab just filename for display
      If ReplaceFirst Then                        'if we have a dummy node
        ReplaceFirst = False                      'turn off the flag
        Set Nd = RootNode.Child                   'set local node to the dummy
        With Nd
          .Key = Item                             'populate it with the current data
          .Text = Tmp                             'stuff filename for display
          .Image = 4                              'file icon
          .SelectedImage = 4
        End With
      Else                    'not a dummy, so create and set a fresh node
        Call Me.TreeViewDir.Nodes.Add(RootNode.Index, tvwChild, Item, Tmp, 4, 4)
      End If
    Next Idx
  End With
'
' clear contents of the temporary sorted folder and file lists
'
  Me.lstDir.Clear
  Me.lstFile.Clear
End Sub

'*******************************************************************************
' Subroutine Name   : lstSelected_Click
' Purpose           : Ignore selections on any clicks in the list of selected items
'*******************************************************************************
Private Sub lstSelected_Click()
  Me.lstSelected.ListIndex = -1
End Sub
