VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTvMultiSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'~cTvMultiSel.cls;modSortCollection.bas;
'Allow multiple entries to be selected in a treeview
'*******************************************************************************
' cTvMultiSel - This class allows you to perform multi-selections in TreeView
'               control.
'
'               Clicking a treeview entry using the Cntrl key will highlight the
'               entry and any additional entries you click. The first selection does
'               not require that the Cntrl key held down, though it is allowed. By
'               default, selecting entries on different branch levels will clear
'               previous selections, as multi-select is allowed only on the same
'               level (nodes having the same parent). See note below to allow
'               cross-branch selections.
'
'               Using the SHIFT key allows you to choose a range, from the previous
'               entry selected (using Cntrl or not) to the current selection. Any
'               previous multi-selections will of course be cleared, as SHIFT works
'               with only the last selection and the current. Shift works on only the
'               same branch level (nodes with the same parent node).
'
' NOTE: Normally, multi-selects are not allowed in entries on different levels
'       (having different parent nodes). Setting the AllowDifferentFolders property
'       to True will allow selections on different levels. However, SHIFT selections
'       will still not allow this, and will work as if AllowDifferentFolders was not
'       set.
'
' Note: This class supports clicking a previously selected multi-select node and
'       thus unselecting it, as is common practive in Windows Explorer.
'
' OPERATIONAL NOTES:
' If you are going to clear a TreeView control of its contents, be sure to ALSO afterward
' invoke the classes' Clear subroutine with a FALSE parameter, or, by first clearing
' the TreeView and then invoking Clear, you need not bother with the parameter.
'
' Items are stored in the list in numerically ascending order.
'
' NOTE: This class used the module: modSortCollection.bas
'
' EXAMPLE ======================================================================
' You should declare a variable (or variables) of the class in the heading of your form
' containing the TreeView(s), and then in the Form_Load event, Set them, and assign
' to them their TreeView control. In the Form_Unload event, you should release their
' resources, thus:
'-------------------------------
''Dim tv As cTvMultiSel         'set aside a variable of the multiselect class
''
''Private Sub Form_Load()
''  Set tv = New cTvMultiSel    'instantiate the treeview multiselect class
''  tv.Init TreeView1           'assign the treeview associated with it to the class
''End Sub
''
''Private Sub Form_Unload()
''  Set tv = Nothing            'release the classes' resources
''End Sub
'-------------------------------
'
' Call the Clear function any time you wish to unselect any selected items in the
' treeview, thus
'-------------------------------
''  TreeView1.Clear             'or better, call ClearTreeView() in modClearTreeView.bas
''  tv.Clear                    'erase any possible muti-selections
'-------------------------------
'
' Call the NodeClick subroutine from within your TreeView's NodeClick event. It
' is OK to pre-parse the Node provided to the event in case you do not want certain
' nodes to be selected. For example, a common practice is to prevent directory nodes
' from being selected, and when they are, then other selections are cleared. One way
' to track this to to append a backslash "\" to the end of a branch nodes' Key
' property. Another way is to inspect the Image property for possible image indexes
' that denote branches. For example:
'
'-------------------------------
''Public Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
''  Dim S As String                       'temp string
''
''  S = Node.Key                          'get path to node
''  If Right$(S, 1) = "\" Then            'folder?
''    tv.Clear                            'yes, so simply clear the multiselect options
''  Else
''    tv.NodeClick Node                   'else process node clicks normally
''  End If
''End Sub
'-------------------------------
'
' To scan the list, do it in a manner simmilar to shown here:
'-------------------------------
''  Dim Idx As Integer
''
''  If tv.SelCount Then                             'if there is a multiselection
''    With tv
''      For Idx = 1 To .SelCount                    'loop through list
''        Debug.Print "Index = " & CStr(.Item(Idx)) 'grab each item
''        Debug.Print "Node = " & TreeView1.Nodes(.Item(Idx))
''      Next Idx                                    'do next
''    End With
''  End If
'-------------------------------
'
' You can manually add to the list using the boolena AddToList() function.
' It will return False if the Init subroutine has not yet been called, or if the
' provided Node index is less than 1 (0 or negative). It will also return False
' if the item is already in the selection list. If it returns True, then the
' Node index had been added, and the node associated with the index is highlighted.
' Be aware that this function follows the very same rules as applies to using the
' CTRL key and picking a node.
'-------------------------------
''  Dim Nd As Node
'''
''' get node info to Nd here...
'''
''  Debug.Print "add node to multi-select list: " & tv.AddToList(Nd.Index)
''  tv.AddToList 20   'add node index 20 to list. Ignore result. We simply
''                    'want to make sure that it in the list of items selected
'-------------------------------
'*******************************************************************************
' NOTE: This class expects the component "Microsoft Windows Common Controls" to
'       be selected in your project's Components list.
'*******************************************************************************

Option Explicit
'
' API goodies
'
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_CONTROL = &H11
Private Const VK_SHIFT = &H10
'
' our private booty
'
Private m_BaseNode As Node              'bottom node for drilling operations
Private m_coltvMultisel As Collection   'multi-select storage list
Private m_LasttvMSkey As Long           'last item selected
Private m_LasttvMSParent As Long        'parent whose children are being multi-selected
Private m_tv As TreeView                'local copy of treeview
Private m_AllowDiffFolders As Boolean   'allow selections in different folders

'*******************************************************************************
' Subroutine Name   : Class_Initialize
' Purpose           : Set aside the collection resource, and init a few variables
'*******************************************************************************
Private Sub Class_Initialize()
  Set m_coltvMultisel = New Collection  'instantiate our collection class
  m_LasttvMSkey = 0                     'init the last key selected to 0 (1 less minimum)
  m_LasttvMSParent = 0                  'indicate no parent node yet selected
  m_AllowDiffFolders = False            'start by not allowing multiple level selections
End Sub

'*******************************************************************************
' Subroutine Name   : Class_Terminate
' Purpose           : When the class is set to nothing, remove the collection
'                   : resource.
'*******************************************************************************
Private Sub Class_Terminate()
  Set m_tv = Nothing                    'release the tv pointer (not needed, but clean...)
  Set m_coltvMultisel = Nothing         'release the collection resource
End Sub

'*******************************************************************************
' Subroutine Name   : Init
' Purpose           : Assign the Treeview control associated with an instance of the
'                   : class. This is used internally.
'*******************************************************************************
Public Sub Init(tv As TreeView)
  Call Clear(False)                     'ensure the list is clean (can be used for
                                        'more than one treeview control, but it is of
                                        'course better to instantiate a separate class.
  Set m_tv = tv                         'assign our local reference pointer to the treeview
End Sub

'*******************************************************************************
' Return or assign a flag indicating if items from multiple levels (differen parents)
' can be selected. By default, only multiple items on the same level can be selected.
' Note that a SHIFT operations stillr requires that items be at the same level.
'*******************************************************************************
Public Property Get AllowDifferentFolders() As Boolean
  If m_tv Is Nothing Then Exit Property       'ignore if the Init has not been called
  AllowDifferentFolders = m_AllowDiffFolders  'return the property state
End Property

Public Property Let AllowDifferentFolders(Flag As Boolean)
  If m_tv Is Nothing Then Exit Property       'ignore if the Init has not been called
  m_AllowDiffFolders = Flag                   'assign the new property state
End Property

'*******************************************************************************
' Get Name          : SelCount
' Purpose           : Return the number of selections stored in the class.
' Outputs           : Return -1 if the Treeview has not be set to the Classes' Init
'                   : subroutine.
'                   :
'                   : Return the number of selections in the treeview
'*******************************************************************************
Public Property Get SelCount() As Integer
  If m_tv Is Nothing Then
    SelCount = -1                     'return impossible count if Init not called
  Else
    SelCount = m_coltvMultisel.Count  'else retun the number of items selected in the tv
  End If
End Property

'*******************************************************************************
' Get Name          : Item
' Purpose           : Return the Treeview Index stored in the class for the indexed
'                   : item. 0 is returned if the treeview has not been set to the
'                   : class via the Init sub, or if the Index is invalid
'*******************************************************************************
Public Property Get Item(Index As Integer) As Long
Attribute Item.VB_MemberFlags = "200"
  If m_tv Is Nothing Then Exit Property       'ignore if the Init has not been called
  If Index > m_coltvMultisel.Count Then Exit Property 'index too high
  If Index < 1 Then Exit Property                     'index too low
  'the above tests will of course invoke errors in the caller's routine, but
  'he-he-heee, they deserve it if they do not init things preoperty...
  Item = CLng(m_coltvMultisel(Index))                 'else pourridge is just right...
End Property

'*******************************************************************************
' Subroutine Name   : NodeClick
' Purpose           : This routine should be called by the NodeClick event for the
'                   : TreeView control. Preparse the Node as may be needed to
'                   : determine if it is a valid selection. If the node can be clicked,
'                   : then call this with the Node object provided to the event.
'*******************************************************************************
Public Sub NodeClick(ByVal Node As MSComctlLib.Node)
  Dim Idx As Long, Lst As Integer, I As Long
  Dim Nd As Node
  Dim Inc As Boolean
  
  If m_tv Is Nothing Then Exit Sub       'ignore if the Init has not been called
'
' set parent if parent not set
'
  If m_coltvMultisel.Count = 0 Then m_LasttvMSParent = Node.Parent.Index
'---
' check the key state. First check if the SHIFT key has been pressed
'---
  If GetKeyState(VK_SHIFT) < 0 Then       'shift key held down?
    Lst = Abs(m_LasttvMSkey)              'yes, get last key (don't care about sign)
    If Lst Then                           'if define, if not same as current parent, kill
      If m_tv.Nodes(Lst).Parent.Index <> Node.Parent.Index Then Lst = 0
    End If
'
' find direction that last is from the current
'
    If Lst Then                           'if last defined, we'll use it as a list term.
      Clear                               'remove any current defs
      Idx = Node.Index                    'get current node index
      Set Nd = Node.Parent.Child          'point to first entry
      Do While Nd.Next.Index              'scan while there are siblings
        If Nd.Index = Idx Then            'found Idx node first?
          Inc = True                      'forward offset from current select
          Exit Do
        ElseIf Nd.Index = Lst Then        'else backward if found Lst node first
          Inc = False
          Exit Do
        End If
        Set Nd = Nd.Next                  'point to next node
      Loop
      
      Set Nd = Node                       'starting node is current selection
      m_coltvMultisel.Add Nd.Index, "K" & CStr(Nd.Index)  'add first to list
      m_tv.Nodes(Idx).BackColor = vbHighlight             'mark it as selected
      Do While Nd.Index <> Lst                            'do all
        If Inc Then
          Set Nd = Nd.Next                                'go down the list
        Else
          Set Nd = Nd.Previous                            'go up the list
        End If
        m_coltvMultisel.Add Nd.Index, "K" & CStr(Nd.Index)  'add item to list
        m_tv.Nodes(Nd.Index).BackColor = vbHighlight        'mark it as selected
      Loop
      m_LasttvMSkey = Node.Index          'save last as current index
      m_LasttvMSParent = Node.Parent.Index 'save parent
    Else                                  'no previous key
      Clear                               'so clear list
      m_LasttvMSkey = Node.Index          'save current as last key
      m_LasttvMSParent = Node.Parent.Index  'save parent
      
      m_coltvMultisel.Add m_LasttvMSkey, "K" & CStr(m_LasttvMSkey)  'add to list
      m_tv.Nodes(m_LasttvMSkey).BackColor = vbHighlight             'mark it as selected
    End If
'---
' not SHIFT, so check for CNTRL...
'---
  ElseIf GetKeyState(VK_CONTROL) < 0 Then 'control key held down?
    If m_LasttvMSkey < 0 Then             'non-Cntrl index (set by a normal click)
      m_LasttvMSkey = -m_LasttvMSkey      'make positive
'
' if we allow multi-level selections, force parent to be always current
'
      If m_AllowDiffFolders Then m_LasttvMSParent = m_tv.Nodes(m_LasttvMSkey).Parent.Index
'
' if parent matches, then add the last key index to the list and mark the item in the tv
' Note: We add it here because a normal pick is not automatically added to our list.
'
      If m_tv.Nodes(m_LasttvMSkey).Parent.Index = m_LasttvMSParent Then
        m_coltvMultisel.Add m_LasttvMSkey, "K" & CStr(m_LasttvMSkey)
        m_tv.Nodes(m_LasttvMSkey).BackColor = vbHighlight
      End If
    End If
'
' pick up the index of the currently selected node (with the Cntrl key)
'
    Lst = Node.Index
'
' if we allow multi-level selections, force parent to be always current
'
    If m_AllowDiffFolders Then m_LasttvMSParent = m_tv.Nodes(Lst).Parent.Index
'
' if parents match, then we will be adding it to our selection list
'
    If m_tv.Nodes(Lst).Parent.Index <> m_LasttvMSParent Then
      Call Clear(True)                                          'clear list
      m_LasttvMSParent = m_tv.Nodes(Lst).Parent.Index 'set new parent
    End If
'
' set an error trap here in case the user is clicking a previously clicked selection,
' which would indicate that they wish to unselect it
'
    On Error Resume Next
    m_coltvMultisel.Add Lst, "K" & CStr(Lst)        'add new node to index
    If Err.Number Then                              'If error, we are unpicking...
      With m_coltvMultisel                          'we will need to locate the duplicate
        For Idx = 1 To .Count
          If .Item(Idx) = Lst Then                  'found it?
            With m_tv.Nodes(Lst)                    'yes
              .Selected = False                     'unselect it
              .BackColor = vbWindowBackground       'remove its selection highlight
            End With
            .Remove Idx                             'then remove it from the collection
            Exit For                                'done with search
          End If
        Next Idx
      End With
'
' if the node index was added to the collection without error
'
    Else
      Node.BackColor = vbHighlight                  'then simply highlight it
    End If
    On Error GoTo 0                                 'done with error trap
    m_LasttvMSkey = Lst                             'save current index as last key
'---
' not SHIFT, or CNTRL, so assume just a normal, Sunday afternoon click event
'---
  Else
    Clear                                           'clear the multi-select list
    m_LasttvMSkey = -Node.Index                     'tag last key as normal key (negative)
  End If
End Sub

'*******************************************************************************
' Function Name     : AddToList
' Purpose           : User-defined Add To List
'*******************************************************************************
Public Function AddToList(NodeIndex As Long) As Boolean
  Dim Prnt As Long
  
  If m_tv Is Nothing Then Exit Function           'ignore if the Init has not been called
  If NodeIndex < 1 Then Exit Function             'if 0 or negative
  If Not AddToList(NodeIndex) Then Exit Function  'item is already in the list
  With m_tv
    Prnt = .Nodes(NodeIndex).Parent.Index         'get parent
    If m_LasttvMSParent <> Prnt Then              'if parents do not match
      If Not m_AllowDiffFolders Then              'if different folders not allowed
        Call Clear                                'then clear the current list
        If Not .SelectedItem Is Nothing Then      'if there is something selected
          .SelectedItem.Selected = False          'then unselect it
        End If
        Call AddToList(NodeIndex)                 'add new index back into the list
      End If
    End If
    .Nodes(NodeIndex).BackColor = vbHighlight     'mark it as selected
    If .SelectedItem Is Nothing Then              'if there are no selected items in list
      .Nodes(NodeIndex).Selected = True           'mark as selected
      m_LasttvMSkey = NodeIndex                   'tag item as last-selected item
      m_LasttvMSParent = Prnt                     'get parent index
    End If
  End With
End Function

'*******************************************************************************
' Subroutine Name   : Clear
' Purpose           : Clear selections in the list
' Inputs            : If flag is True, then clear entries in treeview
'                   : If false, then we are assuming the treeview has been cleared
'*******************************************************************************
Public Sub Clear(Optional ResetNodes As Boolean = True)
  Dim Idx As Integer
  
  If m_tv Is Nothing Then Exit Sub       'ignore if the Init has not been called
'
' if the collection contains entries, and the treeview still contains entries, then
' wind down the collection, un-highlighting nodes as they are found. If you have
' already cleared the treeview, or if you pass False as the optional parameter, then
' un-highlighting will be ignored. However, this should only be done if you plan on
' tossing the contents of the TreeView away, as otherwise nodes, will still appear to
' highlighted, though they will no longer exist in the collection.
'
  m_LasttvMSkey = 0                                     'clear last node index
  m_LasttvMSParent = 0                                  'clear parent index
  With m_coltvMultisel
    If .Count Then                                      'anything to do?
      Do While .Count                                   'yes, flush it
        If ResetNodes And CBool(m_tv.Nodes.Count) Then  'if we should reset nodes (not needed when treeview is rebuilt)
          Idx = .Item(1)                                'get the index
          If Idx <= m_tv.Nodes.Count Then m_tv.Nodes(Idx).BackColor = vbWindowBackground
        End If
        .Remove 1                                       'remove the item from the collection
      Loop 'm_coltvMultisel                             'continue while something to do
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : AddNewItem
' Purpose           : Add a new item to the collection. Stuff long as an 11-character
'                   : string so that the list can be sorted
'*******************************************************************************
Private Function AddNewItem(NewItem As Long) As Boolean
  Dim sNew As String
  
  sNew = CStr(NewItem)                          'string version of long
  sNew = String$(11 - Len(sNew), "0")           'make sure 11 char with leading zeros
  On Error Resume Next
  m_coltvMultisel.Add sNew, "K" & CStr(NewItem) 'add to collection
  If Err.Number Then Exit Function              'return False if already exists
  SortCollection m_coltvMultisel                'sort list
  AddNewItem = True                             'return success
End Function

