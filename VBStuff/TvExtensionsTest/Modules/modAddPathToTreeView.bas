Attribute VB_Name = "modAddPathToTreeView"
Option Explicit
'~modAddPathToTreeView.bas;
'Add a series of chained nodes to a TreeView as a path
'********************************************************************************
' modAddPathToTreeView - The AddPathToTreeView() subroutine will add a chained
'                        series of tree node as a path specification, where the
'                        leftmost item is at the root of the TreeView, and each
'                        succeeding item, separated by a backslash "\", are
'                        children, grandchildren, etc. The depth of the items
'                        are stored as their Key value. A leading slash, if added,
'                        will be ignored in the path.
'EXAMPLE:
'  AddPathToTreeView TreeView1, "RootLevel\Child1\Child2\Child3\Child4"
''' The above command will add the following items and keys:
''' Rootlevel    \RootLevel
''' Child1       \RootLevel\Child1
''' Child2       \RootLevel\Child1\Child2
''' Child3       \RootLevel\Child1\Child2\Child3
''' Child4       \RootLevel\Child1\Child2\Child3\Child4
'
' NOTE: Obviously, Microsoft Windows Common Controls 6.0 (MSCOMCTL.OCX) is used.
'********************************************************************************

Public Sub AddPathToTreeView(TV As TreeView, Path As String)
  Dim VList() As Variant                                'nodes list
  Dim TPath As String                                   'local copy of path
  Dim PathLen As Integer                                'length of temp string
  Dim PathItem As String                                'parent path
  Dim NewItem As String                                 'child node text
  Dim c As String * 1                                   'test character
  Dim i As Integer                                      'loop counter
  
  On Error Resume Next                                  'skip duplicated nodes
  TPath = Trim$(Path)                                   'grab path
  If Left$(TPath, 1) = "\" Then TPath = Mid$(TPath, 2)  'strip any leading slash
  If Right$(TPath, 1) <> "\" Then TPath = TPath & "\"   'require a trailing slash
  PathLen = Len(TPath)                                  'get data length for loop
  NewItem = ""
'
' build node path. Usually inevitable collision duplication of nodes are
'                  protected by the On Error Resume Next command
'
  For i = 1 To PathLen
    c = Mid$(TPath, i, 1)                               'get a character
    If c = "\" Then                                     'backslash?
      If Len(PathItem) Then                             'yes. If data, then add child
        TV.Nodes.Add PathItem, tvwChild, PathItem & "\" & NewItem, NewItem
        PathItem = PathItem & "\" & NewItem             'build parent mode
      Else
        TV.Nodes.Add , , "\" & NewItem, NewItem         'set root item
        PathItem = "\" & NewItem                        'build parent node
      End If
      NewItem = ""                                      'reset for next child
    Else
      NewItem = NewItem & c                             'build child node
    End If
  Next i                                                'scan all characters
End Sub
