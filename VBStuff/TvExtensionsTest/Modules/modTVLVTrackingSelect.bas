Attribute VB_Name = "modTVLVTrackingSelect"
Option Explicit
'~modTVLVTrackingSelect.bas;
'Privide SelectTracking ability to TreeView and ListView Controls
'*******************************************************************************
' modTVLVTrackingSelect - These subroutines allow performing TrackingSelect on
'                         TreeView and ListView controls. Call the appropriate
'                         subroutine from your MouseMove event. Support is added
'                         so that if you move back over an item that is already
'                         selected, that it will be unselected.
' The subroutines are:
'TreeViewTrackingSelect(): Perform TrackingSelect on a TreeView Control.
'ListViewTrackingSelect(): Perform TrackingSelect on a ListView Control.
'
'EXAMPLES:
'Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'  If Shift = vbCtrlMask And Button = 0 Then   'user holding Ctrl down and no button
'    TreeViewTrackingSelect TreeView1, x, y    'select/deselect items
'  End If
'End Sub
'
'Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'  If Shift = vbCtrlMask And Button = 1 Then   'user holding Ctrl down and no button
'    ListViewTrackingSelect ListView1, x, y    'select/deselect items
'  End If
'End Sub
'*******************************************************************************

'*******************************************************************************
' Subroutine Name   : TreeViewTrackingSelect
' Purpose           : TrackingSelect for TreeView.
'*******************************************************************************
Public Sub TreeViewTrackingSelect(TreeView As TreeView, x As Single, y As Single)
  Static LastNode As Node
  Dim AnyNode As Node
  
  With TreeView
    Set AnyNode = .HitTest(x, y)        'get node we are over
    If Not AnyNode Is Nothing Then      'something there?
      If Not AnyNode Is LastNode Then
        Set .DropHighlight = AnyNode    'yes, select it
        .DropHighlight.Selected = Not .DropHighlight.Selected 'flip selection
      End If
    End If
    Set LastNode = AnyNode
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : ListViewTrackingSelect
' Purpose           : TrackingSelect for ListView
'*******************************************************************************
Public Sub ListViewTrackingSelect(ListView As ListView, x As Single, y As Single)
  Static LastItem As ListItem
  Dim AnyItem As ListItem
  
  With ListView
    Set AnyItem = .HitTest(x, y)        'get item we are over
    If Not AnyItem Is Nothing Then
      Set .DropHighlight = AnyItem      'yes, select it
      If Not AnyItem Is LastItem Then
        Set .DropHighlight = AnyItem    'yes, select it
        .DropHighlight.Selected = Not .DropHighlight.Selected 'flip selection
      End If
    End If
    Set LastItem = AnyItem              'save last item
  End With
End Sub

