Attribute VB_Name = "modSortCollection"
Option Explicit
'~modSortCollection.bas;
'Sort a Collection object
'***********************************************************************
' modSortCollection - The SortCollection() will alphabetize a collection
'                     of string values in acending order. The optional
'                     Reserve parameter, if set to TRUE, will sort them in
'                     Decending order.
'EXAMPLE:
'  Dim MyCol As Collection
'  Dim idx As Integer
'
'  Set MyCol = New Collection                'set aside collection
'  With MyCol
'    .Add "Jeff": .Add "David": .Add "Janet" 'fill collection
'    .Add "Gary": .Add "Scott": .Add "Steve"
''display initial list
'    For idx = 1 To .Count                   'display collection
'      Me.Print .Item(idx)
'    Next
'
'    If SortCollection(MyCol) Then           'sort and display sorted
'      Me.Print "----- SORTED LIST ------"
'      For idx = 1 To .Count
'        Me.Print .Item(idx)
'      Next
'    End If
'  End With
'  Set MyCol = Nothing                       'remove collection
'***********************************************************************

Public Function SortCollection(Col As Collection, Optional Reverse As Boolean = False) As Boolean
  Dim str() As String
  Dim Idx As Long, colCount As Long
  Dim IndexLo As Long, IncIndex As Long
  Dim HalfUp As Long, IndexHi As Long
  Dim HalfDown As Long, NumberofItems As Long
  
  NumberofItems = Col.Count
  If NumberofItems <= 0 Then Exit Function
  ReDim str(0 To NumberofItems) As String
  For Idx = 1 To NumberofItems
    str(Idx) = Col(1)
    Col.Remove 1
  Next Idx
  
  HalfDown = NumberofItems              'number of items to sort
  Do While HalfDown \ 2                 'while counter can be halved
    HalfDown = HalfDown \ 2             'back down by 1/2
    HalfUp = NumberofItems - HalfDown   'look in upper half
    IncIndex = 1                        'init index to start of array
    Do While IncIndex <= HalfUp         'do while we can index range
      IndexLo = IncIndex                'set base
      Do
        IndexHi = IndexLo + HalfDown
        If UCase$(str(IndexLo)) > UCase$(str(IndexHi)) Then 'check strings
          str(0) = str(IndexLo)         'swap strings
          str(IndexLo) = str(IndexHi)
          str(IndexHi) = str(0)
          IndexLo = IndexLo - HalfDown  'back up index
        Else
          IncIndex = IncIndex + 1       'else bump counter
          IndexLo = 0                   'allow busting out of 2 loops
          Exit Do
        End If
      Loop While IndexLo > 0            'while more things to check
    Loop
  Loop
'
' check for storage in ascending or decending order
'
  If Reverse Then                         'DESCENDING
    For Idx = NumberofItems To 1 Step -1  'refill sorted collection
      Col.Add str(Idx)
    Next Idx
  Else                                    'ASCENDING
    For Idx = 1 To NumberofItems          'refill sorted collection
      Col.Add str(Idx)
    Next Idx
  End If
  SortCollection = True                   'report success
End Function
