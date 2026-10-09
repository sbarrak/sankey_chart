VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SkDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'  Minimal Scripting.Dictionary replacement (Excel for Mac has no Scripting runtime).
'  Case-insensitive string keys, insertion order kept. Supports Exists, Add, Remove,
'  Keys, Count, CompareMode (ignored) and a default Item property: d(key) = value.
'==============================================================================
Option Explicit

Private mKeys() As String
Private mVals() As Variant
Private mN As Long
Private mIdx As Collection

Private Sub Class_Initialize()
    mN = 0
    ReDim mKeys(0 To 15)
    ReDim mVals(0 To 15)
    Set mIdx = New Collection
End Sub

Private Function Find(ByVal k As String) As Long
    On Error Resume Next
    Find = mIdx("k" & k)
    If Err.Number <> 0 Then Find = 0
End Function

Public Function Exists(ByVal k As Variant) As Boolean
    Exists = (Find(CStr(k)) > 0)
End Function

Public Property Get Count() As Long
    Count = mN
End Property

Public Property Let CompareMode(ByVal v As Long)
    ' keys are always compared case-insensitively
End Property

Public Property Get Item(ByVal k As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim i As Long
    i = Find(CStr(k))
    If i = 0 Then Exit Property
    If IsObject(mVals(i)) Then Set Item = mVals(i) Else Item = mVals(i)
End Property

Public Property Let Item(ByVal k As Variant, ByVal v As Variant)
    Dim i As Long
    i = Find(CStr(k))
    If i = 0 Then Add k, v Else mVals(i) = v
End Property

Public Sub Add(ByVal k As Variant, ByVal v As Variant)
    Dim s As String
    s = CStr(k)
    If Find(s) > 0 Then Err.Raise 457, "SkDict", "This key is already associated with an element of this collection"
    If mN + 1 > UBound(mKeys) Then
        ReDim Preserve mKeys(0 To UBound(mKeys) * 2 + 1)
        ReDim Preserve mVals(0 To UBound(mVals) * 2 + 1)
    End If
    mN = mN + 1
    mKeys(mN) = s
    mVals(mN) = v
    mIdx.Add mN, "k" & s
End Sub

Public Sub Remove(ByVal k As Variant)
    Dim i As Long, j As Long
    i = Find(CStr(k))
    If i = 0 Then Exit Sub
    For j = i To mN - 1
        mKeys(j) = mKeys(j + 1)
        mVals(j) = mVals(j + 1)
    Next j
    mN = mN - 1
    Set mIdx = New Collection
    For j = 1 To mN
        mIdx.Add j, "k" & mKeys(j)
    Next j
End Sub

Public Sub RemoveAll()
    Class_Initialize
End Sub

Public Function Keys() As Variant            ' snapshot array, safe to Remove while looping
    Dim a() As Variant, i As Long
    If mN = 0 Then
        Keys = Array()
        Exit Function
    End If
    ReDim a(0 To mN - 1)
    For i = 1 To mN
        a(i - 1) = mKeys(i)
    Next i
    Keys = a
End Function
