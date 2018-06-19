VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DopeArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Arr As Variant
Private Enum FuncType
    Minf
    Maxf
    Avgf
End Enum
Private Sub Class_Initialize()
'-------------------------------------------------------------
'PURPOSE:set private array variable to nothing
'-------------------------------------------------------------
    Set Arr = Nothing
End Sub
Property Get ArrayVals()
    If Me.IsEmpty Then
        ArrayVals = Null
        Exit Property
    End If
    Set ArrayVals = Arr
End Property

Property Get Count()
'-------------------------------------------------------------
'PURPOSE: get non-zero-index-based count of items in array
'-------------------------------------------------------------
    Count = F_CountArrayItems(Arr)
End Property
Property Get Max()
'-------------------------------------------------------------
'PURPOSE: get max value of items in array. will error out if
'not all items in array are numeric.
'-------------------------------------------------------------
    Max = F_GetFuncVal(Maxf)
End Property
Property Get Min()
'-------------------------------------------------------------
'PURPOSE: get min value of items in array. will error out if
'not all items in array are numeric.
'-------------------------------------------------------------
    Min = F_GetFuncVal(Minf)
End Property
Property Get Avg()
'-------------------------------------------------------------
'PURPOSE: get avg value of items in array. will error out if
'not all items in array are numeric.
'-------------------------------------------------------------
    Avg = F_GetFuncVal(Avgf)
End Property
Property Get IsEmpty()
'-------------------------------------------------------------
'PURPOSE: boolean to advise if the array object is empty or
'not.
'-------------------------------------------------------------
    IsEmpty = Not IsArray(Arr)
End Property
Private Function F_CountArrayItems(tgArr As Variant)
'-------------------------------------------------------------
'PURPOSE: counts items in array that gets passed in. If param
'that gets passed in is not an array it'll return a zero.
'-------------------------------------------------------------
Dim Itm As Variant
Dim Cntr As Long
    'return zero and exit if not array
    If Not IsArray(tgArr) Then
        F_CountArrayItems = 0
        Exit Function
    End If
    'count how many items in array
    For Each Itm In tgArr
        Cntr = Cntr + 1
    Next Itm
    F_CountArrayItems = Cntr
End Function
Public Sub AddNew(ItemOrArray As Variant)
'-------------------------------------------------------------
'PURPOSE: Add items to the array. Will accept a single item
'or an array of items. Should only work with single dimen-
'sion arrays.
'-------------------------------------------------------------
Dim Itm As Variant
    'test to see if item passed is array or single item
    Select Case IsArray(ItemOrArray)
        Case False
            'add single item
             Call S_AddToArray(ItemOrArray, Arr)
        Case True
            'add multiple items
            For Each Itm In ItemOrArray
                Call S_AddToArray(Itm, Arr)
            Next Itm
    End Select
End Sub
Private Function F_GetFuncVal(Func As FuncType)
'-------------------------------------------------------------
'PURPOSE: this private function works in conjuction with the
'min, max, and average properties. It was actually does the
'work to discern the min, max, or avg.
'-------------------------------------------------------------
Dim ErrMsg As String
Dim Itm As Variant
Dim FuncVal As Variant
Dim FuncNm As String
    'test that there are values
    If Me.IsEmpty Then
        F_GetFuncVal = Null
        Exit Function
    End If
    'set up func name
    Select Case Func
        Case FuncType.Minf
            FuncNm = "min"
        Case FuncType.Maxf
            FuncNm = "max"
        Case FuncType.Avgf
            FuncNm = "avg"
    End Select
    'set FuncVal variable to first item of array for max and min calculations
    If Not Func = FuncType.Avgf Then FuncVal = Arr(0)
    'loop throu each item in array
    For Each Itm In Arr
        'test to make sure it is a number
        If Not IsNumeric(Itm) Then
            ErrMsg = "Ya gotsss to have numbers for the " & FuncNm & " calculation. " & vbNewLine & _
                     "'" & Itm & "' isn't a number and kinda wont cut it."
            Err.Raise vbObjectError, "DopeArray.GetFuncVal", ErrMsg
        End If
        'test function logic
        Select Case Func
            Case FuncType.Minf
                If Itm < FuncVal Then FuncVal = Itm
            Case FuncType.Maxf
                If Itm > FuncVal Then FuncVal = Itm
            Case FuncType.Avgf
                FuncVal = FuncVal + Itm
        End Select
    Next Itm
    'return FuncVal value
    Select Case Func
        Case FuncType.Avgf
            'get count of items in array
            Itm = F_CountArrayItems(Arr)
            'get average
            F_GetFuncVal = (FuncVal / Itm)
        Case Else
            F_GetFuncVal = FuncVal
    End Select
    

End Function
Public Sub PrintCount()
'-------------------------------------------------------------
'PURPOSE: quickly get a count of how many items exist. This
'is nicer when compare to having to write out
'"debug.print arrobj.count" you can instead just type out
'"arrobj.printcout" - minor advantage? yes - but that is what
'makes life rich!
'-------------------------------------------------------------
    Debug.Print Me.Count
End Sub
Public Sub RemoveDups()
'-------------------------------------------------------------
'PURPOSE: remove duplicate items from array object. this is
'done by adding all unique values to a dictionary when they
'first appeared in the array. Subsequent duplicate apperances
'will be ignored.
'-------------------------------------------------------------
Dim Dct As New Scripting.Dictionary
Dim Itm As Variant
    'test that array is not empty
    If Me.IsEmpty Then Exit Sub
    'loop through throuth each item in array
    For Each Itm In Arr
        'add to dictionary if not already there
        If Not Dct.Exists(Itm) Then Dct.Add Itm, Null
    Next Itm
    'reset array
    Set Arr = Nothing
    'add keys to private array variable
    Arr = Dct.Keys
End Sub
Public Function Exists(Item)
'-------------------------------------------------------------
'PURPOSE: test whether the item passed in is in the private
'array varibale. This test is case-insensitive.
'-------------------------------------------------------------
Dim Itm As Variant
    'return null and exit if the array is empty
    If Me.IsEmpty Then
        Exists = Null
        Exit Function
    End If
    Exists = F_IsInArray(Item, Arr)
    'if we got to this point than we didn't find a match
    Exists = False
End Function
Private Function F_IsInArray(Item As Variant, tgArr As Variant)
'-------------------------------------------------------------
'PURPOSE: this is the function that drives the test whether
'an item exists or not.
'-------------------------------------------------------------
Dim Itm As Variant
    'assure array was passed in
    If Not IsArray(tgArr) Then
        F_IsInArray = Null
        Exit Function
    End If
    'look for exact match
    For Each Itm In tgArr
        F_IsInArray = (Itm = Item)
        If F_IsInArray Then Exit Function
    Next Itm
    
End Function

Public Sub ReplaceNulls(Optional ReplaceVal As String = vbNullString)
'-------------------------------------------------------------
'PURPOSE: This sub replaces a true null value as recognized
'by the IDE and replaces it with the replaceval passed in.
'-------------------------------------------------------------
Dim Itm As Variant
Dim NewArr As Variant
    If Me.IsEmpty Then Exit Sub
    'loop through each item
    For Each Itm In Arr
        'replace null values
        If IsNull(Itm) Then Itm = ReplaceVal
        'add values to new array
        Call S_AddToArray(Itm, NewArr)
    Next Itm
    'set arr to newarr values
    Arr = NewArr
End Sub
Public Function CompareTwoArrays(GoodArr, TestArr, Optional ArrayType = "items")
'-------------------------------------------------------------
'PURPOSE: this function returns true if all the items in
'testarr match the items in good array. However if testarr has
'items missing from goodarr or if testarr has extra items not
'found in goodarray then it'll return a message detailing what
'is missing (if applicable), what is extra (if applicable),
'and what are all the expected items in the GoodArr.
'the arrayvals property can be used to pass in to the good or
'test arrays if necessary.
'-------------------------------------------------------------
Dim Itm As Variant
Dim Missing As Variant
Dim CntMissing As Long
Dim CntExtra As Long
Dim Extra As Variant
Dim ErrMsg As Variant
Const StringJn As String = vbNewLine & vbTab & "-- "
    'test for items missing from test arr that should be there according to goodarr
    For Each Itm In GoodArr
        If Not F_IsInArray(Itm, TestArr) Then Call S_AddToArray(Itm, Missing)
    Next Itm
    'test for items that are extra in test arr that are not required per good array
    For Each Itm In TestArr
        If Not F_IsInArray(Itm, GoodArr) Then Call S_AddToArray(Itm, Extra)
    Next Itm
    'get counts of items in missing and extra arrays
    CntMissing = F_CountArrayItems(Missing)
    CntExtra = F_CountArrayItems(Extra)
    Select Case (CntMissing + CntExtra) = 0
        Case True
            'return true indicating that arrays match
            CompareTwoArrays = True
        Case False
            'prep messages for missing items and extra items where applicable
            If IsArray(Missing) Then
                Itm = "TestArr missing " & CntMissing & " below " & ArrayType & " that are in GoodArr:" & _
                      StringJn & Join(Missing, StringJn)
                Call S_AddToArray(Itm, ErrMsg)
            End If
            If IsArray(Extra) Then
                'add a new line
                Call S_AddToArray(vbNewLine, ErrMsg)
                Itm = "TestArr has " & CntExtra & " extra " & ArrayType & " not in GoodArr:" & _
                      StringJn & Join(Extra, StringJn)
                Call S_AddToArray(Itm, ErrMsg)
            End If
            'add a new line
            Call S_AddToArray(vbNewLine, ErrMsg)
            'add expected list of items
            Itm = F_CountArrayItems(GoodArr)
            Itm = "Exptected " & Itm & " " & ArrayType & " in GoodArr:" & _
                  StringJn & Join(GoodArr, StringJn)
            Call S_AddToArray(Itm, ErrMsg)
            'return joined messages
            CompareTwoArrays = Join(ErrMsg, vbNullString)
    End Select
End Function

Private Sub S_AddToArray(Item As Variant, tgArr As Variant)
'-------------------------------------------------------------
'PURPOSE: takes an item and and an array variable. If the
'array variable passed is not an array then it is converted
'into an array with the item being the first item in the array.
'if it is alreay an array the the array's size is readjusted'
'and the item passed is added to the end of the array.
'-------------------------------------------------------------
Dim LstPos As Long
    Select Case IsArray(tgArr)
        Case False
            'if tgarr not already array then simply add the item and exit sub
            tgArr = Array(Item)
        Case True
            'capture new upper bound
            LstPos = UBound(tgArr) + 1
            'resize array while keeping all data
            ReDim Preserve tgArr(0 To LstPos)
            'add item to last position
            tgArr(LstPos) = Item
    End Select
End Sub
