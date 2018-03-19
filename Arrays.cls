VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 0 ' Default

Private Declare Sub GetMem2 Lib "msvbvm60" (src As Any, dest As Any)

Public Sub sort(ByRef arr As Variant, Optional sortOrder As sortOrder = ascending)

    Variants.verifyArray arr
    
    Dim length As Integer
    length = UBound(arr) - LBound(arr) + 1
    
    If length <= 0 Then
        Exit Sub
    End If

    If Not Variants.isComparable(arr(LBound(arr))) Then
        Err.Raise E_ARGUMENTOUTOFRANGE, "Arrays.Sort()", "Array must be Variants.IsComparable() for sorting to work."
    End If

    Dim isRef As Boolean
    isRef = IsObject(arr(LBound(arr)))

    Dim a, b As Variant
    a = arr
    ReDim b(LBound(a) To UBound(a))
    Dim aHasTheData As Boolean
    aHasTheData = True

    Dim chunkSize As Integer
    chunkSize = 1
    While chunkSize < length
        Dim sortPos As Integer
        sortPos = LBound(a)
        While sortPos <= UBound(a)
            If aHasTheData Then
                Merge a, b, sortPos, Math.min(sortPos + chunkSize, UBound(a) + 1), Math.min(sortPos + 2 * chunkSize, UBound(a) + 1), isRef, sortOrder
            Else
                Merge b, a, sortPos, Math.min(sortPos + chunkSize, UBound(a) + 1), Math.min(sortPos + 2 * chunkSize, UBound(a) + 1), isRef, sortOrder
            End If
            sortPos = sortPos + 2 * chunkSize
        Wend
        aHasTheData = Not aHasTheData
        chunkSize = chunkSize * 2
    Wend
    
    ' Copy the result over to the input array.
    Dim sourceArray As Variant
    sourceArray = IIf(aHasTheData, a, b)
    Dim runner As Integer
    For runner = LBound(sourceArray) To UBound(sourceArray)
        If isRef Then
            Set arr(runner) = sourceArray(runner)
        Else
            arr(runner) = sourceArray(runner)
        End If
    Next
End Sub

Private Sub Merge(source As Variant, ByRef target As Variant, leftStart As Integer, rightStart As Integer, rightEnd As Integer, isRef As Boolean, sortOrder As sortOrder)
    
    Dim targetRunner, lRunner, rRunner As Integer
    lRunner = leftStart
    rRunner = rightStart
    
    For targetRunner = leftStart To rightEnd - 1
        ' If we have stuff in the left chunk left
        ' and we either don't have stuff in the right chunk anymore
        ' or left is smaller or equal to right (We want to be stable!)
        ' -> then take the left element.
        Dim takeLeft As Boolean
        takeLeft = False
        If lRunner < rightStart Then
            If rRunner >= rightEnd Then
                takeLeft = True
            Else
                If sortOrder = ascending And Math.cmp(source(lRunner), source(rRunner)) <= 0 _
                        Or sortOrder = descending And Math.cmp(source(lRunner), source(rRunner)) >= 0 Then
                    takeLeft = True
                End If
            End If
        End If
        
        If takeLeft Then
            ' -> Take the left element.
            If isRef Then
                Set target(targetRunner) = source(lRunner)
            Else
                target(targetRunner) = source(lRunner)
            End If
            
            lRunner = lRunner + 1
        Else
            ' -> Take the right element.
            If isRef Then
                Set target(targetRunner) = source(rRunner)
            Else
                target(targetRunner) = source(rRunner)
            End If
            
            rRunner = rRunner + 1
        End If
    Next
End Sub

Public Function emptyVariantArray() As Variant()
    emptyVariantArray = Array()
End Function

Public Function emptyIntegerArray() As Integer()
    ' Taken from http://stackoverflow.com/a/21290864/1975049
    Dim i() As Integer
    Dim v As Variant
    v = Array()
    
    Dim NewTypeCode As Integer
    NewTypeCode = vbArray Or vbInteger
    GetMem2 NewTypeCode, v
    i = v
    emptyIntegerArray = i
End Function

Public Function emptyByteArray() As Byte()
    ' Taken from http://stackoverflow.com/a/21290864/1975049
    Dim b() As Byte
    Dim v As Variant
    v = Array()
    
    Dim NewTypeCode As Integer
    NewTypeCode = vbArray Or vbByte
    GetMem2 NewTypeCode, v
    b = v
    emptyByteArray = b
End Function

Public Function elems(arr As Variant) As Long
    elems = UBound(arr) - LBound(arr) + 1
End Function

Public Function toVariantArray(arr As Variant) As Variant()
    Dim varArr() As Variant
    
    If elems(arr) = 0 Then
        toVariantArray = emptyVariantArray
    Else
        ReDim varArr(LBound(arr) To UBound(arr))
        Dim runner As Long
        For runner = LBound(arr) To UBound(arr)
            If IsObject(arr(runner)) Then
                Set varArr(runner) = arr(runner)
            Else
                varArr(runner) = arr(runner)
            End If
        Next
        
        toVariantArray = varArr
    End If
End Function


