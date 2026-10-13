Attribute VB_Name = "HeaderUtils"
' Return a dictionary mapping header name -> zero-based index
Public Function MapHeaders(ByRef sourceHeaders As Variant) As Object
    Const PROC_NAME As String = "MapHeaders"
    Dim dict As Object
    Dim i As Long

    On Error GoTo ErrorHandler

    Set dict = CreateObject("Scripting.Dictionary")

    If Not IsArray(sourceHeaders) Then
        Set MapHeaders = dict
        Exit Function
    End If

    For i = LBound(sourceHeaders) To UBound(sourceHeaders)
        dict(CStr(sourceHeaders(i))) = i
    Next i

    Set MapHeaders = dict
    Exit Function

ErrorHandler:
    Set dict = Nothing
    Set MapHeaders = Nothing
End Function


'===============================================================================
' GetHeaderIndex
'-------------------------------------------------------------------------------
' Purpose:
'   Return the index of a header name within the array returned by
'   GetQCNumberHeaders. The returned index preserves the array's original
'   lower bound (i.e. it is in the same subscript space as the source array).
'
' When to use:
'   Use this function when you need an index that matches the header array's
'   native subscripts (for example, to reference the original array directly
'   with headers(index)). This is useful when GetQCNumberHeaders returns a
'   1-based array (or any non-zero-based array) and you want to retain that
'   convention.
'
' Parameters:
'   headerName (String) - The header value to search for.
'
' Returns:
'   Integer - The matching subscript value as used by the headers array, or
'             -1 if not found.
'
' Remarks:
'   - The function obtains the headers via GetQCNumberHeaders. That function
'     must return a Variant array with defined LBound/UBound.
'   - If the headers array contains duplicate values, the first match is
'     returned.
'   - The function is O(n) in time complexity (linear search).
'   - Because the return type is Integer, it can overflow for very large
'     arrays (greater than 32,767). If you expect very large arrays,
'     prefer GetHeaderIndex0 (which returns Long).
'===============================================================================

Public Function GetHeaderIndex(headerName As String) As Integer
    Dim headers As Variant, i As Integer
    headers = GetQCNumberHeaders
    For i = LBound(headers) To UBound(headers)
        If headers(i) = headerName Then
            GetHeaderIndex = i
            Exit Function
        End If
    Next i
    GetHeaderIndex = -1
End Function

'===============================================================================
' GetHeaderIndex0
'-------------------------------------------------------------------------------
' Purpose:
'   Return a zero-based index of a header name within the array returned by
'   GetQCNumberHeaders. The returned index is normalized to start at 0 by
'   subtracting the array's LBound.
'
' When to use:
'   Use this function when you need a zero-based index (0..N-1), for example
'   when mapping headers to zero-based collections, creating zero-based
'   lookup tables, or when interfacing with libraries/APIs that expect
'   zero-based indices.
'
' Parameters:
'   headerName (String) - The header value to search for.
'
' Returns:
'   Long - The zero-based index (0..N-1) of the matching header, or
'          -1 if not found.
'
' Remarks:
'   - The function obtains the headers via GetQCNumberHeaders. That function
'     must return a Variant array with defined LBound/UBound.
'   - If the headers array is already zero-based, this function returns the
'     same index as GetHeaderIndex; otherwise it normalizes the index.
'   - If the headers array contains duplicate values, the first match is
'     returned.
'   - The function is O(n) in time complexity (linear search).
'   - The return type is Long to avoid integer overflow for large arrays.
'===============================================================================

Public Function GetHeaderIndex0(ByVal headerName As String) As Long
    Dim headers As Variant, i As Long, lb As Long
    headers = GetQCNumberHeaders
    lb = LBound(headers)
    For i = lb To UBound(headers)
        If headers(i) = headerName Then
            GetHeaderIndex0 = i - lb
            Exit Function
        End If
    Next i
    GetHeaderIndex0 = -1
End Function