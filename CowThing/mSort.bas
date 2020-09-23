Attribute VB_Name = "mSort"
Option Explicit

' This module is a version of C.A.R Hoare's Quick Sort algorithm
' extended with TriMedian and InsertionSort by Denis Ahrens
' with all the tips from Robert Sedgewick (Algorithms in C++)
' Author: James Gosling
' Author: Kevin A. Smith
' Visual Basic Adaptation: Peter Wilson (www.midar.com.au)
'
' Peter: This is probably the fastest sort algorithm I've ever seen. Blistering Speed!
' It was originally written for C, and I think I'm the first non-Brazilian to
' travel through time and convert it to Visual Basic!
'
' The code for this sort algorithim is not as easy to understand as a simple bubble-sort,
' but then again, you'd have to be brain-dead to even contemplate the bubble-sort. In fact,
' I'm sorry I even mentioned it!
'
' Hopefully this URL is still valid by the time you read this.
'   http://www.cs.ubc.ca/spider/harrison/Java/sorting-demo.html
'
' If you were going to use VB's internal sort routine (namely the listbox with it's sort
' property set to true, ok I know this is bad, but VB does not have any internal sort routines)
' it would take 1.8 seconds to load 10,000 random numbers and have
' VB automatically sort this list, (then you can read back out the sorted results).
' This "FastQuickSort" routine takes only 0.3 seconds to sort the exact same random list. It's
' also quicker because it doesn't need to spend time updating the contents of a list box.
' However, even if you do decided to sort first, then place results into a list box,
' this sort routine is still quicker than VB's own internal sort routine. See below...
'
' Test 1
'   Place 30,000 random numbers into a ListBox with it's sort property set to true.
'              VB takes 6379ms to sort this list (including displaying the results in the listbox)
'
' Test 2
'   Place 30,000 random numbers into an Array. Use a listbox with it's sort property set to false.
'   FastQuickSort takes 3686ms to sort this list (including displaying the results in the listbox)
'
' Bottome Line: I have just programmed a faster sorting routine than Microsoft, and I did it in VB, in one weekend.
'
' As you can see, you should never again use ListBoxes for sorting anything! Not even themselves! ha ha.
' (Results obtained on my on my DELL XPS T500 PentiumIII 500MHz)

Private Const mc_ModuleName As String = "mSort"

Public Sub FastQSort(ArrayIn As Variant)
    
    ' The Fast Quick Sort is actually made up from the following two sort routines.
    ' (see comments at the start of this module)
    
    ' Quick Presort
    Call QuickSort(ArrayIn, 0, UBound(ArrayIn))
    
    ' Final sort.
    Call InsertionSort(ArrayIn, 0, UBound(ArrayIn))
    
End Sub

Private Sub InsertionSort(ArrayIn As Variant, lngLow As Long, lngHigh As Long)
    
    ' (Please read comments at beginning of this module)
    
    Dim lngN As Long
    Dim lngJ As Long
    Dim lngValue1 As Long
    Dim lngValue2 As Long
    
    lngLow = lngLow + 1
    
    For lngN = lngLow To lngHigh
        lngValue1 = ArrayIn(lngN, 0)
        lngValue2 = ArrayIn(lngN, 1) ' << This line isn't required for a 1-dimensional array sort.
        
        lngJ = lngN
        While (lngJ > lngLow) And (ArrayIn(lngJ - 1, 0) > lngValue1)
            ArrayIn(lngJ, 0) = ArrayIn(lngJ - 1, 0)
            ArrayIn(lngJ, 1) = ArrayIn(lngJ - 1, 1) ' << This line isn't required for a 1-dimensional array sort.
            lngJ = lngJ - 1
        Wend
        ArrayIn(lngJ, 0) = lngValue1
        ArrayIn(lngJ, 1) = lngValue2 ' << This line isn't required for a 1-dimensional array sort.
        
    Next lngN
    
End Sub

Private Sub QuickSort(ArrayIn As Variant, lngLowerBoundry As Long, lngUpperBoundry As Long)

    ' (Please read comments at beginning of this module)
    On Error GoTo errTrap
    
    Dim lngM As Long
    Dim lngMidPoint As Long
    Dim lngJ As Long
    Dim lngValue As Long
    
    lngM = 4
    
    If ((lngUpperBoundry - lngLowerBoundry) > lngM) Then
    
        lngMidPoint = (lngUpperBoundry + lngLowerBoundry) \ 2
        
        If ArrayIn(lngLowerBoundry, 0) > ArrayIn(lngMidPoint, 0) Then Call Swap(ArrayIn, lngLowerBoundry, lngMidPoint)
        If ArrayIn(lngLowerBoundry, 0) > ArrayIn(lngUpperBoundry, 0) Then Call Swap(ArrayIn, lngLowerBoundry, lngUpperBoundry)
        If ArrayIn(lngMidPoint, 0) > ArrayIn(lngUpperBoundry, 0) Then Call Swap(ArrayIn, lngMidPoint, lngUpperBoundry)
        
        lngJ = lngUpperBoundry - 1
        Call Swap(ArrayIn, lngMidPoint, lngJ)
        lngMidPoint = lngLowerBoundry
        lngValue = ArrayIn(lngJ, 0)
        
        Do
            lngMidPoint = lngMidPoint + 1
            While ArrayIn(lngMidPoint, 0) < lngValue
                lngMidPoint = lngMidPoint + 1
            Wend
            
            lngJ = lngJ - 1
            While ArrayIn(lngJ, 0) > lngValue
                lngJ = lngJ - 1
            Wend
            
            If lngJ < lngMidPoint Then Exit Do
            Call Swap(ArrayIn, lngMidPoint, lngJ)
        Loop
        
        Call Swap(ArrayIn, lngMidPoint, lngUpperBoundry - 1)
        Call QuickSort(ArrayIn, lngLowerBoundry, lngJ)
        Call QuickSort(ArrayIn, lngMidPoint + 1, lngUpperBoundry)
        
    End If
    
    Exit Sub
errTrap:
    ' Log error message to logfile (works only for compiled applications - ie. EXE files)
    App.LogEvent Now & ", " & mc_ModuleName & "." & Err.Source & ", " & Err.Number & ", " & Err.Description
    
End Sub

Private Sub Swap(ArrayIn As Variant, i As Long, j As Long)
    
    ' (Please read comments at beginning of this module)
    
    Dim sngAverageZDistance As Single
    Dim lngTempStorage As Long
    
    ' Remember old values (because "part 1" will destroy one of the array's value pair)
    sngAverageZDistance = ArrayIn(i, 0)
    lngTempStorage = ArrayIn(i, 1) ' << This line isn't required for a 1-dimensional array sort.
    
    ' Swap values - part 1
    ArrayIn(i, 0) = ArrayIn(j, 0)
    ArrayIn(i, 1) = ArrayIn(j, 1) ' << This line isn't required for a 1-dimensional array sort.
    
    ' Swap values - part 2
    ArrayIn(j, 0) = sngAverageZDistance
    ArrayIn(j, 1) = lngTempStorage ' << This line isn't required for a 1-dimensional array sort.
        
End Sub


