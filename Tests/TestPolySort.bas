Attribute VB_Name = "TestPolySort"
'@IgnoreModule VariableNotAssigned, UnassignedVariableUsage, LineLabelNotUsed, VariableNotUsed, AssignmentNotUsed, FunctionReturnValueDiscarded, UseMeaningfulName, MultipleDeclarations
'@TestModule
'@Folder("Tests")
'@ModuleDescription("Rubberduck unit tests for PolySort, one case per test.")
Option Explicit
Option Private Module

''
' Comments
''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Rubberduck unit tests for PolySort
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2026.11.11
'
' One case per test. The expected results come from counting the values, a
' check of every neighbouring pair or a linear search. CompareCustom is
' module state of PolySort: the tests that set it reset it, and so does
' TestCleanup.

#Const LateBinding = False

#If LateBinding Then
Private Assert As Object
#Else
Private Assert As AssertClass
#End If


''
' Test module set-up
''

'@ModuleInitialize
Private Sub ModuleInitialize()
#If LateBinding Then
    Set Assert = CreateObject("Rubberduck.AssertClass")
#Else
    Set Assert = New AssertClass
#End If
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub


'@TestCleanup
Private Sub TestCleanup()
    Set CompareCustom = Nothing
End Sub


''
' Helpers
''

'@Description("Returns an array (1 To itemCount) of random whole numbers 0 To distinct - 1.")
Private Function RandomValues(ByVal itemCount As Long, ByVal distinct As Long) As Variant()
    Rnd -1
    Randomize 13
    Dim Values() As Variant: ReDim Values(1 To itemCount)
    Dim i As Long
    For i = 1 To itemCount
        Values(i) = Int(Rnd * distinct)
    Next
    RandomValues = Values
End Function


'@Description("Returns True if every element is >= its predecessor, directly or through an index.")
Private Function Ascending(ByRef arr As Variant, Optional ByRef idx As Variant) As Boolean
    Dim i As Long
    If IsMissing(idx) Then
        For i = LBound(arr) + 1 To UBound(arr)
            If arr(i) < arr(i - 1) Then Exit Function
        Next
    Else
        For i = LBound(idx) + 1 To UBound(idx)
            If arr(idx(i)) < arr(idx(i - 1)) Then Exit Function
        Next
    End If
    Ascending = True
End Function


'@Description("Returns the first position of a value by a linear search.")
Private Function FirstPosition(ByRef arr As Variant, ByVal value As Variant, Optional ByRef idx As Variant) As Long
    Dim i As Long
    If IsMissing(idx) Then
        For i = LBound(arr) To UBound(arr)
            If arr(i) = value Then FirstPosition = i: Exit Function
        Next
    Else
        For i = LBound(idx) To UBound(idx)
            If arr(idx(i)) = value Then FirstPosition = i: Exit Function
        Next
    End If
End Function


'@Description("Returns a Person.")
Private Function NewPerson(ByVal givenName As String, ByVal familyName As String) As Person
    Set NewPerson = New Person
    NewPerson.FirstName = givenName
    NewPerson.Lastname = familyName
End Function


'@Description("Returns three persons, not in the order of their last names (De Vries, Jansen, Smit).")
Private Function People() As Variant()
    Dim Persons(1 To 3) As Variant
    Set Persons(1) = NewPerson("Anna", "Smit")
    Set Persons(2) = NewPerson("Bert", "Jansen")
    Set Persons(3) = NewPerson("Carla", "de Vries")
    People = Persons
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("Sort")
Private Sub Sort_RandomValues()
    On Error GoTo TestFail

    Dim Values() As Variant
    Values = RandomValues(5000, 1000000)
    Sort Values
    Assert.IsTrue Ascending(Values)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Sort")
Private Sub Sort_AllEqual()
    On Error GoTo TestFail

    Dim Values() As Variant
    Values = RandomValues(20000, 1)
    Sort Values
    Assert.IsTrue Ascending(Values)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Sort")
Private Sub Sort_ThreeValues()
    On Error GoTo TestFail

    Dim Values() As Variant
    Values = RandomValues(20000, 3)
    Sort Values
    Assert.IsTrue Ascending(Values)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Sort")
Private Sub Sort_Descending()
    On Error GoTo TestFail

    Dim Values() As Variant
    Values = RandomValues(5000, 3)
    Sort Values, asc:=False
    Assert.AreEqual 2#, Values(1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Sort")
Private Sub Sort_StringsBinaryByDefault()
    On Error GoTo TestFail

    Dim Words() As Variant: Words = Array("b", "A", "c", "B", "a")
    Sort Words
    Assert.AreEqual "A,B,a,b,c", Join(Words, ",")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Sort")
Private Sub Sort_StringsText()
    On Error GoTo TestFail

    Dim Words() As Variant: Words = Array("b", "A", "c", "B", "a")
    Sort Words, method:=vbTextCompare
    Assert.AreEqual "a", LCase$(Words(0))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Sort")
Private Sub Sort_StringsTextIsSorted()
    On Error GoTo TestFail

    Dim Words() As Variant: Words = Array("b", "A", "c", "B", "a")
    Sort Words, method:=vbTextCompare
    Assert.IsTrue IsSorted(Words, method:=vbTextCompare)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Sort")
Private Sub Sort_VariantDates()
    On Error GoTo TestFail

    Dim Days() As Variant: Days = Array(#3/1/2026#, #1/1/2026#, #2/1/2026#)
    Sort Days
    Assert.AreEqual #1/1/2026#, Days(0)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Sort")
Private Sub Sort_TypedDates()
    On Error GoTo TestFail

    Dim Days(0 To 2) As Date
    Days(0) = #3/1/2026#: Days(1) = #1/1/2026#: Days(2) = #2/1/2026#
    Sort Days
    Assert.AreEqual #1/1/2026#, Days(0)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Sort")
Private Sub Sort_ObjectsWithCustomCompare()
    On Error GoTo TestFail

    Dim Persons() As Variant: Persons = People()
    Set CompareCustom = New Person
    Sort Persons
    Assert.AreEqual "Jansen", Persons(2).Lastname
    Set CompareCustom = Nothing

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Sort")
Private Sub Sort_ObjectsWithoutCompareRaisesError()
    On Error GoTo TestFail

    Dim Persons() As Variant: Persons = People()
    Dim ErrNo As Long
    On Error Resume Next
    Sort Persons
    ErrNo = Err.Number: Err.Clear
    On Error GoTo TestFail
    Assert.AreEqual 17&, ErrNo

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Sort")
Private Sub SortByIndex_ThreeValues()
    On Error GoTo TestFail

    Dim Values() As Variant
    Dim SortIndex As Variant
    Values = RandomValues(20000, 3)
    Sort Values, SortIndex
    Assert.IsTrue Ascending(Values, SortIndex)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Sort")
Private Sub SortByIndex_LeavesTheArray()
    On Error GoTo TestFail

    Dim Values() As Variant
    Dim SortIndex As Variant
    Values = RandomValues(1000, 3)
    Dim Before As String: Before = Join(Values, ",")
    Sort Values, SortIndex
    Assert.AreEqual Before, Join(Values, ",")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Search")
Private Sub Search_FirstOfEqualValues()
    On Error GoTo TestFail

    Dim Values() As Variant
    Values = RandomValues(5000, 3)
    Sort Values
    Assert.AreEqual FirstPosition(Values, 1), CLng(Search(Values, 1))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Search")
Private Sub Search_ByIndexFirstOfEqualValues()
    On Error GoTo TestFail

    Dim Values() As Variant
    Dim SortIndex As Variant
    Values = RandomValues(5000, 3)
    Sort Values, SortIndex
    Assert.AreEqual FirstPosition(Values, 2, SortIndex), CLng(Search(Values, 2, SortIndex))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Search")
Private Sub Search_MissingValueIsNull()
    On Error GoTo TestFail

    Dim Values() As Variant
    Values = RandomValues(5000, 3)
    Sort Values
    Assert.IsTrue IsNull(Search(Values, 7))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Search")
Private Sub Search_NextEqualValue()
    On Error GoTo TestFail

    Dim Values() As Variant
    Values = RandomValues(5000, 3)
    Sort Values
    Dim First As Long: First = Search(Values, 1)
    Assert.AreEqual First + 1, CLng(Search(Values, 1, Start:=First))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Search")
Private Sub Search_Strings()
    On Error GoTo TestFail

    Dim Words() As Variant: Words = Array("b", "A", "c", "B", "a")
    Sort Words
    Assert.AreEqual 1&, CLng(Search(Words, "B"))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Search")
Private Sub Search_VariantDates()
    On Error GoTo TestFail

    Dim Days() As Variant: Days = Array(#3/1/2026#, #1/1/2026#, #2/1/2026#)
    Sort Days
    Assert.AreEqual 1&, CLng(Search(Days, #2/1/2026#))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Search")
Private Sub Search_ObjectsWithCustomCompare()
    On Error GoTo TestFail

    Dim Persons() As Variant: Persons = People()
    Set CompareCustom = New Person
    Sort Persons
    Assert.AreEqual 3&, CLng(Search(Persons, NewPerson("Anna", "Smit")))
    Set CompareCustom = Nothing

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Search")
Private Sub Search_IndexOfOtherLengthRaisesError()
    On Error GoTo TestFail

    Dim Values() As Variant
    Dim SortIndex As Variant
    Values = RandomValues(100, 3)
    Sort Values, SortIndex
    ReDim Preserve SortIndex(0 To 50)
    Dim ErrNo As Long
    On Error Resume Next
    Search Values, 1, SortIndex
    ErrNo = Err.Number: Err.Clear
    On Error GoTo TestFail
    Assert.AreEqual 5&, ErrNo

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsSorted")
Private Sub IsSorted_ThreeValues()
    On Error GoTo TestFail

    Dim Values() As Variant
    Values = RandomValues(5000, 3)
    Sort Values
    Assert.IsTrue IsSorted(Values)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsSorted")
Private Sub IsSorted_ByIndexThreeValues()
    On Error GoTo TestFail

    Dim Values() As Variant
    Dim SortIndex As Variant
    Values = RandomValues(5000, 3)
    Sort Values, SortIndex
    Assert.IsTrue IsSorted(Values, SortIndex)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsSorted")
Private Sub IsSorted_Unsorted()
    On Error GoTo TestFail

    Dim Values() As Variant
    Values = RandomValues(5000, 3)
    Assert.IsFalse IsSorted(Values)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
