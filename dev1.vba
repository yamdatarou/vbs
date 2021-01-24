Sub dev1()
  schoolList = getSchoolList()
  schoolDeatailList = getSchoolDeatailList()
  sutudentList = getSutudentList()

  currentRow = 1

  For i = 1 To UBound(schoolList)
    fechRow = Sheet("schoolDeatailList").Range("A:A").Find(schoolList(i, 2)).Row
    s = fechRow - 1
    e = fechRow studentList(i, 3) - 1
    currentRow = currentRow + (e - s)
    Sheet("schoolDeatailList").Range("A" & s & ":" & "E" & e).Copy
    Sheet("new").Range("A" & currentRow).PasteSpecial
  Next
  
End SUb

Function getSchoolList()
  getSchoolList = Sheet("schoolList").Range("A2:C4")
End Function

Function getSchoolDeatailList()
  getSchoolDeatailList = Sheet("schoolDeatailList").Range("A2:B4")
End Function

Function getSutudentList()
  getSutudentList = Sheet("sutudentList").Range("A2:B4")
End Function
