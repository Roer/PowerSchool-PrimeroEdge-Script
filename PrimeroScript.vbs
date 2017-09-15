Const ForReading = 1
'Enter District AUN
strDistAUN = ""
'I did not use these fields, so I left them blank
strApartment = ""
strAddressLine2 = ""
strSpecialInstructions = ""
	
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("student.export.text", ForReading)

'Order fields and add missing fields

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    arrFields = Split(strLine, vbTab)
    strStateID = arrFields(0)
    strLocalID = arrFields(1)
	strSSN = arrFields(2)
	strGrade = arrFields(3)
	strGender = arrFields(4)
	strDOB = arrFields(5)
	strFirstName = arrFields(6)
	strMiddleName = arrFields(7)
	strLastName = arrFields(8)
	strSchoolID = arrFields(9)
	strHomeroom = arrFields(10)
	strStreet = arrFields(11)
	strCity = arrFields(12)
	strState = arrFields(13)
	strZip = arrFields(14)
	strEthnicity = arrFields(15)
	strRace = arrFields(16)
	strGuard_FName = arrFields(17)
	strGuard_LName = arrFields(18)
	strGuard_Email = arrFields(19)
	strGuard_WPhone = arrFields(20)
	strGuard_HPhone = arrFields(21)

    strNewContent = strNewContent &  strStateID &  Chr(9) & strLocalID &  Chr(9) & strSSN &  Chr(9) & strGrade &  Chr(9) & strGender &  Chr(9) & strDOB &  Chr(9) & strDistAUN &  Chr(9) & strFirstName &  Chr(9) & strMiddleName &  Chr(9) & strLastName &  Chr(9) & strSchoolID &  Chr(9) & strHomeroom &  Chr(9) & strApartment &  Chr(9) & strStreet &  Chr(9) & strAddressLine2 &  Chr(9) & strCity &  Chr(9) & strState &  Chr(9) & strZip &  Chr(9) & strEthnicity &  Chr(9) & strRace &  Chr(9) & strGuard_FName & Chr(9) & strGuard_LName &  Chr(9) & strGuard_Email &  Chr(9) & strGuard_WPhone &  Chr(9) & strHPhone &  Chr(9) & strSpecialInstructions & vbCrLf
Loop

objFile.Close

Set objFile = objFSO.CreateTextFile("student.txt")

'Insert correct headers

objFile.WriteLine "StateStudentID" & Chr(9) & "LocalID" & Chr(9) & "SSN" & Chr(9) & "Grade" & Chr(9) & "Gender" & Chr(9) & "BirthDate" & Chr(9) & "DisctrictCode" & Chr(9) & "FirstName" & Chr(9) & "MiddleName" & Chr(9) & "LastName" & Chr(9) & "SchoolCode" & Chr(9) & "Homeroom" & Chr(9) & "Apartment" & Chr(9) & "AddressLine1" & Chr(9) & "AddressLine2" & Chr(9) & "City" & Chr(9) & "State" & Chr(9) & "Zip" & Chr(9) & "Ethnicity" & Chr(9) & "Race" & Chr(9) & "GuardianFirstName" & Chr(9) & "GuardianLastName" & Chr(9) & "GuardianEmail" & Chr(9) & "GuardianWorkPhone" & Chr(9) & "GuardianHomePhone" & Chr(9) & "SpecialInstructions"
objFile.Write strNewContent

objFile.Close


'Rewrite Building ID for Grade 7
'Our district needed to separate two grades out from an otherwise unified building

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("student.txt", ForReading)

Do Until objFile.AtEndOfStream
    strLineFour = objFile.ReadLine
    If InStr(strLineFour, Chr(9) & "7" & Chr(9)) > 0 Then
		strLineFour = Replace(strLineFour, Chr(9) & "####" & Chr(9), Chr(9) & "####" & Chr(9))
	End If
	strNewContentsFour = strNewContentsFour & strLineFour & vbCrLf
Loop

objFile.Close

Set objFile = objFSO.OpenTextFile("student.txt", ForWriting)
objFile.WriteLine strNewContentsFour
objFile.Close

'Rewrite Building ID for Grade 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("student.txt", ForReading)

Do Until objFile.AtEndOfStream
    strLineSeven = objFile.ReadLine
    If InStr(strLineSeven, Chr(9) & "8" & Chr(9)) > 0 Then
		strLineSeven = Replace(strLineSeven, Chr(9) & "####" & Chr(9), Chr(9) & "####" & Chr(9))
	End If
	strNewContentsSeven = strNewContentsSeven & strLineSeven & vbCrLf
Loop

objFile.Close

Set objFile = objFSO.OpenTextFile("student.txt", ForWriting)
objFile.WriteLine strNewContentsSeven
objFile.Close

'Rewrite 0 as KG

Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("student.txt", ForReading)

strText = objFile.ReadAll
objFile.Close
strNewText = Replace(strText, Chr(9) & "0" & Chr(9), Chr(9) & "KG" & Chr(9))

Set objFile = objFSO.OpenTextFile("student.txt", ForWriting)
objFile.WriteLine strNewText
objFile.Close


'Remove lines with SchoolID 700
'We have several buildings in our PowerSchool system that do not corrospond to actual buildings

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("student.txt", ForReading)

Do Until objFile.AtEndOfStream
    strLineOne = objFile.ReadLine
    If InStr(strLineOne, Chr(9) & "700" & Chr(9)) = 0 Then
        strNewContentsOne = strNewContentsOne & strLineOne & vbCrLf
    End If
Loop

objFile.Close

Set objFile = objFSO.OpenTextFile("student.txt", ForWriting)
objFile.Write strNewContentsOne

objFile.Close

'Remove lines with SchoolID 800

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("student.txt", ForReading)

Do Until objFile.AtEndOfStream
    strLineTwo = objFile.ReadLine
    If InStr(strLineTwo, Chr(9) & "800" & Chr(9)) = 0 Then
        strNewContentsTwo = strNewContentsTwo & strLineTwo & vbCrLf
    End If
Loop

objFile.Close

Set objFile = objFSO.OpenTextFile("student.txt", ForWriting)
objFile.Write strNewContentsTwo

objFile.Close

'Remove lines with SchoolID 900

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("student.txt", ForReading)

Do Until objFile.AtEndOfStream
    strLineThree = objFile.ReadLine
    If InStr(strLineThree, Chr(9) & "900" & Chr(9)) = 0 Then
        strNewContentsThree = strNewContentsThree & strLineThree & vbCrLf
    End If
Loop

objFile.Close

Set objFile = objFSO.OpenTextFile("student.txt", ForWriting)
objFile.Write strNewContentsThree

objFile.Close
