
Sub inputbox_msgbox()
'write prg to get marks and show percentage
'Total marks is 150

'Inputbox - it is used to get input from the user
'It can take only one input at a time

Dim marks As Byte

marks = InputBox("Enter the marks", "MARKS", 78)
'Run prg by pressing f5

'Msgbox - it is used to show the output on screen
MsgBox "Percentage = " & marks * 100 / 150 & "%"
End Sub

'Enter two numbers and show the sum of two numbers.
Sub add_num()
Dim num1, num2 As Byte 'num2 is byte but num1 as variant
Dim n1 As Byte, n2 As Byte

n1 = InputBox("Enter the first number")
n2 = InputBox("Enter the second number")

MsgBox "Total is " & n1 + n2
End Sub


'Write a program to Evaluate and display if the number is Even or Odd.
'Remainder -- MOD function
'x=10 MOD 2

Sub even_odd()
Dim num As Integer

num = InputBox("Enter the number")

If num Mod 2 = 0 Then
    MsgBox num & " is a even number"
Else
    MsgBox num & " is a odd number"
End If

End Sub

'Nested IF - used when there are multiple conditions to be checked


Sub IF_ElseIF_example()
'Category of Certificate
Dim marks As Byte

marks = InputBox("Enter the marks", "MARKS", 78)

If marks > 85 Then
    MsgBox "Distinct"
ElseIf marks > 70 And marks <= 85 Then
    MsgBox "Completion"
Else
    MsgBox "No Certificate"
End If

End Sub

'Enter Mobile Num and give name of service provider based on first 3 digits

Sub mobile_provider()
Dim ph As String

ph = InputBox("Enter the phone number")

If Len(ph) <> 10 Then
    MsgBox "Invalid number"
ElseIf Left(ph, 3) = 990 Then
    MsgBox "Airtel"
ElseIf Left(ph, 3) = 901 Then
    MsgBox "BSNL"
ElseIf Left(ph, 3) = 902 Then
    MsgBox "MTS"
ElseIf Left(ph, 3) = 900 Then
    MsgBox "Voda"
Else
    MsgBox "Others"
End If
End Sub

Sub mobile_provider_Select()
Dim ph As String

ph = InputBox("Enter the phone number")

If Len(ph) <> 10 Then
    MsgBox "Invalid number"
Else
    Select Case Left(ph, 3)
        Case 990
            MsgBox "Airtel"
        Case 901
            MsgBox "BSNL"
        Case 902
            MsgBox "MTS"
        Case 900
            MsgBox "Voda"
        Case Else
            MsgBox "Others"
    End Select
End If
End Sub


Sub marks_grade()
Dim marks As Byte
marks = InputBox("enter the marks")

Select Case marks
    Case Is >= 90
        MsgBox "Grade: A"
    Case 80 To 89
        MsgBox "Grade: B"
    Case 70 To 79
        MsgBox "Grade: C"
    Case 60 To 69
        MsgBox "Grade: D"
    Case Is < 60
        MsgBox "Fail"
End Select
End Sub

'Loop - Definite (FOR Next); Conditional loop (DO WHILE, DO UNTIL)

Sub table_13()
 Dim i As Integer

For i = 1 To 10
    MsgBox 13 * i
Next i
End Sub



Sub add_FiveNumbers()

Dim No As String
Dim a As Byte
Dim b As Byte
Dim c As Byte
Dim d As Byte
Dim e As Byte


No = InputBox("Enter the Five Digit Number to Add", "No", 12345)

If Len(No) <> 5 Then
        MsgBox "Please Enter 5 digit numbers" & Len(No)
    Else
     MsgBox "Do the sum"
     a = Left(No, 1)
     b = Mid(No, 2, 1)
     c = Mid(No, 3, 1)
     d = Mid(No, 4, 1)
     e = Right(No, 1)
     f = a + b + c + d + e
     MsgBox "The total sum of " & a & " + " & b & " + " & c & " + " & d & " + " & e & " = " & f
          
 End If

End Sub


Sub check_3_GreatestNumbers()

Dim No As String
Dim a As Byte
Dim b As Byte
Dim c As Byte

No = InputBox("Please enter 3 digit number to check the greatest number", "No", 123)

If Len(No) = 3 Then
    MsgBox "Check the Greatest Number from " & No
    a = Left(No, 1)
    b = Mid(No, 2, 1)
    c = Right(No, 1)
        If a > b And a > c Then
            MsgBox " " & a & " is the Greatest number from " & b & " " & c
                ElseIf b > c And b > a Then
                    MsgBox " " & b & " is the Greatest number from " & a & " " & c
            Else
                    MsgBox " " & c & " is the Greatest number from " & a & " " & b
        End If
Else
    MsgBox "Please enter 3 digit number"
End If


End Sub


Sub check_windChillIndex()

Dim x As Byte
Dim y As Byte

x = InputBox("Please enter the Wind Speed in miles per hour")
y = InputBox("Please enter the temperature in Fahrenheit")

If x >= 0 And x <= 4 Then
    MsgBox "Wind Chill Index  If (0 <= v <= 4) Then WCI = t: " & y
    
        ElseIf x >= 45 Then
            MsgBox "Wind Chill Index  if (v >=45) then WCI = 1.6t - 55 : " & 1.6 * y - 55
        Else
            MsgBox "Wind Chill Index WCI = 91.4 + (91.4 - t)(0.0203v - 0.304(v)1/2 - 0.474):   " & 91.4 + (91.4 - y) * (0.0203 * x - 0.304 * (x) * 1 / 2 - 0.474)
End If
End Sub






'Start num and End num --> the value as well as interval matter

Sub odd_num()
Dim i As Integer

For i = 1 To 20 Step 2
    MsgBox i
Next i
End Sub

Sub countdown()
Dim i As Integer, final As String

For i = 10 To 1 Step -1
    final = final & i & vbCrLf
    'vbtab --> gives a tab space between elements
    'vbcrlf --> works like enter key
Next i

MsgBox final
End Sub

'Enter five digit number and calculate the sum of digits
'12310  --> 7

Sub sum_digits()
Dim num As Integer, final As Integer, i As Integer

num = InputBox("enter the  five digit number")

For i = 1 To Len(CStr(num))
    final = final + Mid(num, i, 1)
Next i
MsgBox "Total =" & final

End Sub

Sub prime_number()
Dim num As Integer, i As Integer

num = InputBox("Enter a number to check")

For i = 2 To num / 2
    If num Mod i = 0 Then
        counter = counter + 1
    End If
Next i

If counter = 0 Then
    MsgBox num & " is a prime number"
Else
    MsgBox num & " is not a prime number"
End If
End Sub


Sub avg_sub()
Dim marks As Integer, i As Integer, j As Integer, total As Integer

For j = 1 To 3
total = 0
For i = 1 To 10
    marks = InputBox("Enter the marks")
    total = total + marks
Next i

MsgBox "Average of Subject " & j & " : " & total / 10
Next j
End Sub

Sub marks_subjects()
Dim i As Integer, marks As Integer, j As Integer, total As Integer
Dim final As String

For j = 1 To 3
total = 0
For i = 1 To 10
marks = InputBox("Enter the marks")
total = total + marks
Next i

final = final & "average of subject" & j & ":" & total / 10 & vbCrLf
Next j
MsgBox final

End Sub



Sub do_while()
Dim i As Integer

Do While i <> 100
    i = i + 10
Loop

MsgBox i
End Sub


Sub mobile_provider_Dowhile()
Dim ph As String

ph = InputBox("Enter the phone number")

Do While Len(ph) <> 10
    MsgBox "Invalid number"
    ph = InputBox("Enter the phone number again")
Loop

If Left(ph, 3) = 990 Then
    MsgBox "Airtel"
ElseIf Left(ph, 3) = 901 Then
    MsgBox "BSNL"
ElseIf Left(ph, 3) = 902 Then
    MsgBox "MTS"
ElseIf Left(ph, 3) = 900 Then
    MsgBox "Voda"
Else
    MsgBox "Others"
End If
End Sub

    
Sub do_until()
Dim i As Integer

Do Until i = 100
    i = i + 10
Loop

MsgBox i
End Sub
    

Sub mobile_provider_goto()
Dim ph As String

line1:
ph = InputBox("Enter the phone number")

If Len(ph) <> 10 Then
    MsgBox "Invalid number"
    GoTo line1
ElseIf Left(ph, 3) = 990 Then
    MsgBox "Airtel"
ElseIf Left(ph, 3) = 901 Then
    MsgBox "BSNL"
ElseIf Left(ph, 3) = 902 Then
    MsgBox "MTS"
ElseIf Left(ph, 3) = 900 Then
    MsgBox "Voda"
Else
    MsgBox "Others"
End If
End Sub


Sub excel_fn()
Dim i As Integer

i = WorksheetFunction.RandBetween(1, 100)
MsgBox i

End Sub

Sub random()
Dim i As Integer, x As Integer, y As Byte, z As Byte
line1:
x = InputBox("Guess the number")
y = WorksheetFunction.RandBetween(1, 100)
z = 1
Do While x <> y
    If x - y > 10 Then
        MsgBox "Number is Too High"
        x = InputBox("Guess the number again")
    ElseIf x - y < -10 Then
        MsgBox "Number is Too Low"
        x = InputBox("Guess the number again")
    Else
        MsgBox "Number lies between +/- 10"
        x = InputBox("Guess the number again")
    End If
    z = z + 1
Loop
    MsgBox "U have guessed the correct number and u have guessd " & z & "times"
End Sub

Sub guessing_game()
Dim i As Integer, guess As Integer, counter As Integer
i = WorksheetFunction.RandBetween(1, 100)
MsgBox i

Do While i <> guess
    guess = InputBox("Enter the number")
    counter = counter + 1
    
    If guess > i + 10 Then
    MsgBox "guess too high"
    ElseIf guess < i - 10 Then
    MsgBox "guess too low "
    ElseIf guess = i Then
        Exit Do
    Else
    MsgBox "close guess"
    End If
Loop
MsgBox "Correct ANswer! You took a total of : " & counter & "chances"

End Sub

Sub bin_or_dec()
Dim i As Byte, x As Integer, y As Integer, flag As Byte
x = InputBox("enter the number")
For i = 1 To Len(Chr(x))
    y = Mid(Chr(x), i, 1)
    If y = 0 Or y = 1 Then
        flag = flag + 1
    End If
Next i
If flag = Len(Chr(x)) Then
    MsgBox "Number is binary and chaging to decimal will be= " & CDec(x)
Else
    MsgBox "Number is decimal and changing to binary will be= " & CByte(x)
End If
End Sub

Sub Triangle()
Dim s1 As Byte, s2 As Byte, s3 As Byte

s1 = InputBox("Enter the first side")
s2 = InputBox("Enter the second side")
s3 = InputBox("Enter the third side")

If s1 = s2 And s1 = s3 Then
    MsgBox "Equilateral Triangle"
ElseIf s1 <> s2 And s2 <> s3 And s3 <> s1 Then
    MsgBox "Scalene Triangle"
Else
    MsgBox "Isosceles Triangle"
End If
End Sub



Sub range_obj()
Range("A1") = 100

'Assign value of 500 to B1 of sheet1
Worksheets("Sheet1").Range("B1") = 500

'Assign a value of 1000 to A1 of another workbook
Workbooks("Abc.xlsx").Worksheets(1).Range("A1") = 1000

End Sub
Sub Worksheet_Deactivate()
MsgBox "You have left sheet2"
End Sub

'workbooks--> Worksheets--> Range
'Range in sheet --> assumption is that range belongs to
'the sheet in which you are writting the code

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
MsgBox "Changed"
End Sub



'When Range obj written in ThisWorkbook --> Assumption is
'Range belongs to the activesheet of the workbook
'in which code is being written

Sub range_obj()
Range("A1") = 100
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
Dim i As Integer
For i = 1 To 10
Range("A" & i) = i
Next i
End Sub




Sub method_1()
Worksheets.Add before:=Worksheets("Sheet1"), Count:=2
Range("B1:B10").Copy
Range("A1").PasteSpecial Paste:=xlPasteFormulas, Transpose:=True
End Sub


Sub workbooks_object()
'Add --> It will open a new workbook
Workbooks.Add

x = Workbooks.Count

Workbooks("abc.xls").Activate

Worksheets(2).Delete
Worksheets("Sheet3").Delete
Workbooks.Open Filename:="C:\Users\2016.xlsm"

Workbooks.Close
Workbooks("Abc.xls").Close
End Sub



Sub workbook_exercise()
Workbooks.Add
Workbooks.Add

MsgBox Workbooks.Count

Workbooks.Open Filename:="c:\use\aaa.xlsx"

Workbooks("Abc.xlsx").Activate
End Sub


Sub worksheets_example()
'Add a new sheet
Worksheets.Add after:=Worksheets("Sheet2"), Count:=3

'Delete a sheet
Worksheets(2).Delete
Worksheets("Sheet3").Delete

'Name the sheet
Worksheets(2).Name = "ABC"
Worksheets("abc").Name = "Sheet2"

'Count the sheets
x = Worksheets.Count

'Activate
Worksheets("Sheet2").Activate

'Hide Sheets - property - Visible
'Visible -True, False, xlVeryHidden
Worksheets("Sheet3").Visible = True 'Unhide the sheet
Worksheets("Sheet3").Visible = False 'Hides sheet but can unhide from excel
Worksheets("Sheet3").Visible = xlVeryHidden 'Hides sheet and only unhide thru excel
End Sub

Sub msgbox_button()
Dim response As VbMsgBoxResult

response = MsgBox("Do you want to save?", vbYesNo + vbQuestion)

If response = vbYes Then
    MsgBox "Save the file"
Else
    MsgBox "Don't save"
End If
End Sub


Sub active_obj()
MsgBox ActiveSheet.Name
MsgBox ActiveCell.Address
MsgBox ActiveCell.Value
End Sub


Sub msgbox_Addsheet()
Dim response As VbMsgBoxResult

response = MsgBox("Do you want to add a new sheet?", vbYesNo + vbQuestion)

If response = vbYes Then
Worksheets.Add
ActiveSheet.Name = InputBox("Enter the name")
End If
End Sub

Sub activate_sheetbyUser()
Dim final As String
Dim user As String
Dim J As Integer

    For i = 1 To Sheets.Count
        'Cells(i, 1) = Sheets(i).Name
        final = final & Sheets(i).Name & ":" & vbCrLf
    Next i
        'MsgBox "Below are the list of Sheet name from the workbook" & final

    user = InputBox("Please Enter the Sheet Name which you would like to Activate :" & final)
    'MsgBox user
    Worksheets(user).Activate

End Sub



Sub add_worksheet_afterSheet2()
Dim i As Integer

    For i = 1 To 2
    Worksheets.Add after:=Worksheets("Sheet2") ', Count:=1
    'MsgBox ActiveSheet.Name
    ActiveSheet.Name = ActiveSheet.Name & "_GD"
    
        'Worksheets.Add after:=Worksheets("Sheet2"), Count:=1
        'MsgBox ActiveSheet.Name
        'ActiveSheet.Name = ActiveSheet.Name & "_GD"
    Next i
End Sub

Sub range_obj()
Range("A1") = 100  'puts 100 in activeworkbook's activesheet
End Sub

Sub range_obj_diffplaces()
Range("A1") = 100

Worksheets("Sheet1").Range("B1") = 500

'Assign a value of 1000 to A1 of another workbook
Workbooks("Abc.xlsx").Worksheets(1).Range("A1") = 1000

End Sub

Sub range_obj_diffplaces_module()
Range("A1") = 100

'Assign value of 500 to B1 of sheet1
Worksheets("Sheet1").Activate
Range("B1") = 500

'Assign a value of 1000 to A1 of another workbook
Workbooks("Abc.xlsx").Worksheets(1).Activate
Range("A1") = 1000

End Sub

Sub range_examples()
'Cont cells
Range("B1:B10") = 1000

'Disont cells
Range("c1,d2,e3,f1:f4") = 5000
End Sub


Sub range_diffvalues()
'Range("A1:A10") = "1:10"
Dim i As Integer

For i = 1 To 10
    Range("A" & i) = i
Next i

End Sub

Sub diagonal()
Dim i As Integer, j As Integer
j = 1
For i = 1 To 5
'A - ascii is 65
    Range(Chr(64 + i) & j) = i
    j = j + 1
Next i
End Sub

'Write prg to get marks of ten students
'Show max, Min and Avg marks

Sub range_marks()
Dim i As Integer

'Range("A1") = "Excel Test Scores"
'
For i = 1 To 10
    Range("A" & i + 1) = InputBox("Enter the marks")
Next i

Range("C1") = "Max Marks"
Range("C2") = "Min Marks"
Range("C3") = "Avg Marks"

'Only values in excel
Range("D1") = WorksheetFunction.Max(Range("A2:A11"))
Range("D2") = WorksheetFunction.Min(Range("A2:A11"))
Range("D3") = WorksheetFunction.Average(Range("A2:A11"))

'To put formula in excel cells
Range("E1").Formula = "=max(A2:A11)"
Range("E2").Formula = "=min(A2:A11)"
Range("E3").Formula = "=average(A2:A11)"
End Sub

'Write prg to show multiplication table of num entered by user
Sub multiplication()
Dim num As Integer, i As Integer

num = InputBox("enter the number")

Range("A1") = "Multiplication Table of :" & num

Range("A3:A12") = num

For i = 1 To 10
    Range("B" & i + 2) = i
Next i

Range("C3:C12").Formula = "=A3*B3"
End Sub


Sub copy_paste()
Range("B3:B12").Copy Range("E3")
Range("B3:B12").Cut Range("E3")
End Sub

Sub pastespl()
Range("C3:C12").Copy
'Destination.pastespecial
Range("E1").PasteSpecial Paste:=xlPasteValues, Transpose:=True
End Sub

'Dynamic Range
Sub dynamic_range()
'range(start cell, end cell)
Range("B3:b12").Select
Range("b3", "b12").Select

'end cell
Range("B3").End(xlDown).Select

'Dynamic Range
Range("b3", Range("B3").End(xlDown)).Select
End Sub

Sub dynamic_Table()
'Ctrl+A
'Always use your starting cell of table to do currentregion
Range("A3").CurrentRegion.Select

'write a program to fill all the empty cells in the table with hyphen
Range("A1").CurrentRegion.Replace "", "-"
End Sub

Sub obj_Variable()
Dim car As String
car = InputBox("Enter the name of the car")

Range("G2") = "Total Sales"
Range("G3") = "Avg Sales"
Range("G4") = "Total Sales for " & car & " car"

Dim rng_sales As Range, rng_car As Range

Set rng_sales = Range("D2", Range("D2").End(xlDown))
Set rng_car = Range("B2", Range("B2").End(xlDown))

Range("H2") = WorksheetFunction.sum(rng_sales)
Range("H3") = WorksheetFunction.Average(rng_sales)
Range("H4") = WorksheetFunction.SumIf(rng_car, car, rng_sales)

Range("E2:E10").Name = "COGS"
MsgBox "count of COGS = " & Range("cogs").Count

End Sub


'Activate worksheet Report1
'activate cell c5 and calculate formula from Raw Data Sheet i4:i212

Sub report1()
Workbooks("VBA Class Exercise 1_v2_April2014.xlsm").Activate
Worksheets("Report 1").Activate

    Range("C5").Formula = WorksheetFunction.Sum(Worksheets("Raw Data").Range("i4:i212"))
    Range("C6").Formula = WorksheetFunction.Average(Worksheets("Raw Data").Range("i4:i212"))
    Range("C7").Formula = WorksheetFunction.Count(Worksheets("Raw Data").Range("i4:i212"))
    Range("C8").Formula = WorksheetFunction.AverageIfs(Worksheets("Raw Data").Range("i4:i212"), Worksheets("Raw Data").Range("e4:e212"), "Sunset")
    '=AVERAGEIFS(Sales,Product,"Sunset")
    Range("C9").Formula = WorksheetFunction.Min(Worksheets("Raw Data").Range("i4:i212"))
    Range("C10").Formula = WorksheetFunction.Max(Worksheets("Raw Data").Range("i4:i212"))

End Sub

' Report3

Sub report3()
Dim i As Integer
    i = 4
 Workbooks("VBA Class Exercise 1_v2_April2014.xlsm").Activate
 Worksheets("Report 3").Activate
    'Range("B4").Formula = WorksheetFunction.SumIfs(Worksheets("Raw Data").Range("i4:i212"), Worksheets("Raw Data").Range("D4:D212"), Range("A4"))
    'Range("C4").Formula = WorksheetFunction.CountIfs(Worksheets("Raw Data").Range("D4:D212"), Range("A4"))
    
    For i = 4 To 17
        Range("B" & i).Formula = WorksheetFunction.SumIfs(Worksheets("Raw Data").Range("i4:i212"), Worksheets("Raw Data").Range("D4:D212"), Range("A" & i))
        Range("C" & i).Formula = WorksheetFunction.CountIfs(Worksheets("Raw Data").Range("D4:D212"), Range("A" & i))
    Next i
  
End Sub


Sub report3a()
Worksheets("report 3").Range("A3").CurrentRegion.Clear
 Worksheets("Raw Data").Activate
Dim rng_customer As Range, rng_sales As Range
 
Set rng_customer = Range("D3", Range("D3").End(xlDown))
Set rng_sales = Range("I3", Range("I3").End(xlDown))
 
rng_customer.Copy Worksheets("Report 3").Range("A3")
 
Worksheets("report 3").Activate
Range("A3").CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlYes
 
Range("B3") = "Total Sales"
Range("C3") = "No. of Sales"
 
'How many unique customers
Dim x As Integer, i As Integer
x = Range("A4", Range("A4").End(xlDown)).Count
 
'Populate Sales and count
For i = 1 To x
Range("B" & i + 3) = WorksheetFunction.SumIf(rng_customer, Range("A" & i + 3), rng_sales)
Range("C" & i + 3) = WorksheetFunction.CountIf(rng_customer, Range("A" & i + 3))
Next i
End Sub



Sub select_unique_data()

Dim i As Integer
Dim x As Integer

Workbooks("VBA Class Exercise 1_v2_April2014.xlsm").Activate
Worksheets("Raw Data").Activate

'Worksheets("Report 3").Activate
'Worksheets("Raw Data").Range("D4:D212").Copy Range("E4")
'Worksheets("Raw Data").Range("D4", Range("D4").End(xlDown)).Copy Range("E4")

Range("D4", Range("D4").End(xlDown)).Copy Worksheets("report 3").Range("E4")


Worksheets("Report 3").Activate

'Range("E4").CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlYes
ActiveSheet.Range("E4", Range("E4").End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlNo

x = Range("E4", Range("E4").End(xlDown)).Count

MsgBox x
    
    For i = 1 To x
       Range("F" & i + 3).Formula = WorksheetFunction.SumIfs(Worksheets("Raw Data").Range("i4", Range("i4").End(xlDown)), Worksheets("Raw Data").Range("D4", Range(d4).End(xlDown)), Range("D" & i + 1))
       Range("G" & i + 3).Formula = WorksheetFunction.CountIfs(Worksheets("Raw Data").Range("D4", Range("D4").End(xlDown)), Range("D" & i + 1))
   
        Range("F" & i).Formula = WorksheetFunction.SumIfs(Worksheets("Raw Data").Range("i4:i212"), Worksheets("Raw Data").Range("D4:D212"), Range("A" & i + 3))
        Range("G" & i + 3).Formula = WorksheetFunction.CountIfs(Worksheets("Raw Data").Range("D4:D212"), Range("A" & i + 3))
    
   Next i

End Sub




Sub report3_auto()
'Worksheets("report 3").Range("A3").CurrentRegion.Clear
 Worksheets("Raw Data").Activate
Dim rng_customer As Range, rng_sales As Range
 
Set rng_customer = Range("D3", Range("D3").End(xlDown))
Set rng_sales = Range("I3", Range("I3").End(xlDown))
 
rng_customer.Copy Worksheets("Report 3").Range("A3")
 
Worksheets("report 3").Activate
Range("A3").CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlYes
 
Range("B3") = "Total Sales"
Range("C3") = "No. of Sales"
 
'How many unique customers
Dim x As Integer, i As Integer
x = Range("A4", Range("A4").End(xlDown)).Count
 
'Populate Sales and count
For i = 1 To x
Range("B" & i + 3) = WorksheetFunction.SumIf(rng_customer, Range("A" & i + 3), rng_sales)
Range("C" & i + 3) = WorksheetFunction.CountIf(rng_customer, Range("A" & i + 3))
Next i
End Sub


Sub report_6()
Worksheets("Sheet1").Range("B2").CurrentRegion.Clear

Worksheets("Raw Data").Activate

Dim rng As Range, rng1 As Range

Dim rng_customer As Range, rng_sales As Range, rng_region As Range
 
Set rng_customer = Range("D4", Range("D4").End(xlDown))
Set rng_sales = Range("I4", Range("I4").End(xlDown))
Set rng_region = Range("B4", Range("B3").End(xlDown))


Set rng = Range("D4", Range("D4").End(xlDown))
Set rng1 = Range("B4", Range("B4").End(xlDown))


rng.Copy Worksheets("Sheet1").Range("A3")
rng1.Copy Worksheets("Sheet1").Range("C3")

Worksheets("Sheet1").Activate
Range("A3").CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlNo
Range("C3").CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlNo
Range("C3", Range("C3").End(xlDown)).Copy
Range("B2").PasteSpecial Transpose:=True
Range("C3", Range("C3").End(xlDown)).Delete

Dim x As Integer, i As Integer
x = Range("A3", Range("A3").End(xlDown)).Count

Dim y As Integer, J As Integer, col_start As Integer
y = Range("B2", Range("B2").End(xlToRight)).Count
MsgBox y



For i = 1 To x
Range("B" & i + 2) = WorksheetFunction.SumIfs(rng_sales, rng_customer, Range("A" & i + 2), rng_region, Range("B2"))
Range("C" & i + 2) = WorksheetFunction.SumIfs(rng_sales, rng_customer, Range("A" & i + 2), rng_region, Range("C2"))
Range("D" & i + 2) = WorksheetFunction.SumIfs(rng_sales, rng_customer, Range("A" & i + 2), rng_region, Range("D2"))
Range("E" & i + 2) = WorksheetFunction.SumIfs(rng_sales, rng_customer, Range("A" & i + 2), rng_region, Range("E2"))
Range("F" & i + 2) = WorksheetFunction.SumIfs(rng_sales, rng_customer, Range("A" & i + 2), rng_region, Range("F2"))
Next i

'For j = 1 To y
'MsgBox Range("B" & j + 2)
'Range("B" & j + 2) = WorksheetFunction.SumIfs(rng_sales, rng_customer, Range("A" & j + 2), rng_region, Range("B2"))
'Next j


End Sub





'Pivot Tables - codes are to be written in module

Sub pivot_table_example()
Dim ptc As PivotCache, pt As PivotTable

'Storing the data for Pivot Table
Worksheets("filename").Activate

'Range("H6").CurrentRegion.Delete --> This command is used to delete previous pivot if we create pivot in same area thru prog

Set ptc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("A1").CurrentRegion)

'Add new sheet
Worksheets.Add

'pivot table--> pivot fields  --> pivot items
'set pt=activesheet.pivottables.add(pivot cache, destination to start pivot table
Set pt = ActiveSheet.PivotTables.Add(ptc, Range("A4"))

'Drag and drop field
'Create a summary for total sales of each car type
pt.PivotFields("CarType").Orientation = xlRowField
pt.PivotFields("Sales").Orientation = xlDataField

'Averages sale and count of sales
pt.PivotFields("Sales").Orientation = xlDataField
pt.PivotFields("Sales").Orientation = xlDataField

'Change default function
pt.PivotFields("Sum of Sales2").Function = xlAverage
pt.PivotFields("Sum of Sales3").Function = xlCount

'Change the Caption
pt.PivotFields("Average of Sales2").Caption = "Avg. Sales"
pt.PivotFields("Count of Sales3").Caption = "No. of Sales"
pt.PivotFields("Sum of Sales").Caption = "Total Sales"
End Sub


Sub pivot_group_dates()
'yearwise and month wise sales summary for all cars

Dim ptc As PivotCache, pt As PivotTable
Worksheets("filename").Activate

Set ptc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("A1").CurrentRegion)

Worksheets.Add

Set pt = ActiveSheet.PivotTables.Add(ptc, Range("A4"))

pt.PivotFields("CarType").Orientation = xlColumnField
pt.PivotFields("Date").Orientation = xlRowField
pt.PivotFields("Sales").Orientation = xlDataField

Range("A6").Group Start:=True, End:=True, periods:=Array(False, False, False, False, True, False, True)

'Start:=True, End:=True --> Auto mode
'End:=#12/3/2016# --> Date and time given in #
End Sub
Sub pivot_group_Numbers()
'yearwise and month wise sales summary for all cars

Dim ptc As PivotCache, pt As PivotTable
Worksheets("filename").Activate

Set ptc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("A1").CurrentRegion)

Worksheets.Add

Set pt = ActiveSheet.PivotTables.Add(ptc, Range("A4"))

pt.PivotFields("Sales").Orientation = xlRowField
pt.PivotFields("Cartype").Orientation = xlDataField

Range("A5").Group Start:=25000, End:=40000, by:=5000
End Sub



Sub report5()

Dim ptc As PivotCache, pt As PivotTable


Worksheets("Sheet5").Activate
If Worksheets("Sheet5").Range("A4") <> Empty Then
Worksheets("Sheet5").Range("A4").CurrentRegion.Delete
End If


Worksheets("Raw Data").Activate

Set ptc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("A3").CurrentRegion)
Worksheets("Sheet5").Activate
Set pt = ActiveSheet.PivotTables.Add(ptc, Worksheets("Sheet5").Range("A3"))

'Drag and drop filed
'create a summary for total sales
pt.PivotFields("Product").Orientation = xlRowField
pt.PivotFields("Sales").Orientation = xlDataField

'average sale and count
pt.PivotFields("Sales").Orientation = xlDataField
pt.PivotFields("Sales").Orientation = xlDataField

'Change the default function
pt.PivotFields("Sum of Sales2").Function = xlCount
pt.PivotFields("Sum of Sales3").Function = xlMax


'Change the name of the caption
pt.PivotFields("Count of Sales2").Caption = "Count of Sale"
pt.PivotFields("Max of Sales3").Caption = "Max of Sale"


End Sub



Sub clear_sheet()

Worksheets("Sheet5").Activate
If Worksheets("Sheet5").Range("A4") <> Empty Then
Worksheets("Sheet5").Range("A4").CurrentRegion.Delete
End If

End Sub


'How to group a fileds
'Yearwise and month wise sale summary for all product

Sub yearwise_summ()
Dim ptc As PivotCache, pt As PivotTable


Worksheets("Sheet5").Activate
If Worksheets("Sheet5").Range("A4") <> Empty Then
Worksheets("Sheet5").Range("A4").CurrentRegion.Delete
End If


Worksheets("Raw Data").Activate

Set ptc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("A3").CurrentRegion)
Worksheets("Sheet5").Activate
Set pt = ActiveSheet.PivotTables.Add(ptc, Worksheets("Sheet5").Range("A3"))

'Drag and drop filed
'create a summary for total sales
pt.PivotFields("Date").Orientation = xlRowField
pt.PivotFields("Product").Orientation = xlColumnField
pt.PivotFields("Sales").Orientation = xlDataField


Range("A5").Group start:=True, End:=True, Periods:=Array(False, False, False, False, False, False, True)

'Start:=true, End:=True --> Auto mode


End Sub





Sub report5_yearly()

Dim ptc As PivotCache, pt As PivotTable


Worksheets("Sheet5").Activate
If Worksheets("Sheet5").Range("A4") <> Empty Then
Worksheets("Sheet5").Range("A4").CurrentRegion.Delete
End If


Worksheets("Raw Data").Activate

Set ptc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("A3").CurrentRegion)
Worksheets("Sheet5").Activate
Set pt = ActiveSheet.PivotTables.Add(ptc, Worksheets("Sheet5").Range("A3"))

'Drag and drop filed
'create a summary for total sales
pt.PivotFields("Product").Orientation = xlRowField
pt.PivotFields("Sales").Orientation = xlDataField
pt.PivotFields("Date").Orientation = xlColumnField

'average sale and count
pt.PivotFields("Sales").Orientation = xlDataField
pt.PivotFields("Sales").Orientation = xlDataField

'Change the default function
pt.PivotFields("Sum of Sales2").Function = xlCount
pt.PivotFields("Sum of Sales3").Function = xlMax


'Change the name of the caption
pt.PivotFields("Count of Sales2").Caption = "Count of Sale"
pt.PivotFields("Max of Sales3").Caption = "Max of Sale"


'Range("A6").Group Start:=True, End:=True, Periods:=Array(False, False, False, False, False, False, True)

End Sub




Sub Calculated_field()
'Show product wise profit

Dim ptc As PivotCache, pt As PivotTable


Worksheets("Sheet5").Activate
If Worksheets("Sheet5").Range("A4") <> Empty Then
Worksheets("Sheet5").Range("A4").CurrentRegion.Delete
End If

Worksheets("Raw Data").Activate
Set ptc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("A3").CurrentRegion)

Worksheets("Sheet5").Activate
Set pt = ActiveSheet.PivotTables.Add(ptc, Worksheets("Sheet5").Range("A3"))

pt.PivotFields("Product").Orientation = xlRowField
pt.PivotFields("Sales").Orientation = xlDataField
'pt.PivotFields("Sales").Orientation = xlRowField

pt.CalculatedFields.Add Name:="Profit", Formula:="Sales*2%"
pt.PivotFields("Profit").Orientation = xlDataField

End Sub



Sub pivot_show_value_as()
'compare Audi sales with other cars
Dim ptc As PivotCache, pt As PivotTable
Worksheets("filename").Activate

Set ptc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("A1").CurrentRegion)

Worksheets.Add

Set pt = ActiveSheet.PivotTables.Add(ptc, Range("A4"))

pt.PivotFields("Sales").Orientation = xlDataField
pt.PivotFields("Cartype").Orientation = xlRowField

Range("b5").PivotField.Calculation = xlPercentDifferenceFrom
Range("b5").PivotField.BaseField = "CarType"
Range("b5").PivotField.BaseItem = "Audi"
End Sub

Sub calculated_field()
'show cartype wise profit
'2% of sales

Dim ptc As PivotCache, pt As PivotTable
Worksheets("filename").Activate

Set ptc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("A1").CurrentRegion)

Worksheets.Add

Set pt = ActiveSheet.PivotTables.Add(ptc, Range("A4"))

pt.PivotFields("Sales").Orientation = xlDataField
pt.PivotFields("Cartype").Orientation = xlRowField

pt.CalculatedFields.Add Name:="Profit", Formula:="=Sales*2%"
pt.PivotFields("Profit").Orientation = xlDataField

End Sub



Sub compare_Calculated()
'Show product wise profit

Dim ptc As PivotCache, pt As PivotTable


Worksheets("Sheet5").Activate
If Worksheets("Sheet5").Range("A4") <> Empty Then
Worksheets("Sheet5").Range("A4").CurrentRegion.Delete
End If

Worksheets("Raw Data").Activate
Set ptc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("A3").CurrentRegion)

Worksheets("Sheet5").Activate
Set pt = ActiveSheet.PivotTables.Add(ptc, Worksheets("Sheet5").Range("A3"))

pt.PivotFields("Customer").Orientation = xlRowField
pt.PivotFields("Sales").Orientation = xlDataField
'pt.PivotFields("Sales").Orientation = xlRowField


Range("b4").PivotField.Calculation = xlPercentDifferenceFrom
Range("b4").PivotField.BaseField = "Customer"
Range("b4").PivotField.BaseItem = "Amazon.com"
End Sub




'Cells - not an object but a property of Range object
Sub cells_example()
'cells represent only one cell
'D5 --> Cells(row, col) --> Cells(5,"d") or cells (5,4)

'cells - when we are working in table
'cells(i,j)

Cells(4, 2) = 500  'cell -->B4
Range("B1:B10").Cells(5, 4) = 1000 'B1 start to behave like A1 of excel
End Sub

Sub cells_colour_alt_rows()
Dim i As Integer, j As Integer

For j = 1 To 5
    For i = 1 To 13 Step 2
        Range("E5").Cells(i, j).Interior.Color = 65535
    Next i
Next j
End Sub


Sub activecell_example()
'the members of range can be used with activecell
MsgBox ActiveSheet.Name
MsgBox ActiveSheet.Name & ActiveCell.Address
MsgBox ActiveCell.Row
MsgBox ActiveCell.Column
MsgBox ActiveCell.Value
End Sub

Sub activecell_exercise()
'Find range from activecell
Range(ActiveCell, ActiveCell.End(xlDown)).Select

Dim total As Integer
total = WorksheetFunction.sum(Range(ActiveCell, ActiveCell.End(xlDown)))

'Reach last cell
ActiveCell.End(xlDown).Select

'Offset is used to move about in excel based activecell
'Object.offset(row index, col index).select/copy/value
'row index, col index - can be positive or negative

ActiveCell.Offset(1, 0) = total
End Sub


'Write prg to give random numbers from cells D10 to G20 and bold all the cells
'which have even number in it


Sub random_number()
Dim i As Integer, J As Integer
Dim k As Integer, z As Integer

Range("D10").CurrentRegion.Delete
    
    For J = 1 To 4
        For i = 1 To 10 'Step 2
            'Range("D10").Cells(i, j).Interior.Color = 65535
            Range("D10").Cells(i, J) = WorksheetFunction.RandBetween(1, 1000)
            k = Range("D10").Cells(i, J)
                
                z = k Mod 2
                If z = 0 Then
                    'MsgBox k
                    Range("D10").Cells(i, J).Font.Bold = True
                    Range("D10").Cells(i, J).Font.Size = 12
                Else
                    Range("D10").Cells(i, J).Font.Bold = False
                    Range("D10").Cells(i, J).Font.Italic = True
                End If
        Next i
    Next J
End Sub


Sub transpose_example()

Do While ActiveCell.Offset(1, 0) <> Empty
    'Come one cell down from empty cell
    ActiveCell.Offset(1, 0).Select
    
    'Select and copy the record
    Range(ActiveCell, ActiveCell.End(xlDown)).Copy
    
    'Paste in cell above
    ActiveCell.Offset(-1, 0).PasteSpecial Transpose:=True
    
    'delete the transposed data
    ActiveCell.Offset(1, 0).Select
    Range(ActiveCell, ActiveCell.End(xlDown)).Delete
Loop

End Sub



Sub active_cell()
ActiveSheet.Name
MsgBox ActiveSheet.Name & ActiveCell.Address

ActiveCell.Address
ActiveCell.Row
ActiveCell.Column
ActiveCell.Value

End Sub

Sub activecell_sumdown()
'active cell
'last cell from active cell

Range(ActiveCell, ActiveCell.End(xlDown)).Select

Dim total As Integer
total = WorksheetFunction.Sum(Range(ActiveCell, ActiveCell.End(xlDown)))

MsgBox total
'Reach the last cell

ActiveCell.End(xlDown).Select

'offset is used to move about in excel based activecell
'object.offset(row index, col index) 'select/copy/value
'row index, col can be postive or negative


ActiveCell.Offset(1, 0) = total

End Sub


Sub create_firstmarco_record()


Do While ActiveCell.Offset(1, 0) <> Empty

ActiveCell.End(xlDown).Select
Range(ActiveCell, ActiveCell.End(xlDown)).Copy
ActiveCell.Offset(-1, 0).PasteSpecial Transpose:=True

ActiveCell.Offset(1, 0).Select
Range(ActiveCell, ActiveCell.End(xlDown)).Delete
Loop



End Sub




Sub Autofilter()
Worksheets("filename").Activate

'Filter in CarType
Range("A1").CurrentRegion.Autofilter Field:=2, Criteria1:=Worksheets("Summary").Range("C3")

'Filter on Customer City
Range("A1").CurrentRegion.Autofilter Field:=3, Criteria1:=Worksheets("Summary").Range("C4"), _
Operator:=xlOr, Criteria2:=Worksheets("Summary").Range("D4")

'Filter on Sales column
Range("A1").CurrentRegion.Autofilter Field:=4, Criteria1:=">" & Worksheets("Summary").Range("C5")
'Code to remove criteria from one col
'Range("A1").CurrentRegion.Autofilter Field:=3
    
'Removes filter from entire data
'Selection.Autofilter
End Sub

'Sub, Function, Event
'Function - user defined function (udf)
'Coding for function should only be done in Modules

Function VLookup(Lookup_Value As Variant, Table_Array As Range, Col_Index_Num As Byte, Range_Lookup As Boolean) As Variant
'The variables declared in the first line should be used in prg
'the final output has to be stored in the name of function
'No Inputbox and no MSGBOX
End Function

'Create a function which removes enter sign from the cells
Function Remove_Enter(text As String) As String
Dim i As Integer

For i = 1 To Len(text)
If Mid(text, i, 1) <> Chr(10) Then
    Remove_Enter = Remove_Enter & Mid(text, i, 1)
End If
Next i
End Function

'In order to use the function in all the workbook saveas the function as Addin and
'activate the addin

'Create a function which can count the occurence of a character in a text
'Eg: text --> filename, occurence of --> e, output --> 2

Function Countletter(text As String, let1 As String) As String
Dim i As Integer
Dim j As Integer
Dim len1 As Integer
Dim len2 As Integer
Dim Count As Integer

len1 = Len(let1)
len2 = Len(text) 'Total length of text

For i = 1 To len2
If Mid(text, i, len1) = let1 Then
 Count = Count + 1
 End If
 
Next i

Countletter = Count
 
End Function


'Calling the SUB
Private Sub test1()
MsgBox "Hello"
End Sub

Sub test2()
test1 'this will call the sub test1
MsgBox "Bye"
End Sub



'create a function which removes the enter line feed or enter sign from the cell
'char10 is to identify enter
'no input bix and no msgbox

Function remove_Enter(text As String) As String
Dim i As Integer
'How many data type
'scan each and every character and see enter and remove it
    For i = 1 To Len(text)
            If Mid(text, i, 1) <> Chr(10) Then
            remove_Enter = remove_Enter & Mid(text, i, 1)
            End If
    Next i
End Function

'In order to use the function in all the workbook save it as active adam .xlam

Function Count_letter(text As String, let1 As String) As String
Dim i As Integer
Dim J As Integer
Dim len1 As Integer
Dim len2 As Integer
Dim Count As Integer
len1 = Len(let1)
len2 = Len(text)
    For i = 1 To len2
        If Mid(text, i, len1) = let1 Then
            Count = Count + 1
        End If
    Next i
Countletter = Count
End Function





Sub report4()

Worksheets("Sheet1").Activate
    Range("A3").CurrentRegion.AutoFilter Field:=4, Criteria1:=Worksheets("Report 4").Range("B3"), Operator:=xlOr, Criteria2:=Worksheets("Report 4").Range("D3")
    Range("A3").CurrentRegion.AutoFilter Field:=9, Criteria1:=">" & Worksheets("Report 4").Range("B4"), Operator:=xlAnd, Criteria2:="<" & Worksheets("Report 4").Range("D4")
        Range("A3").CurrentRegion.Copy
Worksheets("Report 4").Activate
    Range("A6").PasteSpecial

End Sub


'Q27
'Write a procedure such that when you select a cell in a table, corresponding row and column should gets shaded.


Sub shade_rowcol_Q27()
ActiveCell.Activate
    Range(ActiveCell, ActiveCell.End(xlDown)).Select
    Range(ActiveCell, ActiveCell.End(xlDown)).Interior.Color = 65535
    Range(ActiveCell, ActiveCell.End(xlUp)).Interior.Color = 65535
    Range(ActiveCell, ActiveCell.End(xlToLeft)).Interior.Color = 65535
    Range(ActiveCell, ActiveCell.End(xlToRight)).Interior.Color = 65535
End Sub

'Q26
'Write the function which calculates the average of top n values in a Range.

Sub cal_topn()
Dim startp As String
Dim endp As String
Dim i As Integer
Dim cou As String
Dim final As String

    startp = Range("J2")
    endp = Range("L2")
        MsgBox startp
        MsgBox endp

    'Range("J4").Formula = WorksheetFunction.Average(Worksheets("Report 4").Range("i7:i19"))
    '=AVERAGE(LARGE(I8:I20,{1,2}))
    'Range("K4").Formula = "=AVERAGE(LARGE(I8:I20,{1,2,3,4}))"
    Range("J4").Formula = WorksheetFunction.Average(Worksheets("Report 4").Range(startp & ":" & endp))
    Range("J5").Formula = "=AVERAGE(LARGE(" & startp & ":" & endp & ",{1,2}))"

i = Range("J3")
        MsgBox i
         For i = 1 To i - 1
         cou = cou & i & ","
          'final = final & Sheets(i).Name & ":" & vbCrLF
         Next i
        final = cou & WorksheetFunction.Max(i)
            MsgBox final
            MsgBox startp & ":" & endp & ",{" & final & "})"
        'Range("L5").Formula = "=AVERAGE(LARGE(" & startp & ":" & endp & ",{1,2,3}))"
        Range("L3").Formula = "=AVERAGE(LARGE(" & startp & ":" & endp & ",{" & final & "}))"
End Sub


'Q24
'Write a program that picks the students’ names and their test scores from excel sheet and outputs the following information:
'The average score
'Names of all students whose test scores are below the average, with an  appropriate message
'Highest test score and the name of all students having the highest score

Sub question24()


    Worksheets("Sheet30GD").Activate
    
    Range("A1", Range("A1").End(xlDown)).Select
    Range("A1", Range("A1").End(xlDown)).Copy
    Worksheets("Sheet29").Activate
    Range("A1").PasteSpecial
    
    Worksheets("Sheet30GD").Activate
    Range("D1", Range("D1").End(xlDown)).Select
    Range("D1", Range("D1").End(xlDown)).Copy
    Worksheets("Sheet29").Activate
    Range("B1").PasteSpecial
    
        Range("E2").Formula = WorksheetFunction.Average(Worksheets("Sheet29").Range("B2:B20"))
        Range("A1").CurrentRegion.AutoFilter Field:=2, Criteria1:="<" & Worksheets("Sheet29").Range("E2")
        Range("A1").CurrentRegion.Copy
        Range("D5").PasteSpecial
    'Switching off the suto filter
    Range("A1").AutoFilter
    
        Range("A1").CurrentRegion.AutoFilter Field:=2, Criteria1:=">" & Worksheets("Sheet29").Range("E2")
        Range("A1").CurrentRegion.Copy
        Range("H5").PasteSpecial
        'Range("i6").CurrentRegion.Sort key1:=Range("i6"), Order1:=xlAscending
        'Range.Sort key1:=Range("A1"), Order1:=xlAscending
    Range("A1").AutoFilter
 
End Sub




Sub fill_BlanksCellwithValue_Dash()
    Worksheets("Sheet6").Activate
    Range("H6").Activate
    Range("H6").CurrentRegion.Select
    
    Dim cell As Range
    Dim InputValue As String
    On Error Resume Next
        InputValue = InputBox("Enter value that will fill empty cells in selection Fill Empty Cells")
    
        For Each cell In Selection
            If IsEmpty(cell) Then
                cell.Value = InputValue
            End If
        Next
    
End Sub




Sub DeleteRows()
    
    Worksheets("Sheet6").Activate
    Range("B19").Activate
    Range("B19").CurrentRegion.Select
    
    
    Dim iStart As Integer
    Dim iCount As Integer
    Dim iStep As Integer
    Dim k As Integer
    k = 0
    iStep = 4   'Delete every 4th row
    Application.ScreenUpdating = False
    iStart = 1
    iCount = Selection.Rows.Count
    
    'MsgBox "iCount" & iCount
    'Find ending row to start deleting
    
    Do While iStart <= iCount
        k = k + iStep
        iStart = k
        'MsgBox "istat k " & iStart
        Selection.Rows(k).Delete
        Loop
   Application.ScreenUpdating = True

End Sub






Sub check_emailAddress()
    Dim emailaddress As String, n As Long
    Dim sItem As Variant
    ReDim sArray(1 To 2)
    
line1:
        emailaddress = InputBox("Please Enter Email Address")
        
        n = Len(emailaddress) - Len(Application.Substitute(emailaddress, "@", ""))
            
            If n <> 1 Then
                MsgBox ("You have entered an incorrect Email address Please check @")
                GoTo line1
            End If
            
        sArray(1) = Left(emailaddress, InStr(1, emailaddress, "@", 1) - 1)
        sArray(2) = Application.Substitute(Right(emailaddress, Len(emailaddress) - Len(sArray(1))), "@", "")
    
    
    For Each sItem In sArray
            'There should be atleast one character before @.
                If Len(sItem) <= 0 Then
                    MsgBox ("You have entered an incorrect Email address")
                    GoTo line1
                End If
            
            For n = 1 To Len(sItem)
                c = LCase(Mid(sItem, n, 1))
    'It must not contain any special character but only alphanumeric, underscore, period and dash or hyphen.
                If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
                MsgBox ("You have entered an incorrect Email address No Special Characters Allowed")
                    GoTo line1
                End If
            Next
            'Extreme characters must not be period or dot (.)
            If Left(sItem, 1) = "." Or Right(sItem, 1) = "." Then
                MsgBox ("You have entered an incorrect Email address Check .")
                GoTo line1
            End If
         Next
        
        'There must be atleast one period or dot after @
                If InStr(sArray(2), ".") <= 0 Then
                    MsgBox ("There must be atleast one period or dot after @")
                    GoTo line1
                End If
        'After the last dot or period, there must be either exactly 2 or 3 characters.
        n = Len(sArray(2)) - InStrRev(sArray(2), ".")
                If n <> 2 And n <> 3 Then
                    MsgBox ("After the last dot or period, there must be either exactly 2 or 3 characters")
                    GoTo line1
                End If
        'It must not contain 2 or more consecutive periods or dots.
                If InStr(emailaddress, "..") > 0 Then
                    MsgBox ("It must not contain 2 or more consecutive periods or dots.")
                    GoTo line1
                End If
    
    MsgBox ("You have entered a Valid Email Address")
    
End Sub


Function CheckEmail(ByVal emailaddress As String)
    Dim sArray As Variant, sItem As Variant
    Dim n As Long, c As String
    'Find the number of @, it should be exactly one.
    n = Len(emailaddress) - Len(Application.Substitute(emailaddress, "@", ""))
    If n <> 1 Then CheckEmail = MsgBox("Incorrect Email Address"): CheckEmail = False: Exit Function
    
    ReDim sArray(1 To 2)
    sArray(1) = Left(emailaddress, InStr(1, emailaddress, "@", 1) - 1)
    sArray(2) = Application.Substitute(Right(emailaddress, Len(emailaddress) - Len(sArray(1))), "@", "")
    
    For Each sItem In sArray
        'There should be atleast one character before @.
        If Len(sItem) <= 0 Then CheckEmail = MsgBox("There should be atleast one character before @."): CheckEmail = False: Exit Function
        For n = 1 To Len(sItem)
            c = LCase(Mid(sItem, n, 1))
                           'It must not contain any special character but only alphanumeric, underscore, period and dash or hyphen.
            If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then CheckEmail = MsgBox("Incorrect Email Address No Special Characters"): CheckEmail = False: Exit Function
        Next
        'Extreme characters must not be period or dot (.)
        If Left(sItem, 1) = "." Or Right(sItem, 1) = "." Then CheckEmail = MsgBox("Extreme characters must not be period or dot (.)"): CheckEmail = False: Exit Function
     Next
    'There must be atleast one period or dot after @
    If InStr(sArray(2), ".") <= 0 Then CheckEmail = MsgBox("There must be atleast one period or dot after @"): CheckEmail = False: Exit Function
    'After the last dot or period, there must be either exactly 2 or 3 characters.
    n = Len(sArray(2)) - InStrRev(sArray(2), ".")
    If n <> 2 And n <> 3 Then CheckEmail = MsgBox("After the last dot or period, there must be either exactly 2 or 3 characters"): CheckEmail = False: Exit Function
    'It must not contain 2 or more consecutive periods or dots.
    If InStr(emailaddress, "..") > 0 Then CheckEmail = MsgBox("It must not contain 2 or more consecutive periods or dots."): CheckEmail = False: Exit Function
    CheckEmail = True
End Function


' *** Team Lunch Project



Sub Button1_Click()

Dim i As Integer
Dim name As String

Worksheets("Project Statement").Activate
Range("G8").CurrentRegion.Delete

name = InputBox("Please enter your name")

i = Range("A7") + 1
Range("A7") = i
    
    Range("B8") = i
    Range("B9") = name

Worksheets("DataBase").Activate
    Range("A" & i) = i
    Range("B" & i) = name
Worksheets("Menu").Activate

End Sub

Sub Button5_Click()
Dim z As Integer
Dim i As Integer


Worksheets("Menu").Activate
z = Range("H6")
MsgBox z

Worksheets("Project Statement").Activate
i = Range("A7")
'
'    If z = 1 Then
'        Range("B10") = "Veg Paneer"
'    ElseIf z = 2 Then
'        Range("B10") = "Non Veg Mutton"
'    Else
'        Range("B10") = "Non Veg Chicken"
'    End If

Worksheets("DataBase").Activate
    If z = 1 Then
        Range("C" & i) = "Veg Paneer"
    ElseIf z = 2 Then
        Range("C" & i) = "Non Veg Mutton"
    Else
        Range("C" & i) = "Non Veg Chicken"
    End If


Worksheets("Project Statement").Activate



End Sub
Sub Button2_Click()

Worksheets("Project Statement").Activate
Range("G8").CurrentRegion.Delete

    Worksheets("DataBase").Activate
    Range("A1").CurrentRegion.Copy

        Worksheets("Project Statement").Activate
        Range("G8").PasteSpecial

End Sub



' Sales Monthly report

Sub Button3_Click()

Worksheets("Project Statement").Activate
        Range("A19").CurrentRegion.Delete

Worksheets("Data").Activate
        Range("A1").CurrentRegion.AutoFilter Field:=3, Criteria1:="=" & Range("M3")
        Range("A1").CurrentRegion.AutoFilter Field:=1, Criteria1:=">=01/01/" & Range("O3"), Operator:=xlAnd, Criteria2:="<=12/31/" & Range("O3")
        Range("A1").CurrentRegion.Copy

Worksheets("Project Statement").Activate
        Range("A19").PasteSpecial

End Sub


' =INDEX(L2:L9,M2)


' *** Final Dashboard
Sub Button2_Click()


Dim abc As PivotCache, pt As PivotTable
Dim Mkt As String

Worksheets("Market Wise Sales Summary").Activate
If Worksheets("Market Wise Sales Summary").Range("B8") <> Empty Then
Worksheets("Market Wise Sales Summary").Range("B8").CurrentRegion.Delete
End If


Worksheets("Raw Data").Activate
Mkt = Range("V3")
Set abc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("A1").CurrentRegion)
'Set abc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("A1:Q65536")) '.CurrentRegion)


Worksheets("Market Wise Sales Summary").Activate
Set pt = ActiveSheet.PivotTables.Add(abc, Worksheets("Market Wise Sales Summary").Range("B8"))


pt.PivotFields("Business_Segment").Orientation = xlRowField
pt.PivotFields("Market").Orientation = xlRowField
pt.PivotFields("Sales_Amount").Orientation = xlDataField



pt.PivotFields("Business_Segment").ClearAllFilters

pt.PivotFields("Business_Segment").PivotFilters.Add Type:=xlCaptionEquals, Value1:=Mkt  '"Maintenance and Repair"

'
'pt.PivotFields("Business_Segment").PivotFilters. _
'    Add Type:=xlCaptionEquals, Value1:="Maintenance and Repair"



End Sub



Sub Button1_Click()

Worksheets("Raw Data").Activate


    Range("A1", Range("A1").End(xlDown)).Select
    Range("A1", Range("A1").End(xlDown)).Copy
    Worksheets("Quater Wise Summary 2").Activate
    Range("B7").PasteSpecial

Worksheets("Raw Data").Activate
    Range("B1", Range("B1").End(xlDown)).Select
    Range("B1", Range("B1").End(xlDown)).Copy
    Worksheets("Quater Wise Summary 2").Activate
    Range("C7").PasteSpecial
    
Worksheets("Raw Data").Activate
    Range("I1", Range("I1").End(xlDown)).Select
    Range("I1", Range("I1").End(xlDown)).Copy
    Worksheets("Quater Wise Summary 2").Activate
    Range("D7").PasteSpecial
    
    
Worksheets("Raw Data").Activate
    Range("J1", Range("J1").End(xlDown)).Select
    Range("J1", Range("J1").End(xlDown)).Copy
    Worksheets("Quater Wise Summary 2").Activate
    Range("E7").PasteSpecial
    
Worksheets("Raw Data").Activate
    Range("M1", Range("M1").End(xlDown)).Select
    Range("M1", Range("M1").End(xlDown)).Copy
    Worksheets("Quater Wise Summary 2").Activate
    Range("A7").PasteSpecial


End Sub




Sub Sheet5_Button1_Click()

Sheets("Quarter Wise Summary").Select
' Pivot table created through excel just refreshing it
ActiveSheet.PivotTables("PivotTable12").RefreshTable

End Sub



