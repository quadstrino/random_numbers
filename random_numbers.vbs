' ++++++++++++++++++++++++++++++++++++++++++++++++
' This script was created to automate the process
' of inputting random numbers into a column of a
' spreadsheet.
'
' It will prompt the user to input a number, then 
' to select the cell in the spreadsheet. After that
' it will simulate the user typing the number and
' pressing enter, until all the numbers have been 
' entered into the spreadsheet.
'  
' Created by Levi McClaferty | 2025 
' ++++++++++++++++++++++++++++++++++++++++++++++++

Set WshShell = CreateObject("WScript.Shell")

' === Prompt user for how many numbers to generate ===

Do
    user_input = InputBox( _
        "Enter the max value of the random numbers you want to generate." & vbCrLf & _
        "You have 3 seconds to click into your spreadsheet after you click OK.", _
        "Random Number Generator")

    If IsNumeric(user_input) Then
        num_of_entries = CInt(user_input)
    Else
        MsgBox "Please enter a valid number.", vbExclamation, "Invalid Input"
    End If
Loop Until IsNumeric(user_input) And num_of_entries >= 1

' === Setup Time Intervals ===

delay_interval = 300 '  time between entries in milliseconds
pause_for_user = 3000 ' delay to give user time to select appropriate cell in spreadsheet

' === 3-second delay for user to click on cell in spreadsheet ===

WScript.Sleep pause_for_user

' === Generate and fill array ===

Dim numbers()
ReDim numbers(num_of_entries - 1)

For i = 0 To num_of_entries - 1
    numbers(i) = i + 1
Next

' === Randomize array ===  (using Fisher-Yates shuffle!)

Randomize
For i = num_of_entries - 1 To 1 Step -1
    j = Int(Rnd * (i + 1))
    _number = numbers(i)
    numbers(i) = numbers(j)
    numbers(j) = _number
Next

' === Type values into spreadsheet ===

For i = 0 To num_of_entries - 1
    WshShell.SendKeys numbers(i)
    WshShell.SendKeys "{ENTER}"
    WScript.Sleep delay_interval
Next