<div align="center">

## 4 function Calculator with 8 level memory


</div>

### Description

This is a 4 function calculator with 8 - level memory, it performs calculations
 
### More Info
 
Nothing, it works nicely

Values you calculated

None, I hope :)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[James Bergeron](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/james-bergeron.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/james-bergeron-4-function-calculator-with-8-level-memory__1-1189/archive/master.zip)

### API Declarations

```

		NONE
```


### Source Code

```
Dim calcarray(0 To 3) As Double
Dim holder As Integer
Dim operation As Integer
Dim decicount As Integer
Dim newnum As Integer
Dim clearcount As Integer
Dim memstorebut(1 To 8) As Double
Dim location As Single
Private Sub clear_Click()
     If clearcount = 0 Then
       txtcal.Text = ""
       clearcount = 1
     Else
       calcarray(0) = 0
       clearcount = 0
     End If
     decicount = 0
End Sub
Private Sub cmdInfo_Click()
  Dim Sure As String
  Sure = "Created By James Bergeron, For more info e-mail berg0036@algonquinc.on.ca"
  Rem Get results from the button click (action)
  ButtonClicked = MsgBox(Sure, 0 + 256 + 32, "Info")
End Sub
Private Sub decimal_Click()
  clearcount = 0
  If decicount = 0 Then
    txtcal.Text = txtcal.Text + decimal.Caption
    decicount = 1
  Else
    txtcal.Text = txtcal.Text
  End If
End Sub
Private Sub digit_Click(Index As Integer)
  If newnum = 1 Then
   txtcal.Text = ""
   txtcal.Text = txtcal.Text + digit(Index).Caption
   calcarray(holder) = txtcal.Text
   newnum = 0
  Else
  txtcal.Text = txtcal.Text + digit(Index).Caption
  calcarray(holder) = txtcal.Text
  End If
  clearcount = 0
End Sub
Private Sub equal_Click()
  Select Case operation
    Case 1
       txtcal.Text = calcarray(holder - 1) + calcarray(holder)
       calcarray(0) = txtcal.Text
    Case 2
       txtcal.Text = calcarray(holder - 1) - calcarray(holder)
       calcarray(0) = txtcal.Text
    Case 3
       txtcal.Text = calcarray(holder - 1) * calcarray(holder)
       calcarray(0) = txtcal.Text
    Case 4
       If calcarray(holder) = 0 Then
         txtcal.Text = "Error, can't divide by 0"
       Else
         txtcal.Text = calcarray(holder - 1) / calcarray(holder)
         calcarray(0) = txtcal.Text
       End If
    Case Else
      txtcal.Text = txtcal.Text
  End Select
  operation = 5
  holder = 0
  decicount = 0
  newnum = 1
  clearcount = 0
End Sub
Private Sub Form_Load()
operation = 0
location = 0
decicount = 0
End Sub
Private Sub memclear_Click()
  clearcount = 0
  For i = 1 To 8
    memstorebut(i) = 0
  Next i
  location = 0
End Sub
Private Sub memrecall_Click()
  clearcount = 0
  newnum = 1
  If location >= 1 Then
    txtcal.Text = memstorebut(location)
    calcarray(holder) = memstorebut(location)
    memstorebut(location) = 0
    location = location - 1
  End If
End Sub
Private Sub memstore_Click()
  clearcount = 0
  If location <= 7 And txtcal.Text > "" Then
    location = location + 1
    memstorebut(location) = txtcal.Text
  End If
End Sub
Private Sub mult_Click()
  Call equal_Click
  holder = holder + 1
  operation = 3
End Sub
Private Sub plus_Click()
   Call equal_Click
   holder = holder + 1
   operation = 1
End Sub
Private Sub div_Click()
  Call equal_Click
  holder = holder + 1
  operation = 4
End Sub
Private Sub sub_Click()
  Call equal_Click
  holder = holder + 1
  operation = 2
End Sub
```

