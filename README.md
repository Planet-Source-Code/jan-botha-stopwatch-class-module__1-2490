<div align="center">

## StopWatch Class Module


</div>

### Description

I bet you're one of those programmers that want to time something in your code, aren't you? Well, with this code you can! This, is a Stopwatch! Not one of those that screw will up when midnight occurs while you're timing. This is... no, not Y2K compliant, infact, it's midnight-compliant! (Fully documented code!)
 
### More Info
 
Start a new project, add a new class module. Name the class module "CStopWatch"

Put 3 Command buttons on the form (Command1, Command2, Command3)

Number of seconds that have elapsed while you timed.

You'll have to change your system time if you want to see if this really is midnight-compliant!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jan Botha](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jan-botha.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jan-botha-stopwatch-class-module__1-2490/archive/master.zip)





### Source Code

```
'***************************************************************************
'*PUT THE FOLLOWING INTO A CLASS MODULE. NAME THE CLASS MODULE "CStopWatch"*
'***************************************************************************
Private m_StartTime As Single
Private m_StopTime As Single
Const cSecsInDay As Long = 86400
Public Enum cPauseConstants 'I'm not gonna explain this, consult VB Help if you want to know what it does
  cSeconds = 0
  cMinutes = 1
  cHours = 2
End Enum
Public Sub StartTiming()
  m_StartTime = Timer
End Sub
Public Sub StopTiming()
  m_StopTime = Timer
End Sub
Public Function TimeElapsed() As Single
  Dim tempTimeElapsed
  tempTimeElapsed = m_StopTime - m_StartTime 'see how many seconds passed since stopwatch has started
  If tempTimeElapsed < 0 Then 'if value of above is less than 0, assume that timing started before midnight and ended after midnight
    TimeElapsed = tempTimeElapsed + cSecsInDay 'add number of seconds in a day to the negative number and you have the time that has elapsed
   Else 'if it's a positive number...
    TimeElapsed = tempTimeElapsed
  End If
End Function
'****************************************************************************
'*To use the functions in your program, paste the following code into a form*
'****************************************************************************
'This goes in the Declaration Section
Dim TimeKeeper as CStopWatch
'Press command1 to start timing
Private Sub Command1_Click()
  Set TimeKeeper = New CStopWatch
  TimeKeeper.StartTiming
End Sub
'Press command2 to stop timing
Private Sub Command2_Click()
  TimeKeeper.StopTiming
End Sub
'Press command3 to display the number of seconds that have elapsed, in a MsgBox
Private Sub Command3_Click()
  Dim Elapsed as Single
  Elapsed = TimeKeeper.TimeElapsed
  MsgBox Elapsed
End Sub
'Please give comments and suggestions on this code. It's basically my first
'class module. Email me at: <c03jabot@prg.wcape.school.za>
```

