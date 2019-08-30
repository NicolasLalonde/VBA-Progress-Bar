VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadingBar 
   Caption         =   "Processing..."
   ClientHeight    =   1140
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "LoadingBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoadingBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BASIC USAGE:

' The Loading bar defines many methods for changing small parameters that you need not concern yourself with in regular use cases
' for regular use cases, the following suffice
'
' To launch/show the progress bar:                      initialize(total number of tasks)
' To update the progress by one task:                   incrementProgress
' To set the number of tasks completed manually:        setProgress(number of tasks completed)
' To change the description:                            setDescription(new description); or you can use the optional description parameter of initialize, incrementProgress, newSubTask, incrementSubProgress and endSubTask
' To close the progress bar:                            closes automatically when progress hits 100%; or you can use the terminate method (do not use the unload function)


' Subtasks:
' The loading bar supports subtasks which add a second loading bar to the bottom of the loading bar, representing subtask progress
' The primary bar will automatically fill according to the subtask's worth as the subtask increases its progress
'
' To launch a subtask:                                  newSubTask(total number of subtasks for the subtask); note, the totalWorth optional parameter can be used to make the subtask worth any number of tasks, eg newSubTask(10, totalWorth:=2)
' To update the progress by one subtask:                incrementSubProgress
' To set the number of subtasks completed manually:     setSubProgress(number of subtasks completed)
' To end a subtask:                                     endSubTask; or closes automatically when it hits 100%; or closes automatically when a new subtask is created


'FULL INTERFACE:

' initialize(Double, [String]):
'                         sets the number of tasks to be completed
'                         optionally changes the description
'                         shows the form

' setTotal(Double):       sets the number of tasks to be completed

' incrementProgress([String], [Double]):
'                         adds to the number of tasks completed,
'                         optionally changes the description under the progress bar
'
' newSubTask(Double, [String], [Double]):
'                         displays the subProgress bar and creates a new subtask with a given total
'                         optionally changes the description under the progress bar
'                         optionally takes a number (default 1) which corresponds to the number of total tasks which the subtask is worth
'                         ends any previous subtask
'
' endSubTask([String]):   hides the subprogress bar and resets all data pertaining to the subtask
'                         ensures progress is incremented by exactly the number given when the sub task was created
'                         optionally changes the description under the progress bar

' incrementSubProgress([String], [Double]):
'                         equivalent of incrementProgress for the subprogress bar
'                         the progress bar will also grow as the subtask gets completed, proportional to its worth
'
' setProgress(Double):    sets the current number of tasks completed

' setSubProgress(Double): sets the current number of subtasks completed

' dispPctProgress:        display progress as a percentage of total tasks done

' dispFracProgress:       display progress as a fraction (completed tasks / total tasks)

' showSubProgress:        displays the sub-progress bar

' hideSubProgress:        hides the sub-progress bar

' setDescription(String): changes the description of the current task below the progress bar

' setRefreshRate(Single): sets the minimum time to elapse before refreshing the bar

' setTitle(String):       sets the title of the loading bar window

' setListen(Boolean):     sets whether a doEvent line is ran while updating the view
'                         this prevents the form from going blank occasionally with a message of "not responding", at a slight performance cost

' setAutoClose(Boolean):  sets whether the loading bar window will automatically close when progress hits 100%
'                         and whether subtasks will automatically complete if they hit 100%

' terminate:              closes the progress window


'DEFAULTS:
' total and subtotal get set to 100
' progress shown as percentage
' no sub progress
' subtotal worth is set to 0
' Description = ""
' refreshRate of 0 (all updates are drawn)
' doEvents/Listen/preventNotResponding = True
' close when finished/autoclose = True


' INFO:
' drawing the bar continuously for very small tasks can bottleneck the actual processing of data as well as make the progress bar flash
' the bar can be set to refresh only after a certain time (in seconds, eg 0.05) has elapsed for these situations
' the bar will not draw any updates given during this period (although the object's values still get updated)
' all functions which update the view take an optional boolean parameter forceRefresh, which forces
' the progress bar to refresh regardless of the refresh rate set, useful if the bar seems to get stuck at the wrong task
' by default the refreshRate is 0, so all updates are "forced"


Private total As Double
Private progress As Double
Private subtotal As Double
Private subtotalWorth As Double
Private subprogress As Double
Private showFraction As Boolean
Private lastUpdated As Single
Private refreshRate As Single
Private preventNotResponding As Boolean
Private autoclose As Boolean
Private closed As Boolean


Private Sub UserForm_Initialize()
  total = 100
  progress = 0
  subtotal = 100
  subtotalWorth = 0
  subprogress = 0
  showFraction = False
  lastUpdated = Timer
  refreshRate = 0
  preventNotResponding = True
  autoclose = True
  closed = False
  
  setDescription ""
  hideSubProgress
End Sub

Public Sub initialize(totalTasks As Double, Optional description As String = "&NOTHING&")
  setTotal totalTasks
  If description <> "&NOTHING&" Then setDescription description
  closed = False
  Me.Show vbModeless
End Sub

Public Sub setTotal(totalTasks As Double)
  total = totalTasks
End Sub

Public Sub incrementProgress(Optional description As String = "&NOTHING&", Optional inc As Double = 1, Optional forceRefresh As Boolean = False)
  progress = progress + inc
  If description <> "&NOTHING&" Then setDescription description
  updateView forceRefresh
End Sub

Public Sub newSubTask(subtasktotal As Double, Optional description As String = "&NOTHING&", Optional totalWorth As Double = 1, Optional forceRefresh As Boolean = False)
  If subtotalWorth <> 0 Then endSubTask
  showSubProgress
  setDescription description
  subtotal = subtasktotal
  subtotalWorth = totalWorth
  subprogress = 0
  updateView forceRefresh
End Sub

Public Sub endSubTask(Optional description As String = "&NOTHING&", Optional forceRefresh As Boolean = False)
  Dim tempSub As Double
  tempSub = subtotalWorth
  subtotalWorth = 0
  subprogress = 0
  hideSubProgress
  incrementProgress description, tempSub
  updateView forceRefresh
End Sub

Public Sub incrementSubProgress(Optional description As String = "&NOTHING&", Optional inc As Double = 1, Optional forceRefresh As Boolean = False)
  subprogress = subprogress + inc
  If description <> "&NOTHING&" Then setDescription description
  updateView forceRefresh
End Sub


Public Sub setProgress(newprogress As Double, Optional forceRefresh As Boolean = False)
  progress = newprogress
  updateView forceRefresh
End Sub

Public Sub setSubProgress(newprogress As Double, Optional forceRefresh As Boolean = False)
  subprogress = newprogress
  updateView forceRefresh
End Sub



Public Sub dispFracProgress(Optional forceRefresh As Boolean = False)
  showFraction = True
  updateView forceRefresh
End Sub

Public Sub dispPctProgress(Optional forceRefresh As Boolean = False)
  showFraction = False
  updateView forceRefresh
End Sub

Public Sub showSubProgress(Optional forceRefresh As Boolean = False)
  Me.subBar.Visible = True
  Me.subFrame.Visible = True
  updateView forceRefresh
End Sub

Public Sub hideSubProgress(Optional forceRefresh As Boolean = False)
  Me.subBar.Visible = False
  Me.subFrame.Visible = False
  updateView forceRefresh
End Sub

Public Sub setDescription(description As String, Optional forceRefresh As Boolean = False)
  Me.description.Caption = description
  updateView forceRefresh
End Sub

Public Sub setRefreshRate(seconds As Single)
  refreshRate = seconds
End Sub

Public Sub setTitle(title As String, Optional forceRefresh As Boolean = False)
  Me.Caption = title
  updateView forceRefresh
End Sub

Public Sub setListen(doevent As Boolean)
  preventNotResponding = doevent
End Sub

Public Sub setAutoClose(automaticallyCloseProgressBar As Boolean)
  autoclose = automaticallyCloseProgressBar
End Sub

Public Sub terminate()
  If Not closed Then Me.Hide
  closed = True
End Sub



Private Sub updateView(Optional forceRefresh As Boolean = False)
  If enoughTimeElapsed Or forceRefresh Then
    Dim subtotalContribution As Double, realProgress As Double
    subtotalContribution = subtotalWorth * min(subprogress / subtotal, 1)
    realProgress = progress + subtotalContribution
  
    Me.Label1.Caption = IIf(showFraction, _
        Format(IIf(realProgress < total, realProgress, total), "#") & "/" & Format(total, "#") & " Completed", _
        IIf(realProgress < total, Format((realProgress / total) * 100, "#"), 100) & "% Completed")
    Me.Bar.width = 1 + (realProgress / total) * 200
    Me.subBar.width = 1 + (subprogress / subtotal) * 200
    Me.Repaint
    lastUpdated = Timer
    If preventNotResponding Then DoEvents
  End If
  If subprogress >= subtotal And subtotal > 0 And autoclose Then endSubTask
  If progress >= total And autoclose Then terminate
End Sub




Private Function min(a As Double, b As Double) As Double
  min = IIf(a <= b, a, b)
End Function

Private Function enoughTimeElapsed() As Boolean
  enoughTimeElapsed = Abs(Timer - lastUpdated) >= refreshRate
End Function
