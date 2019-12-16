# VBA-Progress-Bar
Progress bar to indicate progress for VBA subroutines which take a long time to run

The idea for using text highlighting to create the visual bar comes from https://www.excel-easy.com/vba/examples/progress-indicator.html

## BASIC USAGE:

 The Loading bar defines many methods for changing small parameters that you need not concern yourself with in regular use cases.
 
 For regular use cases, the following suffice:

 To launch/show the progress bar:                      initialize(total number of tasks)
 
 To update the progress by one task:                   incrementProgress
 
 To set the number of tasks completed manually:        setProgress(number of tasks completed)
 
 To change the description:                            setDescription(new description); or you can use the optional description parameter of initialize, incrementProgress, newSubTask, incrementSubProgress and endSubTask
 
 To close the progress bar:                            closes automatically when progress hits 100%; or you can use the terminate method (do not use the unload function)


 
 #### Subtasks:
 
 The loading bar supports subtasks which add a second loading bar to the bottom of the loading bar, representing subtask progress. The primary bar will automatically fill according to the subtask's worth as the subtask increases its progress.

+ To launch a subtask:                                  

 newSubTask(total number of subtasks for the subtask); note, the totalWorth optional parameter can be used to make the subtask worth any number of tasks, eg newSubTask(10, totalWorth:=2)
 
+ To update the progress by one subtask:                

 incrementSubProgress
 
+ To set the number of subtasks completed manually:    

 setSubProgress(number of subtasks completed)
 
+ To end a subtask:                                    

 endSubTask; or closes automatically when it hits 100%; or closes automatically when a new subtask is created


## FULL INTERFACE:

*Note: Optional parameters are denoted with square brackets* 

 **initialize(Double, [String])**:
 
   sets the number of tasks to be completed
                         optionally changes the description
                         shows the form

 **setTotal(Double)**:
 
   sets the number of tasks to be completed

 **incrementProgress([String], [Double])**:
 
   adds to the number of tasks completed,
                         optionally changes the description under the progress bar

 **newSubTask(Double, [String], [Double])**:
 
   displays the subProgress bar and creates a new subtask with a given total
                         optionally changes the description under the progress bar
                         optionally takes a number (default 1) which corresponds to the number of total tasks which the subtask is worth
                         ends any previous subtask

 **endSubTask([String])**:   
 
   hides the subprogress bar and resets all data pertaining to the subtask
                         ensures progress is incremented by exactly the number given when the sub task was created
                         optionally changes the description under the progress bar

 **incrementSubProgress([String], [Double])**:
 
   equivalent of incrementProgress for the subprogress bar
                         the progress bar will also grow as the subtask gets completed, proportional to its worth

 **setProgress(Double)**:
 
 sets the current number of tasks completed

 **setSubProgress(Double)**:
  
  sets the current number of subtasks completed

 **dispPctProgress**:        
 
 display progress as a percentage of total tasks done

 **dispFracProgress**:
 
 display progress as a fraction (completed tasks / total tasks)

 **showSubProgress**:
 
 displays the sub-progress bar

 **hideSubProgress**:        
 
 hides the sub-progress bar

 **setDescription(String)**:
 
 changes the description of the current task below the progress bar

 **setRefreshRate(Single)**: 
 
 sets the minimum time (in seconds) to elapse before refreshing the bar

 **setTitle(String)**:
 
 sets the title of the loading bar window

 **setListen(Boolean)**:
 
 sets whether a doEvent line is ran while updating the view
                         this allows MS word/excel to perform other actions while the script is running. Prevents the form from going blank occasionally with a message of "not responding", at a slight performance cost

 **setAutoClose(Boolean)**:  
 
 sets whether the loading bar window will automatically close when progress hits 100%
                         and whether subtasks will automatically complete if they hit 100%

 **terminate**:
 
 closes the progress window


## DEFAULTS:

 total and subtotal get set to 100
 
 progress shown as percentage
 
 no sub progress
 
 subtotal worth is set to 0
 
 Description = ""
 
 refreshRate of 0 (all updates are drawn)
 
 doEvents/Listen/preventNotResponding = True
 
 close when finished/autoclose = True


 ### Refresh info:
 Drawing the bar continuously for very small tasks can bottleneck the actual processing of data as well as make the progress bar flash.
 
 The bar can be set to refresh only after a certain time (in seconds, eg 0.05) has elapsed for these situations.
 
 The bar will not draw any updates given during this period (although the object's values still get updated).
 All functions which update the view take an optional boolean parameter forceRefresh, which forces the progress bar to refresh regardless of the refresh rate set, useful if the bar seems to get stuck at the wrong task (because another task started but the bar hasn't refreshed yet so the task name was not changed). By default the refreshRate is 0, so all updates are "forced".
