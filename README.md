# General-VBA-Error-Handler
A really simple error handler that gives a general message to the end user about how to fix an error.

## Purpose

In continuation of my dump of all the small project I've made onto GitHub to keep everything in the same place. Here is an Error Handler that I made for many Excel files that I've worked on.

The point of this is to have a few generalized Error Handler that can be used on pretty much ANY excel file. When an error happens then a very generalized message will appear to the end user sometimes with ways that MIGHT fix the error.

It is released into the Public Domain so do with it what you want.

I may or may not update this; because as I've said I'm trying to just put the random code bits I'ver made throughout the years all on Github as a central place to be able to reference them.

## How to Use

Download the ErrorHandling.bas and then in Excel under Visual Basic, Right click on Modules and select "Import File". Select ErrorHandling.bas from here you downloaded.

To have the error handler work you must first make sure that gEnableErrorHandling is True (this is useful to turn off when tweaking code because otherwise you won't get the option to "Debug" and error).

Now for each and every function / sub that you want to use this for follow the example sub below.

```visual-basic

Private Sub ThisIsMyExample()
If gEnableErrorHandling Then On Error GoTo errorHandler
 '**Main Code**
Exit Sub
errorHandler:
    Dim ErrorAction As Integer
    ErrorAction = errHandler(Err.Number)
    Select Case ErrorAction
        Case 0:
            'Ignore Error
            Resume
        Case 1:
            'End Macro
            Exit Sub
        Case 2:
            ErrorAction = MsgBox("Do you wish to close excel?", vbYesNo, "Close?")
            If ErrorAction = vbYes Then
                ThisWorkbook.Close
            Else
                Exit Sub
            End If
        Case 3:
            'Skip error causing line
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

```

It is VERY important that the line :

```visual-basic
If gEnableErrorHandling Then On Error GoTo errorHandler
```

Is right after you call the sub / function.

It is also important that you change the 

```visual-basic
Exit Sub
End Sub
```

To 

```visual-basic
Exit Function
End Function
```

If you have a function not a sub.


What will then happen is when an error happens the macro will try to do the following depending on what the error is.

* Do Nothing (just keep running)

* Exit Sub / function

* Prompt the user to close the workbook

* Skip the line causing the error.
