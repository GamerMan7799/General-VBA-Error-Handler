Attribute VB_Name = "ErrorHandling"
'---------------------------------------------------------------------------------------
' Module    : ErrorHandling
' Author    : Patrick Rye
' Date      : 7/15/2015
' Purpose   : A very generic way for the code to handle errors, modify this to suit your needs
'---------------------------------------------------------------------------------------

Option Explicit
Public Const gEnableErrorHandling As Boolean = True 'Global varaiable if you should or should not enable error handling
'Full list of errors can be found at: https://support.microsoft.com/en-us/kb/146864
Public Function errHandler(ByRef ErrorNumber As Integer) As Integer
Dim SomeNum As Integer
'On Error GoTo errHandler
'Error Code for the error code, it is not recommended to have this as if it is the same error every time it will get stuck.

'Action Return codes are:
'0 - Do Nothing (Just ignore the error) aka Resume
'1 - Exit sub / function
'2 - Close Workbook
'3 - Skip line, aka Resume Next

Select Case ErrorNumber
    Case 0:
        'No Error
        errHandler = 0 'Do nothing
    Case 3, 5, 10, 13, 16, 17, 20, 35, 59, 62, 63, 74, 91, 92, 93, 94, 323, 328, 361, 364, 365, 380, 381, 382, 383, 385, 387, 393, 394, 422, 423, 424, 425, 429, 430, 432, 438, 440, 442, 443, 445, 446, 447, 448, 451, 452, 454, 455, 457, 458, 459, 460, 461, 480, 1000, 1001, 1004, 1005, 1006, 31004, 31018, 31027, 31032:
        '3 = Return Without GoSub
        '5 = Invalid Procedure
        '10 = Duplicate Definition
        '13 = Type Mismatch
        '16 = String Formula too complex
        '17 = Can't perform requested operation
        '20 = Resume without Error
        '35 = Sub or function not defined.
        '59 = bad record length
        '62 = Input past end of line
        '63 = Bad Record number
        '74 = Can't rename with different drive (aka you tried to movea file to a diff
        '91 = Object variable not set
        '92 = For loop not initialized
        '93 = invalid pattern string
        '94 = invalid use of Null
        '323 = can't load module invalid format
        '361 = Can't Load or unload this object
        '364 = Object was unloaded
        '365 = Unable to laod within this context
        '328 = Illegal parameter; can't write arrays
        '380 = Invalid property value (version 97)
        '381 = Invalid property-array index (version 97)
        '382 = Property Set can't be executed at run time (version 97)
        '383 = Property Set can't be used with a read-only property (version 97)
        '385 = Need property-array index (version 97)
        '387 = Property Set not permitted (version 97)
        '393 = Property Get can't be executed at run time (version 97)
        '394 = Property Get can't be executed on write-only property (version 97)
        '422 = Property not found (version 97)
        '423 = Property or method not found
        '424 = Object Required
        '425 = Invalid object use (version 97)
        '429 = ActiveX component can't create object or return reference to this object (version 97)
        '430 = Class doesn 't support OLE Automation
        '432 = File name or class name not found during Automation operation (version 97)
        '438 = Object doesn 't support this property or method
        '440 = OLE Automation error
        '442 = Connection to type library or object library for remote process has been lost (version 97)
        '443 = Automation object doesn't have a default value (version 97)
        '445 = Object doesn 't support this action
        '446 = Object doesn 't support named arguments
        '447 = Object doesn 't support current locale settings
        '448 =  Named Not argument
        '451 = Object not a collection
        '452 = Invalid Ordinal
        '454 = Code Not resource
        '455 = Code resource lock error
        '457 = This key is already associated with an element of this collection (version 97)
        '458 = Variable uses a type not supported in Visual Basic (version 97)
        '459 = This component doesn't support events (version 97)
        '460 = Invalid clipboard format (version 97)
        '461 = Specified format doesn't match format of data (version 97)
        '480 = Can 't create AutoRedraw image (version 97)
        '1000 = Classname does not have propertyname property
        '1001 = Classname does not have methodname method
        '1004 = Methodname method of classname class failed
        '1005 = Unable to set the propertyname property of the classname
        '1006  = Unable to get the propertyname property of the classname
        '31004 = No Object
        '31018 = Class is not set
        '31027 = Unable to activate object (version 97)
        '31032 = Unable to create embedded object (version 97)
        'These mean that there is something wrong with the actual code.
        SomeNum = MsgBox("There is an issue with the code. Please contact the maker of this workbook.", vbOKOnly, "Error: " & ErrorNumber)
        errHandler = 1 'Exit the running macro
    Case 6, 7, 14, 28, 31001:
        '6 = Overflow
        '7 = Out of Memory
        '14 = Out of String Space (This can be caused by a string being too long as well)
        '28 = Out of Stack Space
        '31001 = Out of memory
        SomeNum = MsgBox("Excel could not allocate the memory it needed to run this macro. There are a couple ways to try to fix this:" & vbCrLf & _
            "1)Close Background Apps" & vbCrLf & "2) Restart Excel" & vbCrLf & "3) Restart your computer" & vbCrLf & _
            "4) Get 64-bit Windows Office on you computer (contact I.T.)", vbOKOnly, "Error: " & ErrorNumber)
        errHandler = 1 'Exit the running macro
    Case 9:
        '9 = subscript out of range
        SomeNum = MsgBox("Macro could not find something it was looking for." & vbCrLf & "Would you like the macro to try and finish? (This could be very bad).", vbYesNo, "Error: 9")
        If SomeNum = vbYes Then
            errHandler = 3
        Else
            errHandler = 1
        End If
    Case 11:
        '11 = Division by Zero
        SomeNum = MsgBox("Oh God you divided by 0!! The whole world is going to end now!!!", vbOKOnly, "Oh God Why!? Error: 11")
        errHandler = 1
    Case 18:
        '18 = User interrupt occurred
        'User is trying to stop the macro so stop it
        errHandler = 1
    Case 47, 48, 49, 51, 298, 325, 327, 335, 368, 453:
        '47 = too man DLL application clients
        '48 = Error in loading DLL
        '49 = Bad DLL calling convention
        '51 = Internal Error
        '298 = system DLL could not be loaded
        '325 - Invalid format in resource file
        '327 = Data value named was not found
        '335 = Could not access system registry
        '368 = The specified file is out of date. This program requires a newer version.
        '453 = Specified DLL function not found
        'Something is wrong with excel as an application,
        'It might be missing files or installed is incomplete
        SomeNum = MsgBox("Excel application cannot find required files it needs. It may not have installed correctly." _
            & vbCrLf & "First try to close and reopen the file, but if that doesn't work reinstall excel", vbOKOnly, "Error: " & ErrorNumber)
        errHandler = 2
    Case 52, 53, 54:
        '52 = Bad file name or number
        '53 = File Not found
        '54 = Bad file mode
        SomeNum = MsgBox("The file path enter is not possible or file was not found, please double check it and try again.", vbOKOnly, "Error: " & ErrorNumber)
        errHandler = 1
    Case 55:
        '55 = File already open
        'Nothing bad just trying to open something already open, just ignore it.
        errHandler = 3
    Case 57:
        '57 = Device I/O error
        SomeNum = MsgBox("Device I/O had an error, please ensure it is working properly or restart your computer and try again.", vbOKOnly, "Error: 57")
        errHandler = 2
    Case 58:
        '58 = file already exits
        SomeNum = MsgBox("A file by this name already exists in this folder, please change the name, the location or delete the old file before you try again.", vbOKOnly, "Error: 58")
        errHandler = 1
    Case 61:
        '61 = Disk full
        SomeNum = MsgBox("There is not enough diskspace to save required file.", vbOKOnly, "Error: 61")
        errHandler = 1
    Case 67:
        '67 = too many files
        SomeNum = MsgBox("You have too many files open and cannot open anymore. Close some and try again", vbOKOnly, "Error: 67")
        errHandler = 1
    Case 68, 71:
        '68 = Device unavailable
        '71 = Disk not ready
        SomeNum = MsgBox("The device you are trying to access is unavailable, check that it is working properly or your network access and try again.", vbOKOnly, "Error: " & ErrorNumber)
        errHandler = 1
    Case 70:
        '70 = Permission denied
        SomeNum = MsgBox("You do not have permission to write to this location. Please try again.", vbOKOnly, "Error: 70")
        errHandler = 1
    Case 75, 76:
        '75 = Path/file access error
        '76 = Path not found
        SomeNum = MsgBox("Excel cannot find/access specificed file.", vbOKOnly, "Error: " & ErrorNumber)
        errHandler = 1
    Case 95:
        '95 = User-defined error
        'User forced an error, try and skip the line
        errHandler = 3
    Case 320, 321:
        '320 = Can't use character device names in specified file names
        '321 = invalid file format
        SomeNum = MsgBox("Save name is invalid change it and try again.", vbOKOnly, "Error: " & ErrorNumber)
        errHandler = 1
        'Save Name or file name is invalid
    Case 322, 735:
        '322 = Can't Create necessary temporary file
        '735 = Can 't save file to Temp directory (version 97)
        SomeNum = MsgBox("Excel cannot create necessary temporary files, this could happen if you do not have enough disk space" & _
            "or do not have write access to you TEMP folder. Check these and try again.", vbOKOnly, "Error: 322")
        errHandler = 2
    Case 336, 337, 338, 363:
        '336 = ActiveX component not correctly registered
        '337 = ActiveX component not found
        '338 = ActiveX component did not correctly run
        '363 = Specificed ActiveX control not found
        SomeNum = MsgBox("There was an issue with the ActiveX Components. Close the workbook and try again. Otherwise contact the creator of this workbook.", vbOKOnly, "Error: " & ErrorNumber)
        errHandler = 2
    Case 360:
        '360 = object already loaded
        errHandler = 3
    Case 400, 402:
        '400 = Form already displayed; can't show modally (version 97)
        '402 = Code must close topmost modal form first (version 97)
        'There is already a user form open, it has to close before the new one can be closed.
        'Use the isUserFormLoaded to check each form and close it.

        errHandler = 0 'Do nothing because we closed all open forms
    Case 419:
        '419 = Permission to use object denied (version 97)
        SomeNum = MsgBox("You cannot access an object, check that the worksheet is not protected.", vbOKOnly, "Error: 419")
        errHandler = 1
    Case 449, 450, 1002, 1003:
        '449 = Argument not optional
        '450 = Wrong number of arguments
        '1002 = Missing required argument argumentname
        '1003 = Invalid number of arguments (versions 5.0 and 7.0)
        SomeNum = MsgBox("You or the macro have tried to use a function incorrectly, double check what you entered and try again.", vbOKOnly, "Error: " & ErrorNumber)
        errHandler = 1
    Case 481, 485:
        '481 = Invalid picture (version 97)
        '485 = Invalid picture type (version 97)
        SomeNum = MsgBox("This picture is invalid, please try a different one.", vbOKOnly, "Error: " & ErrorNumber)
        errHandler = 1
    Case 482, 483, 484, 486:
        '482 = Printer error (version 97)
        '483 = Printer driver does not support specified property (version 97)
        '484 = Problem getting printer information from the system. Make sure the printer is set up correctly
        '486 = Can 't print form image to this type of printer (version 97)
        SomeNum = MsgBox("There is an error with your printer, please check that it is working properly and can print this format.", vbOKOnly, "Error: " & ErrorNumber)
        errHandler = 1
    Case Else
        'An Error not defined has occured.
        'Since we don't know what it is end the macro
        SomeNum = MsgBox("Error!" & vbCrLf & Error(ErrorNumber), vbOKOnly, "Error: " & ErrorNumber)
        errHandler = 1
End Select
Exit Function
errorHandler:
    Dim ErrorAction As Integer
    ErrorAction = errHandler(Err.Number)
    Select Case ErrorAction
        Case 0:
            'Ignore Error
            Resume
        Case 1:
            'End Macro
            Exit Function
        Case 2:
            ErrorAction = MsgBox("Do you wish to close excel?", vbYesNo, "Close?")
            If ErrorAction = vbYes Then
                ThisWorkbook.Close
            Else
                Exit Function
            End If
        Case 3:
            'Skip error causing line
            Resume Next
        Case Else
            Exit Function
    End Select

   On Error GoTo 0
   Exit Function

errHandler_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure errHandler of Module ErrorHandling"
End Function
'EXAMPLE SUB
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
Function IsUserFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object
    IsUserFormLoaded = False
    For Each UForm In VBA.UserForms
        If UForm.Name = UFName Then
            IsUserFormLoaded = True
            Exit For
        End If
    Next
Exit Function
errorHandler:
    Dim ErrorAction As Integer
    ErrorAction = errHandler(Err.Number)
    Select Case ErrorAction
        Case 0:
            'Ignore Error
            Resume
        Case 1:
            'End Macro
            Exit Function
        Case 2:
            ErrorAction = MsgBox("Do you wish to close excel?", vbYesNo, "Close?")
            If ErrorAction = vbYes Then
                ThisWorkbook.Close
            Else
                Exit Function
            End If
        Case 3:
            'Skip error causing line
            Resume Next
        Case Else
            Exit Function
    End Select
End Function
