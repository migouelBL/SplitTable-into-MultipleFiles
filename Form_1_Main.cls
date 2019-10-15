Option Compare Database

Public Function StartUp()
Application.Echo False

DoCmd.NavigateTo "acNavigationCategoryObjectType"
DoCmd.RunCommand acCmdWindowHide
DoCmd.ShowToolbar "Ribbon", acToolbarNo
DoCmd.OpenForm "1_Main"

Application.Echo True
End Function

Public Function Unload()
Dim aob As AccessObject
Application.Echo False

'Close Forms-----------------
For Each aob In CurrentProject.AllForms
    DoCmd.Close acForm, aob.Name, acSaveNo
Next aob

DoCmd.SelectObject acForm, , True
DoCmd.ShowToolbar "Ribbon", acToolbarYes
Application.Echo True
End Function
