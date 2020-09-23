Attribute VB_Name = "modWMI_CPU"
Option Explicit

Function GetProcessor() As String
    Dim objWMIService, colItems, objItem

    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select Name from Win32_Processor", , 48)
    
    For Each objItem In colItems
        GetProcessor = objItem.Name
    Next
End Function

