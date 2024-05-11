Attribute VB_Name = "Modul4"
Public Static Sub getProcess()
    Dim Process As Object
 
    
For Each Process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
    
    
            
    If Process.caption = "System" Then
        e = 1
    ElseIf Process.caption = "System Idle Process" Then
        e = 2
   
    Else
        UserForm2.ListBox1.AddItem (Process.caption)
    End If
  
Next
UserForm2.CommandButton2.Enabled = True
End Sub
Public Static Sub killProcess()
    Dim LBItem As Long
    For LBItem = 0 To UserForm2.ListBox1.ListCount - 1
        If UserForm2.ListBox1.Selected(LBItem) = True Then
            TaskKill = CreateObject("WScript.Shell").Run("taskkill /f /im " & ListBox1.List(LBItem), 0, True)
            temp = UserForm2.ListBox1.List(LBItem)
            UserForm2.ListBox1.RemoveItem (LBItem)
            If (UserForm2.OptionButton4.Value = True) Then
                 MsgBox (temp + " wurde beendet")
            ElseIf UserForm2.OptionButton4.Value = False And UserForm2.OptionButton3.Value = True Then
                
                  MsgBox (temp + " was killed")
            Else
                   MsgBox (temp + " was killed")
                End If
                        
                         
                     End If
Next
End Sub
Public Static Sub clearProcesslist()
    UserForm2.ListBox1.Clear
    UserForm2.CommandButton2.Enabled = False
End Sub
Public Static Sub App_loop()
    If UserForm2.OptionButton5 = False Then
    
    Else
    ' Reapeat every 5 Minutes
    Application.OnTime Now() + TimeValue("00:4:59"), "getProcess"
    Application.OnTime Now() + TimeValue("00:5:00"), "clearProcesslist"
    
    End If
    ' Create the Loop every 5 Minutes and 2 Seconds
    Application.OnTime Now() + TimeValue("00:5:02"), "App_loop"
End Sub

