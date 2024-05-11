VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Task Manager"
   ClientHeight    =   13590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9405
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BackupString1 As String
Dim BackupString2 As String
Dim BackupString3 As String
Dim BackupString4 As String
Dim BackupString5 As String

Private Sub CommandButton1_Click()
    Modul4.getProcess
 
    
    
End Sub

Private Sub CommandButton2_Click()
    Modul4.killProcess
    
End Sub

Private Sub CommandButton5_Click()
   OpenTask = CreateObject("WScript.Shell").Run(TextBox1.Value)
   ListBox1.AddItem (TextBox1.Value)
   If (OptionButton4.Value = True) Then
    MsgBox ("Prozess: " + TextBox1.Value + " wurde erstellt")
   ElseIf OptionButton4.Value = False And OptionButton3.Value = True Then
   
    MsgBox ("Process: " + TextBox1.Value + " is created!")
   Else
     MsgBox ("Process: " + TextBox1.Value + " is created!")
   End If
   
   CommandButton2.Enabled = True
   TextBox1.Value = ""
   
End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub OptionButton3_Click()
    UserForm2.CommandButton1.caption = BackupString1
    UserForm2.CommandButton2.caption = BackupString2
    UserForm2.CommandButton3.caption = BackupString3
    UserForm2.CommandButton4.caption = BackupString4
    UserForm2.CommandButton5.caption = BackupString5
End Sub

Private Sub OptionButton4_Click()
    UserForm2.CommandButton1.caption = "Bekomme Prozesse"
    UserForm2.CommandButton2.caption = "Beende ausgewählter Prozess"
    UserForm2.CommandButton3.caption = "Beenden"
    UserForm2.CommandButton4.caption = "Lösche Prozessliste"
    UserForm2.CommandButton5.caption = "Starte Prozess"
End Sub

Private Sub UserForm_Initialize()
    Label1.ForeColor = RGB(0, 150, 0)
    Label2.ForeColor = RGB(0, 150, 0)
    Label3.ForeColor = RGB(0, 150, 0)
    UserForm2.BackColor = RGB(0, 0, 0)
    Frame1.BackColor = RGB(0, 0, 0)
    Frame1.ForeColor = RGB(255, 255, 255)
    Frame2.BackColor = RGB(0, 0, 0)
    Frame2.ForeColor = RGB(255, 255, 255)
    Frame3.BackColor = RGB(0, 0, 0)
    Frame3.ForeColor = RGB(255, 255, 255)
    ListBox1.BackColor = RGB(0, 0, 0)
    ListBox1.ForeColor = RGB(0, 150, 0)
    CommandButton2.Enabled = False
    OptionButton1.Value = True
    BackupString1 = CommandButton1.caption
    BackupString2 = CommandButton2.caption
    BackupString3 = CommandButton3.caption
    BackupString4 = CommandButton4.caption
    BackupString5 = CommandButton5.caption
    OptionButton5.Value = False
    Modul4.App_loop
    
    
End Sub

Private Sub CommandButton3_Click()
    End
End Sub

Private Sub CommandButton4_Click()
    ListBox1.Clear
    CommandButton2.Enabled = False
    
End Sub

Private Sub OptionButton1_Click()
    UserForm2.BackColor = RGB(0, 0, 0)
    Frame1.BackColor = RGB(0, 0, 0)
    Frame1.ForeColor = RGB(0, 150, 0)
    ListBox1.BackColor = RGB(0, 0, 0)
    ListBox1.ForeColor = RGB(0, 150, 0)
 
    
    
End Sub

Private Sub OptionButton2_Click()
    UserForm2.BackColor = RGB(255, 255, 255)
    Frame1.BackColor = RGB(255, 255, 255)
    Frame1.ForeColor = RGB(0, 0, 0)
    ListBox1.BackColor = RGB(255, 255, 255)
    ListBox1.ForeColor = RGB(0, 0, 0)
End Sub

Private Sub UserForm_Click()

End Sub

