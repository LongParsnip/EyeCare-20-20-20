Attribute VB_Name = "A_Startup"
Sub Main()

    'Only allow 1 instance to run.
    If App.PrevInstance Then
        AppActivate (App.Title)
    Else
        frmMain.Show
    End If
    
End Sub

