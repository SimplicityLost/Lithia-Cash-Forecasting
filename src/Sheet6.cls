VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call RawReformatter
    MsgBox ("The data is ready for categorization")
End Sub

Private Sub CommandButton2_Click()
    Application.ScreenUpdating = False
    Call Categorize
    Application.ScreenUpdating = True
    MsgBox ("Categorization complete.")
    
End Sub
