Attribute VB_Name = "ReckonTest"

Sub ReckonTest()
  Dim x As Integer
    Dim MyTimer As Double
    Dim Max
    Dim MagicNumber
    Max = 900
         
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo HandleCancel:
   
    For x = 0 To Max
         
        MyTimer = Timer
        Do: Loop While Timer - MyTimer < 0.03
        DoEvents
         
        Application.StatusBar = Reckon(x, Max)
        
    Next x

HandleCancel:
    Application.StatusBar = False
End Sub


