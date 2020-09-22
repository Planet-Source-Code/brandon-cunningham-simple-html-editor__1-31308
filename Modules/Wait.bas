Attribute VB_Name = "WaitModule"
'This code is Intellectual Propriety of Brandon Cunningham.
Sub Wait(Time As Integer)
  Dim Count As Long
  Count = Timer
  Do While Timer - Count < Time
    DoEvents
  Loop
End Sub
