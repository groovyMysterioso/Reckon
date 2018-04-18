Attribute VB_Name = "TimeReckon"
Option Explicit

Dim progressLastChecked, progressLastUpdated
Dim progressTimeSample As New Collection
Dim lastEstimate

Public Function Reckon(index, count)

    Dim timeCycle
       
    timeCycle = IIf(index < 1, 0, (Timer - progressLastChecked)) ' a timeCycle is the space from one iteration to the next
        
    If progressTimeSample.count > 5 Then progressTimeSample.Remove (1)
        progressTimeSample.Add (timeCycle)
    
    If Timer - progressLastUpdated > 1 Then
        
        lastEstimate = EstimateString(count, index)
        progressLastUpdated = Timer
    
    End If
        
    Reckon = lastEstimate
    progressLastChecked = Timer
    
End Function

Function AverageCycle(count, index)

    Dim cycleSum, sampleCycle
    
    For Each sampleCycle In progressTimeSample
        cycleSum = cycleSum + sampleCycle
    Next
        AverageCycle = Split(Format(cycleSum / progressTimeSample.count * (count - index) / 86400, "n s"))
        
End Function

Function EstimateString(count, index)
    
    Dim timeArray
    timeArray = AverageCycle(count, index)
    If timeArray(0) = 0 And timeArray(1) = 0 Then
        EstimateString = "Calculating"
    Else
        EstimateString = IIf(timeArray(0) <> 0, timeArray(0) & " minutes and  ", "") & timeArray(1) & " seconds remaining"
    End If
    
End Function

