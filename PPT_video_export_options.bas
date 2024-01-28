Attribute VB_Name = "Module1"
Sub MP4_720_60fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 720p 60.mp4", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=720, _
    FramesPerSecond:=60, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MP4_720_120fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 720p 120.mp4", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=720, _
    FramesPerSecond:=120, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MP4_1080_60fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 1080p 60.mp4", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=1080, _
    FramesPerSecond:=60, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MP4_1080_120fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 1080p 120.mp4", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=1080, _
    FramesPerSecond:=120, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MP4_4K_30fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 4K 30.mp4", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=2160, _
    FramesPerSecond:=30, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MP4_4K_60fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 4K 60.mp4", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=2160, _
    FramesPerSecond:=60, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MP4_4K_120fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 4K 120.mp4", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=2160, _
    FramesPerSecond:=120, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub WMV_720_60fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 720p 60.wmv", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=720, _
    FramesPerSecond:=60, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub WMV_720_120fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 720p 120.wmv", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=720, _
    FramesPerSecond:=120, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub WMV_1080_60fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 1080p 60.wmv", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=1080, _
    FramesPerSecond:=60, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub WMV_1080_120fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 1080p 120.wmv", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=1080, _
    FramesPerSecond:=120, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub WMV_4K_30fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 4K 30.wmv", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=2160, _
    FramesPerSecond:=30, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub WMV_4K_60fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 4K 60.wmv", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=2160, _
    FramesPerSecond:=60, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub WMV_4K_120fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 4K 120.wmv", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=2160, _
    FramesPerSecond:=120, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MOV_480_30fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 480p 30.mov", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=480, _
    FramesPerSecond:=30, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MOV_480_60fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 480p 60.mov", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=480, _
    FramesPerSecond:=60, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MOV_480_120fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 480p 120.mov", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=480, _
    FramesPerSecond:=120, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MOV_720_30fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 720p 30.mov", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=720, _
    FramesPerSecond:=30, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MOV_720_60fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 720p 60.mov", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=720, _
    FramesPerSecond:=60, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MOV_720_120fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 720p 120.mov", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=720, _
    FramesPerSecond:=120, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MOV_1080_30fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 1080p 30.mov", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=1080, _
    FramesPerSecond:=30, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MOV_1080_60fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 1080p 60.mov", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=1080, _
    FramesPerSecond:=60, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MOV_1080_120fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 1080p 120.mov", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=1080, _
    FramesPerSecond:=120, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MOV_4K_30fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 4K 30.mov", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=2160, _
    FramesPerSecond:=30, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MOV_4K_60fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 4K 60.mov", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=2160, _
    FramesPerSecond:=60, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
Sub MOV_4K_120fps()
If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    ActivePresentation.CreateVideo FileName:=Environ("USERPROFILE") & "\Desktop\PowerPoint Video 4K 120.mov", _
    UseTimingsAndNarrations:=True, _
    VertResolution:=2160, _
    FramesPerSecond:=120, _
    Quality:=100
Else
    MsgBox "There is another conversion to video in progress"
End If
End Sub
