Attribute VB_Name = "Modmain"


' This is the module containing the main game loop an various important Procedures





'********************************************************************************
'********************************************************************************
' This is the main game loop. Once the stage is set up the control is passed to this sub everything is
'controlled from here. Befor this sub is called some initialisatins are performed in the modinit module


Public Sub Gameloop()




' On GETTICKCOUNT
'this function  returns the no of seconds passed after last midnight
' this has no such utility as such but this is a very useful function coz you can know the time taken by a
'operation by calling this function twice- once before opp. and one after the opp. and subtracting the values.
'this is what is done here but a bit diff. way

'T2 updates whenever the do loop executes. but T1 updates only when game-frame runs
' if a frame completes in very less time ie. less than Fpslimiter then the prog waits till time becomes fpslimiter
' So Frames Per Second Is limited to a certain value

'||||||||||||||||||||||||\_____________/||||||||||||||||||||||||||||||||||||\


' Update T1 for 1st time
T1 = GetTickCount  'for frame limit






Do

'Update T2 every time Do loops
T2 = GetTickCount                                                     'for frame limit

' Check if time bet frames is greater than Fpslimiter value
If (T2 - T1) >= FpsLimiter Then                                 'if frame rate is under control




tFps = tFps + 1                                                         'tFps stores the Frames executed In the current second
                                                    'It resets to Zero every new sec. after transferring its value to the Fps variable


Call CheckKeys                                     '   Call the sub to Check some basic Key presses Like Escape Key
                                                               ' this thing Goes evem if game is paused



' See if the game is in Running mode
If Gamestate = running Then                       'If yes Then Proceed
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


' Call the sub which checks for things to be done every frame
Call Regularevents


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If



                                                                      ' the menu also has to be updated every time
Menu.Update

' Call the sub to Render the canvas ie copy the buffer to somethin you can see
Call RenderCanvas

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' Update T1
T1 = GetTickCount
End If                                                            'End if For frme checker


' Do keyboard and mouse events
DoEvents

Loop                                                             ' loop the do loop
End Sub
'********************************************************************************
'********************************************************************************
' This sub renders the canves ie. copys the buffer(frmpics.buffer picturebox ) to the canvas
'(the canvas picturebox present in the main form) . the user can see the canvas and not the buffer.

'Double Buffering- It is done to eleminate Flickering which happens if graphics are erased and redrawn each time
                           'So what we can do is to create a buffer erase and redraw it every time and the copy it to
                            'the canvas(actually visible to the user). the canvas is never erased so no flickering



Private Sub RenderCanvas()
' If you dont know what bitblt is or how it works the see some other articles on it
' i can't explain every thing


BitBlt dhoom1.Canvas.hdc, 0, 0, Canvas_width, Canvas_height, Frmpics.Buffer.hdc, 0, 0, vbSrcCopy
Frmpics.Buffer.Cls                                                    ' clear he buffer after it has been copyed to the canvas



End Sub
'********************************************************************************
'********************************************************************************
' Events done every frame
Private Sub Regularevents()

Call Updateclasses                                          ' All classes have to be updated everytime the frame executes

End Sub
'********************************************************************************
'********************************************************************************
' Updates the instances of the classes that are present
' All class-objects are updates from here except the menu object which is updated directly
' from the mainloop since it has to be updated even if the game is paused
Private Sub Updateclasses()

Road.Update                                             '<-- Update the road object uhich is a instance of clsroad



For i = 1 To 10                                         ' <--there are 10 instances of vehicleclass hence 1 to 10
If Vehicle(i).Alive = True Then                   '<-- if vehicle is not alive then thee is no need of updating so
                                                                  '<--first check if the vehicle is alive.

Vehicle(i).Update                                        ' <--Call the Update method to update
End If                                                          '<-- And if for vehicle alive check
Next                                                             ' <--See next vehicle

Player.Update                                             ' <--there is one player and it has to be alive . so update it
Display.Update                                             ' <--Display update


End Sub
'********************************************************************************
'********************************************************************************
