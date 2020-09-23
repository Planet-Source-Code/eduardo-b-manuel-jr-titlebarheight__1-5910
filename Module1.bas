Attribute VB_Name = "ModTitBar"
Option Explicit
'This is a very small function which is
'potentially useful for those who know where/how to use it.
'I used this function on several of my projects already
'if you're wondering if this code has any use at all.
'I'll post them later. Now back to the function.
'It returns the size of the titlebar of the form
'passed on the first parameter in any scaling scheme
'(even user defined) that you supply in the second parameter
'(if you do not supply a value, the ScaleMode of the
'target form--Target.scalemode--is used).
'***Please Give me credit if you use this
'********* D o n G ************'
Public Function TitleBarHeight(Target As Form, Optional RetScaleMode)
Dim RetScale As Integer
'For convenience, I used the With statement. Look it up in the help file if you don't know how to use it.
With Target
    'Check if a return scalemode is supplied. if not, use Target's scalemode(Target.ScaleMode)
    If IsMissing(RetScaleMode) Then RetScaleMode = .ScaleMode
    'The ScaleY function cannot return a value in a user defined scaling scheme.
    'The following line tests if RetScaleMode is equal to 0 or User. If such is the case
    'we will use a scalemode value of 1(Twips) to avoid the funky-shitty-ultra-crap runtime-error
    'because ScaleY does not return a value in User mode(Shame,Really)
    If RetScaleMode = 0 Then RetScale = 1 Else RetScale = RetScaleMode
    'Subtract the height measured from Titlebar to bottom border using the ScaleY function to obtain the desired(correct) value
    TitleBarHeight = .ScaleY(.Height, 1, RetScale) - .ScaleY(.ScaleHeight, .ScaleMode, RetScale)
    'if the Scalemode you wish to return is 0 or User mode then convert the result of the operation above to User mode
    If RetScaleMode = 0 Then TitleBarHeight = (.ScaleHeight / (.Height - TitleBarHeight)) * TitleBarHeight
End With
End Function
