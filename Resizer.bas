Attribute VB_Name = "Module1"
Public masterImage As Shape
Public master_pf As PictureFormat
Public currentSlide As Integer

' This function sets a master image and saves its properties for later use
Public Sub SelectMaster_Click()

    On Error GoTo ErrorHandler
   
    Set masterImage = ActiveWindow.Selection.ShapeRange(1)
    Set master_pf = masterImage.PictureFormat()

    currentSlide = ActiveWindow.View.Slide.SlideNumber
   
   
    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case -2147188160
            MsgBox "No Shape selected to create Master"
    End Select

       
   
   
End Sub

'Resize and Reposition chosen Object
Public Sub ResizeReposition_Click()

    On Error GoTo ErrorHandler
   
    'Prompt the user to select the image to be resized and repositioned
    Set selectedImage = ActiveWindow.Selection.ShapeRange(1)
    selectedImage.LockAspectRatio = msoFalse

    'Resize the selected image based on the dimensions of the master image
    selectedImage.Width = masterImage.Width
    selectedImage.Height = masterImage.Height

    'Reposition the selected image based on the position of the master image
    selectedImage.Top = masterImage.Top
    selectedImage.Left = masterImage.Left
   
   



    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 91
            MsgBox "No Master Shape selected to Resize"
    End Select
   
End Sub


'Resize Reposition and Crop chosen Object
Public Sub Resize_and_Crop()

    On Error GoTo ErrorHandler

    'Prompt the user to select the image to be resized and repositioned
    Set selectedImage = ActiveWindow.Selection.ShapeRange(1)
   
    selectedImage.LockAspectRatio = msoFalse
   

    Set selected_pf = selectedImage.PictureFormat()
   
    selected_pf.CropBottom = master_pf.CropBottom
    selected_pf.CropLeft = master_pf.CropLeft
    selected_pf.CropRight = master_pf.CropRight
    selected_pf.CropTop = master_pf.CropTop
   




    'Resize the selected image based on the dimensions of the master image
    selectedImage.Width = masterImage.Width
    selectedImage.Height = masterImage.Height

    'Reposition the selected image based on the position of the master image
    selectedImage.Top = masterImage.Top
    selectedImage.Left = masterImage.Left
   
   
    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 91
            MsgBox "No Master Shape selected to Resize"
    End Select
   
   
End Sub


'Resize Reposition all mso pricture objects on selected slides
Public Sub ResizeRepositionAll()

    On Error GoTo ErrorHandler

Dim sld As PowerPoint.Slide

For Each sld In ActiveWindow.Selection.SlideRange
    Dim shp As Shape
    For Each shp In sld.Shapes
        If shp.Type = msoPicture Then
            Set selectedImage = shp
           
            'Resize the selected image based on the dimensions of the master image
            selectedImage.LockAspectRatio = msoFalse
            selectedImage.Width = masterImage.Width
            selectedImage.Height = masterImage.Height
       
            'Reposition the selected image based on the position of the master image
            selectedImage.Top = masterImage.Top
            selectedImage.Left = masterImage.Left
           
           
           
           
        End If
    Next shp
Next sld


    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 91
            MsgBox "No Master Shape selected to Resize"
    End Select

End Sub


'Resize Reposition and Crop all mso pricture objects on selected slides
Public Sub Resize_and_Crop_all()

    On Error GoTo ErrorHandler

Dim sld As PowerPoint.Slide


For Each sld In ActiveWindow.Selection.SlideRange
    Dim shp As Shape
    For Each shp In sld.Shapes
        If shp.Type = msoPicture Then
            Set selectedImage = shp
           
            Set selected_pf = selectedImage.PictureFormat()
   
            selectedImage.LockAspectRatio = msoFalse
           
            selected_pf.CropBottom = master_pf.CropBottom
            selected_pf.CropLeft = master_pf.CropLeft
            selected_pf.CropRight = master_pf.CropRight
            selected_pf.CropTop = master_pf.CropTop
           
            'Resize the selected image based on the dimensions of the master image
            selectedImage.LockAspectRatio = msoFalse
            selectedImage.Width = masterImage.Width
            selectedImage.Height = masterImage.Height
       
            'Reposition the selected image based on the position of the master image
            selectedImage.Top = masterImage.Top
            selectedImage.Left = masterImage.Left
           
           
           
           
        End If
    Next shp
Next sld


    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 91
            MsgBox "No Master Shape selected to Resize"
    End Select

End Sub

'Send back all image objects
Public Sub sendBack()

Dim iSlide As Slide
Dim iShapes As Shapes
Dim iShape As Shape

For Each iSlide In ActiveWindow.Selection.SlideRange
    Set iShapes = iSlide.Shapes
    For i = iShapes.Count To 1 Step -1
        Set iShape = iShapes(i)
        If iShape.Type = msoPicture Then
            iShape.ZOrder msoSendToBack
        End If
    Next i
Next iSlide

For Each iSlide In ActiveWindow.Selection.SlideRange
    Set iShapes = iSlide.Shapes
    For i = iShapes.Count To 1 Step -1
        Set iShape = iShapes(i)
        If iShape.Type = msoPicture Then
            iShape.ZOrder msoSendToBack
        End If
    Next i
Next iSlide

End Sub

'Send to front all image objects
Public Sub sendFront()

Dim oSlide As Slide
Dim oShapes As Shapes
Dim oShape As Shape

For Each oSlide In ActiveWindow.Selection.SlideRange
    Set oShapes = oSlide.Shapes
    For i = oShapes.Count To 1 Step -1
        Set oShape = oShapes(i)
        If oShape.Type = msoPicture Then
            oShape.ZOrder msoBringToFront
        End If
    Next i
Next oSlide

For Each oSlide In ActiveWindow.Selection.SlideRange
    Set oShapes = oSlide.Shapes
    For i = oShapes.Count To 1 Step -1
        Set oShape = oShapes(i)
        If oShape.Type = msoPicture Then
            oShape.ZOrder msoBringToFront
        End If
    Next i
Next oSlide


End Sub

'Automatic Slide createion for each selected image from file explorer. Title will be created
Sub InsertPicturesAndCreateSlides()

    On Error GoTo ErrorHandler


Dim myPresentation As PowerPoint.Presentation
Dim mySlide As PowerPoint.Slide
Dim myLayoutSlide As PowerPoint.Slide
Dim myShape As PowerPoint.Shape
Dim myPicture As String
Dim myTextbox As PowerPoint.Shape
Dim shapeCount As Integer
Dim currentSlideIndex As Integer 'variable to hold the index of the current slide


Dim sngHeight As Single
  Dim sngWidth As Single
  Dim bTemp As Boolean

  bTemp = False ' by default

  With ActivePresentation.PageSetup
    sngHeight = .slideHeight
    sngWidth = .slideWidth
  End With


Set myPresentation = ActivePresentation
Set myLayoutSlide = myPresentation.Slides(currentSlide)


currentSlideIndex = ActiveWindow.Selection.SlideRange(1).SlideIndex

'Open the file dialog to select pictures
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = True
    .Show
    'Loop through each selected picture
    For i = 1 To .SelectedItems.Count
        myPicture = .SelectedItems(i)
        'Create a new slide
        Set mySlide = myPresentation.Slides.AddSlide(currentSlideIndex + 1, myLayoutSlide.CustomLayout)
        'Insert the picture on the slide
       
       
        If mySlide.Shapes.Count >= 1 Then
           
 
        shapeCount = mySlide.Shapes.Count

        For k = shapeCount To 1 Step -1
            If mySlide.Shapes(k).ZOrderPosition <> shapeCount Then
                mySlide.Shapes(k).Delete
            End If
        Next k
       
        Set Title = mySlide.Shapes.Item(1)
       
        text_str11 = Replace(myPicture, ".png", "")
        text_str10 = Replace(text_str11, ".jpg", " ")
        text_str = Replace(text_str10, "_", " ")
        Title.TextFrame.TextRange.Text = Right(text_str, Len(text_str) - InStrRev(text_str, "\", -1, vbTextCompare))

       
        End If
       
        If mySlide.Shapes.Count = 0 Then
        Set Title = mySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 26.456, 15.11, 731.48, 41.574)
        text_str11 = Replace(myPicture, ".png", "")
        text_str10 = Replace(text_str11, ".jpg", " ")
        text_str = Replace(text_str10, "_", " ")
        Title.TextFrame.TextRange.Text = Right(text_str, Len(text_str) - InStrRev(text_str, "\", -1, vbTextCompare))
        Title.TextFrame.TextRange.Font.Name = "Calibri(Headings)"
        Title.TextFrame.TextRange.Font.Size = 32
        Title.TextFrame.TextRange.Font.Color.RGB = RGB(135, 17, 98)

       
       
        End If
       
       ' Add the new picture shape
        Set new_pic = mySlide.Shapes.AddPicture(myPicture, msoFalse, msoTrue, 0, 0)
       
       ' Position the picture shape to match the master image position and size
        new_pic.Left = masterImage.Left
        new_pic.Top = masterImage.Top
        new_pic.Height = masterImage.Height

       
    Next i
End With

    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case -2147188160
            MsgBox "Please select Master Shape first"
    End Select


End Sub

' Automatic Slide creation for each selected image from file explorer. No Title
Sub InsertPictures_NT()

    On Error GoTo ErrorHandler
    
    ' Declare variables
    Dim myPresentation As PowerPoint.Presentation
    Dim mySlide As PowerPoint.Slide
    Dim myLayoutSlide As PowerPoint.Slide
    Dim myShape As PowerPoint.Shape
    Dim myPicture As String
    Dim myTextbox As PowerPoint.Shape
    Dim shapeCount As Integer
    Dim currentSlideIndex As Integer ' Variable to hold the index of the current slide
    Dim sngHeight As Single
    Dim sngWidth As Single
    Dim bTemp As Boolean
  
    bTemp = False ' By default
    
    ' Get the slide dimensions
    With ActivePresentation.PageSetup
        sngHeight = .slideHeight
        sngWidth = .slideWidth
    End With
    
    ' Set references to the active presentation and layout slide
    Set myPresentation = ActivePresentation
    Set myLayoutSlide = myPresentation.Slides(currentSlide)
    
    ' Get the index of the current slide
    currentSlideIndex = ActiveWindow.Selection.SlideRange(1).SlideIndex
    
    ' Open the file dialog to select pictures
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        .Show
        
        ' Loop through each selected picture
        For i = 1 To .SelectedItems.Count
            myPicture = .SelectedItems(i)
            
            ' Create a new slide
            Set mySlide = myPresentation.Slides.AddSlide(currentSlideIndex + 1, myLayoutSlide.CustomLayout)
            
            ' Insert the picture on the slide
           
            ' Delete existing shapes on the slide if there are any
            If mySlide.Shapes.Count >= 1 Then
                shapeCount = mySlide.Shapes.Count
                
                For k = shapeCount To 1 Step -1
                    If mySlide.Shapes(k).ZOrderPosition <> shapeCount Then
                        mySlide.Shapes(k).Delete
                    End If
                Next k
            End If
            
            ' Add the new picture shape
            Set new_pic = mySlide.Shapes.AddPicture(myPicture, msoFalse, msoTrue, 0, 0)
            
            ' Position the picture shape to match the master image position and size
            new_pic.Left = masterImage.Left
            new_pic.Top = masterImage.Top
            new_pic.Height = masterImage.Height
        Next i
    End With
    
    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case -2147188160
            MsgBox "Please select Master Shape first"
    End Select
End Sub

'Delete all Image objects on selected slides
Sub Delete()

Dim aSlide As Slide
Dim aShapes As Shapes
Dim aShape As Shape

For Each aSlide In ActiveWindow.Selection.SlideRange
    Set aShapes = aSlide.Shapes
    For i = aShapes.Count To 1 Step -1
        Set aShape = aShapes(i)
        If aShape.Type = msoPicture Then
            aShape.Delete
        End If
    Next i
Next aSlide

End Sub


'Find and replace
Sub ReplaceTextInSlides()
    Dim sld As Slide
    Dim shp As Shape
    Dim findText As String
    Dim replaceText As String
   
    findText = Trim(InputBox("Enter the text you want to find:"))
    replaceText = Trim(InputBox("Enter the text you want to replace it with:"))
   
    For Each sld In ActiveWindow.Selection.SlideRange
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    shp.TextFrame.TextRange.Text = Replace(shp.TextFrame.TextRange.Text, findText, replaceText)
                End If
            End If
        Next shp
    Next sld
End Sub
Sub CreateTextBox()

    Dim mySlide As Slide
    Dim myShape As Shape
   
    Set mySlide = ActiveWindow.View.Slide
   
    Set myShape = mySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, mySlide.Master.Width / 6, mySlide.Master.Height - 50, mySlide.Master.Width / 4, 50)
   
    myShape.Fill.ForeColor.RGB = RGB(255, 255, 255) 'White background
    myShape.Line.ForeColor.RGB = RGB(135, 17, 98) 'Border color
    myShape.Line.Weight = 1 'Border line width in points
    myShape.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0) 'Black text
   
  
    myShape.TextFrame.TextRange.Text = "Your text here"
   
End Sub









