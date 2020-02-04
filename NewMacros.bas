Attribute VB_Name = "NewMacros"
Sub insertImage()
'
' insertImage Macro
'
Dim i As Integer
i = 1




Dim bookmark As Bookmarks
Dim sFolder As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a Folder"
        If .Show = -1 Then
            sFolder = .SelectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then

 Dim bmk As bookmark
    Dim msg As String
    For Each bmk In ActiveDocument.Range.Bookmarks
      
       
       vS = ActiveDocument.Bookmarks(bmk.Name).Range _
        .InlineShapes.AddPicture(sFolder & "\" & i & ".jpg").ScaleHeight = 50
       
       i = i + 1
    
    Next bmk

   
    
Call SetupAllPictureSize

    End If

End Sub


Sub SetupAllPictureSize()
  Dim objInlineShape As InlineShape
  Dim objShape As Shape
  
  For Each objInlineShape In ActiveDocument.InlineShapes
    objInlineShape.Height = 50
    objInlineShape.Width = 50
  Next objInlineShape

  For Each objShape In ActiveDocument.Shapes
    objShape.Height = 50
    objShape.Width = 50
  Next objShape
End Sub

