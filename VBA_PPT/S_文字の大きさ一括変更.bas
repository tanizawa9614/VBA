Attribute VB_Name = "S_文字の大きさ一括変更"
Option Explicit

Sub 文字の大きさ一括変更()
    Uf_文字の大きさ一括変更.Show
End Sub

Public Sub ChangeStringSize(ByVal S_Size As Double)
    
    Dim sld As Slide
    Dim shp As Shape
    Dim S As Double
    
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            With shp
                If .Type <> msoPlaceholder Then
                    If .HasTextFrame Or .Type = msoGroup Then
                        With .TextFrame2.TextRange.Font
                            .size = S_Size
                        End With
                    End If
                End If
            End With
        Next shp
    Next sld
    
End Sub
