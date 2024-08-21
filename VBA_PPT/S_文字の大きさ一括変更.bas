Attribute VB_Name = "S_•¶š‚Ì‘å‚«‚³ˆêŠ‡•ÏX"
Option Explicit

Sub •¶š‚Ì‘å‚«‚³ˆêŠ‡•ÏX()
    Uf_•¶š‚Ì‘å‚«‚³ˆêŠ‡•ÏX.Show
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
