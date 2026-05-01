Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Shapes
Imports System.Windows
Imports System.Collections.Generic
Imports System.Linq

Public Module ChartHelper

    Public Sub DrawLineChart(canvas As Canvas,
                              data As Queue(Of Double),
                              maxValue As Double,
                              lineColor As Color,
                              Optional showGrid As Boolean = False,
                              Optional showLabels As Boolean = False,
                              Optional unit As String = "%")

        canvas.Children.Clear()
        If data Is Nothing OrElse data.Count < 2 Then Return

        Dim w = canvas.ActualWidth
        Dim h = canvas.ActualHeight
        If w <= 1 OrElse h <= 1 Then Return

        Dim values = data.ToArray()
        Dim n = values.Length
        Dim actualMax = If(maxValue > 0, maxValue, Math.Max(values.Max(), 0.1))

        Dim padLeft = If(showLabels, 32.0, 4.0)
        Dim padRight = 4.0
        Dim padTop = 6.0
        Dim padBottom = 4.0
        Dim chartW = w - padLeft - padRight
        Dim chartH = h - padTop - padBottom

        ' ChartHelper.vb içindeki DrawLineChart metodunun ilgili kısmı:

        If showGrid Then
            ' DÜZELTME: 3 yerine 4'e kadar saydırıyoruz ki %100 çizgisi de çizilsin
            For i As Integer = 1 To 4
                Dim yPos = h - (h * (i / 4.0))
                If yPos < 0 Then yPos = 0

                ' Arka plan kesik çizgileri
                Dim gridLine As New Shapes.Line With {
                    .X1 = 0, .Y1 = yPos, .X2 = w, .Y2 = yPos,
                    .Stroke = New SolidColorBrush(Color.FromArgb(40, 200, 200, 255)),
                    .StrokeDashArray = New DoubleCollection({2, 2})
                }
                canvas.Children.Add(gridLine)

                ' %25, %50, %75, %100 Etiketleri
                If showLabels Then
                    Dim lbl As New TextBlock With {
                        .Text = $"{i * 25}{unit}",
                        .Foreground = New SolidColorBrush(Color.FromArgb(120, 150, 150, 180)), ' Biraz daha belirgin renk
                        .FontSize = 10
                    }
                    Canvas.SetLeft(lbl, 2)
                    Canvas.SetTop(lbl, yPos - 14)
                    canvas.Children.Add(lbl)
                End If
            Next
        End If

        Dim pts(n - 1) As Point
        For i = 0 To n - 1
            Dim x = padLeft + (i / (n - 1.0)) * chartW
            Dim v = Math.Min(Math.Max(values(i), 0), actualMax)
            Dim y = padTop + chartH * (1 - v / actualMax)
            y = Math.Max(padTop, Math.Min(padTop + chartH, y))
            pts(i) = New Point(x, y)
        Next

        Dim fillPts As New PointCollection()
        fillPts.Add(New Point(padLeft, padTop + chartH))
        For Each pt In pts : fillPts.Add(pt) : Next
        fillPts.Add(New Point(padLeft + chartW, padTop + chartH))

        Dim fillPoly As New Polygon With {
            .Points = fillPts,
            .Fill = New LinearGradientBrush(
                Color.FromArgb(70, lineColor.R, lineColor.G, lineColor.B),
                Color.FromArgb(8, lineColor.R, lineColor.G, lineColor.B),
                New Point(0, 0), New Point(0, 1)
            ),
            .StrokeThickness = 0
        }
        canvas.Children.Add(fillPoly)

        Dim linePts As New PointCollection(pts)
        Dim polyline As New Polyline With {
            .Points = linePts,
            .Stroke = New SolidColorBrush(lineColor),
            .StrokeThickness = If(showGrid, 2.0, 1.6),
            .StrokeLineJoin = PenLineJoin.Round,
            .StrokeStartLineCap = PenLineCap.Round,
            .StrokeEndLineCap = PenLineCap.Round
        }
        canvas.Children.Add(polyline)

        If pts.Length > 0 Then
            Dim last = pts(pts.Length - 1)
            Dim outerDot As New Ellipse With {
                .Width = 10, .Height = 10,
                .Fill = New SolidColorBrush(Color.FromArgb(50, lineColor.R, lineColor.G, lineColor.B)),
                .IsHitTestVisible = False
            }
            Canvas.SetLeft(outerDot, last.X - 5)
            Canvas.SetTop(outerDot, last.Y - 5)
            canvas.Children.Add(outerDot)

            Dim innerDot As New Ellipse With {
                .Width = 5, .Height = 5,
                .Fill = New SolidColorBrush(lineColor),
                .IsHitTestVisible = False
            }
            Canvas.SetLeft(innerDot, last.X - 2.5)
            Canvas.SetTop(innerDot, last.Y - 2.5)
            canvas.Children.Add(innerDot)
        End If
    End Sub

    Public Sub DrawUsageBar(canvas As Canvas, pct As Double, barColor As Color)
        canvas.Children.Clear()
        Dim w = canvas.ActualWidth
        Dim h = canvas.ActualHeight
        If w <= 0 OrElse h <= 0 Then Return

        Dim bg As New Rectangle With {
            .Width = w, .Height = h,
            .Fill = New SolidColorBrush(Color.FromArgb(60, 40, 40, 80)),
            .RadiusX = 4, .RadiusY = 4
        }
        canvas.Children.Add(bg)

        Dim fillW = Math.Min(w * (pct / 100.0), w)
        If fillW > 0 Then
            Dim fill As New Rectangle With {
                .Width = fillW, .Height = h,
                .Fill = New SolidColorBrush(barColor),
                .RadiusX = 4, .RadiusY = 4
            }
            canvas.Children.Add(fill)
        End If
    End Sub

    ' YENİ: Çekirdek ızgarası için mini bar çizimi
    Public Sub DrawMiniBar(canvas As Canvas, pct As Double, barColor As Color)
        canvas.Children.Clear()
        Dim w = canvas.ActualWidth
        Dim h = canvas.ActualHeight
        If w <= 0 OrElse h <= 0 Then Return

        Dim bg As New Rectangle With {
            .Width = w, .Height = h,
            .Fill = New SolidColorBrush(Color.FromArgb(40, 40, 40, 80)),
            .RadiusX = 3, .RadiusY = 3
        }
        canvas.Children.Add(bg)

        Dim fillH = Math.Min(h * (pct / 100.0), h)
        If fillH > 0 Then
            Dim fill As New Rectangle With {
                .Width = w,
                .Height = fillH,
                .Fill = New SolidColorBrush(barColor),
                .RadiusX = 3, .RadiusY = 3
            }
            Canvas.SetTop(fill, h - fillH)
            canvas.Children.Add(fill)
        End If
    End Sub

End Module