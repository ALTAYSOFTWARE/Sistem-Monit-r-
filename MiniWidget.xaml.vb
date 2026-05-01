Imports System.Windows
Imports System.Windows.Input
Imports System.Windows.Threading
Imports Sistem_Monitörü.SystemMonitor

' DİKKAT: "Namespace SystemMonitor" satırını MainWindow'daki gibi sildik.

Partial Public Class MiniWidget
    Inherits Window

    Private _hw As HardwareService
    Private _timer As New DispatcherTimer()
    Private _totalRam As Long

    Public Sub Setup(hwService As HardwareService, totalRamMB As Long)
        _hw = hwService
        _totalRam = totalRamMB
        _timer.Interval = TimeSpan.FromSeconds(1)
        AddHandler _timer.Tick, AddressOf OnTick
        _timer.Start()
    End Sub

    Private Sub OnTick(sender As Object, e As EventArgs)
        If _hw Is Nothing Then Return

        ' 1. CPU (Metin formatlama hatası klasik birleştirme ile çözüldü)
        Dim cpuUsage = _hw.GetCpuUsage()
        txtMiniCpu.Text = Math.Round(cpuUsage, 0).ToString() & "%"

        ' 2. RAM
        Dim usedRam = _totalRam - _hw.GetAvailableRamMB()
        Dim ramPct = If(_totalRam > 0, (usedRam / _totalRam) * 100, 0)
        txtMiniRam.Text = Math.Round(ramPct, 0).ToString() & "% (" & Math.Round(usedRam / 1024.0, 1).ToString() & " GB)"

        ' 3. PING
        If DateTime.Now.Second Mod 3 = 0 Then
            Dim p = _hw.GetPing()
            If p >= 0 Then
                txtMiniPing.Text = p.ToString() & " ms"
            Else
                txtMiniPing.Text = "Hata"
            End If
        End If
    End Sub

    Private Sub Window_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Me.DragMove() ' Ekrandan sürüklemeyi sağlar
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As RoutedEventArgs)
        _timer.Stop()
        Application.Current.MainWindow.Show() ' Ana ekranı geri getir
        Me.Close()
    End Sub
End Class