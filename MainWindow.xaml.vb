Imports System.Collections.ObjectModel
Imports System.Linq
Imports System.Runtime.Remoting
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Threading



Partial Public Class MainWindow
        Inherits Window
    Private _allTasks As List(Of ScheduledTaskInfo) = New List(Of ScheduledTaskInfo)()
    Private ReadOnly _hw As New HardwareService()
        Private ReadOnly _timer As New DispatcherTimer()
        Private _alarmConfig As New AlarmConfig()

        Private ReadOnly _cpuQ As New Queue(Of Double)
        Private ReadOnly _ramQ As New Queue(Of Double)
        Private ReadOnly _diskQ As New Queue(Of Double)
        Private ReadOnly _netQ As New Queue(Of Double)
        Private Const HISTORY_LEN As Integer = 60

        Private _totalRam As Long = 0
        Private _activeView As String = "dashboard"
        Private _alarmSnoozeUntil As DateTime = DateTime.MinValue
        Private ReadOnly _alarmHistory As New ObservableCollection(Of AlarmHistoryEntry)()
        Private _allProcesses As List(Of SysProcessInfo) = New List(Of SysProcessInfo)()
        Private _allServices As List(Of ServiceInfo) = New List(Of ServiceInfo)()
        Private _allPrograms As List(Of InstalledProgram) = New List(Of InstalledProgram)()
        Private _allFirewallRules As List(Of FirewallRule) = New List(Of FirewallRule)()

        Private _tickCount As Integer = 0
        Private _isStressTesting As Boolean = False
        Private _lastCpuPct As Double = 0
        Private _lastRamPct As Double = 0
        Private _lastUsedRamMB As Long = 0

        Private _trayIcon As System.Windows.Forms.NotifyIcon
        Private ReadOnly _configPath As String = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SysMon_Ayarlar.txt")

        Private _coreCanvases As New List(Of Canvas)()
        Private _coreLabels As New List(Of TextBlock)()

        ' YENİ: GEÇMİŞ DEĞİŞKENLERİ
        Private ReadOnly _cpuHistory60min As New Queue(Of Double)()
        Private _cpuMinuteSum As Double = 0
        Private _cpuMinuteCount As Integer = 0

        Public Sub New()
            InitializeComponent()
            InitQueues()
            InitTrayIcon()
            InitApp()
            AddHandler Me.KeyDown, AddressOf MainWindow_KeyDown
        End Sub

        Private Sub MainWindow_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyboardDevice.Modifiers = ModifierKeys.Control Then
                Select Case e.Key
                    Case Key.D1 : NavToView("dashboard")
                    Case Key.D2 : NavToView("cpu")
                    Case Key.D3 : NavToView("ram")
                    Case Key.D4 : NavToView("disk")
                    Case Key.D5 : NavToView("network")
                    Case Key.D6 : NavToView("firewall")
                    Case Key.D7 : NavToView("processes")
                    Case Key.D8 : NavToView("services")
                    Case Key.D9 : NavToView("startup")
                    Case Key.H : NavToView("history")
                    Case Key.D0 : NavToView("alarms")
                    Case Key.W : BtnMiniWidget_Click(Nothing, Nothing)
                End Select
            End If
        End Sub

        Private Sub InitQueues()
            For i = 1 To HISTORY_LEN
                _cpuQ.Enqueue(0) : _ramQ.Enqueue(0) : _diskQ.Enqueue(0) : _netQ.Enqueue(0)
                _cpuHistory60min.Enqueue(0) ' Grafiğin sağdan dolması için sıfırlarla başlatılır
            Next
        End Sub

        Private Sub InitTrayIcon()
            _trayIcon = New System.Windows.Forms.NotifyIcon()
            _trayIcon.Icon = System.Drawing.Icon.ExtractAssociatedIcon(Reflection.Assembly.GetExecutingAssembly().Location)
            _trayIcon.Text = "Sistem Monitörü"
            _trayIcon.Visible = False
            AddHandler _trayIcon.DoubleClick, Sub()
                                                  Me.Show()
                                                  Me.WindowState = WindowState.Normal
                                                  _trayIcon.Visible = False
                                              End Sub
        End Sub

        Private Sub LoadSettings()
            Try
                If System.IO.File.Exists(_configPath) Then
                    Dim cfg As New Dictionary(Of String, String)
                    For Each line In System.IO.File.ReadAllLines(_configPath)
                        Dim idx = line.IndexOf("="c)
                        If idx > 0 Then cfg(line.Substring(0, idx).Trim()) = line.Substring(idx + 1).Trim()
                    Next

                    Dim getBool = Function(k As String, def As Boolean) As Boolean
                                      Dim v As String = Nothing
                                      If cfg.TryGetValue(k, v) Then
                                          Dim r As Boolean
                                          If Boolean.TryParse(v, r) Then Return r
                                      End If
                                      Return def
                                  End Function

                    Dim getInt = Function(k As String, def As Integer) As Integer
                                     Dim v As String = Nothing
                                     If cfg.TryGetValue(k, v) Then
                                         Dim r As Integer
                                         If Integer.TryParse(v, r) Then Return r
                                     End If
                                     Return def
                                 End Function

                    Dim getStr = Function(k As String, def As String) As String
                                     Dim v As String = Nothing
                                     If cfg.TryGetValue(k, v) Then Return v
                                     Return def
                                 End Function

                    _alarmConfig.CpuEnabled = getBool("CpuEnabled", True)
                    _alarmConfig.CpuThreshold = getInt("CpuThreshold", 90)
                    _alarmConfig.RamEnabled = getBool("RamEnabled", True)
                    _alarmConfig.RamThreshold = getInt("RamThreshold", 85)
                    _alarmConfig.DiskEnabled = getBool("DiskEnabled", True)
                    _alarmConfig.DiskThreshold = getInt("DiskThreshold", 90)
                    _alarmConfig.ThemeColor = getStr("ThemeColor", "#00D4FF")
                    _alarmConfig.RunAtStartup = getBool("RunAtStartup", False)
                    _alarmConfig.MinimizeToTray = getBool("MinimizeToTray", True)
                    _alarmConfig.AlarmSoundEnabled = getBool("AlarmSoundEnabled", True)
                    _alarmConfig.SnoozeDurationSec = getInt("SnoozeDurationSec", 15)
                End If
            Catch : End Try

            chkCpuAlarm.IsChecked = _alarmConfig.CpuEnabled : txtCpuThreshold.Text = _alarmConfig.CpuThreshold.ToString() : sldrCpu.Value = _alarmConfig.CpuThreshold
            chkRamAlarm.IsChecked = _alarmConfig.RamEnabled : txtRamThreshold.Text = _alarmConfig.RamThreshold.ToString() : sldrRam.Value = _alarmConfig.RamThreshold
            chkDiskAlarm.IsChecked = _alarmConfig.DiskEnabled : txtDiskThreshold.Text = _alarmConfig.DiskThreshold.ToString()
            chkStartup.IsChecked = _alarmConfig.RunAtStartup : chkTray.IsChecked = _alarmConfig.MinimizeToTray
            chkAlarmSound.IsChecked = _alarmConfig.AlarmSoundEnabled : txtSnoozeSec.Text = _alarmConfig.SnoozeDurationSec.ToString()

            Try
                Dim newColor = CType(ColorConverter.ConvertFromString(_alarmConfig.ThemeColor), Color)
                Me.Resources("AccentColor") = New SolidColorBrush(newColor) : Me.Resources("CpuColor") = New SolidColorBrush(newColor)
            Catch : End Try
        End Sub

        Private Sub SaveSettings()
            Try
                Dim lines As New List(Of String) From {
                    $"CpuEnabled={_alarmConfig.CpuEnabled}", $"CpuThreshold={_alarmConfig.CpuThreshold}",
                    $"RamEnabled={_alarmConfig.RamEnabled}", $"RamThreshold={_alarmConfig.RamThreshold}",
                    $"DiskEnabled={_alarmConfig.DiskEnabled}", $"DiskThreshold={_alarmConfig.DiskThreshold}",
                    $"ThemeColor={_alarmConfig.ThemeColor}", $"RunAtStartup={_alarmConfig.RunAtStartup}",
                    $"MinimizeToTray={_alarmConfig.MinimizeToTray}", $"AlarmSoundEnabled={_alarmConfig.AlarmSoundEnabled}",
                    $"SnoozeDurationSec={_alarmConfig.SnoozeDurationSec}"
                }
                System.IO.File.WriteAllLines(_configPath, lines)
            Catch : End Try
        End Sub

        Private Async Sub InitApp()
            LoadSettings()
            _hw.Initialize()
            _totalRam = _hw.GetTotalRamMB()

            txtPcName.Text = Environment.MachineName
            txtSysInfo.Text = _hw.GetSystemInfoText()
            txtDeepHw.Text = $"Anakart: {_hw.DeepInfo.Motherboard}{vbCrLf}BIOS: {_hw.DeepInfo.BIOS}{vbCrLf}RAM Hızı: {_hw.DeepInfo.RamSpeed}"
            txtKisayollar.Text = "Ctrl+1..0: Sekmeler  |  Ctrl+H: Geçmiş  |  Ctrl+W: Mini Ekran"

            If lvAlarmHistory IsNot Nothing Then lvAlarmHistory.ItemsSource = _alarmHistory

            _timer.Interval = TimeSpan.FromSeconds(1)
            AddHandler _timer.Tick, AddressOf OnTick
            _timer.Start()

            RefreshDiskDrives()
            BuildCoreGrid()

            Dim myIp = Await _hw.GetPublicIpAsync()
            txtPublicIp.Text = myIp
        End Sub

    Private Sub OnTick(sender As Object, e As EventArgs)
        _tickCount += 1

        Dim cpu = CDbl(_hw.GetCpuUsage())
        Dim availRam = _hw.GetAvailableRamMB()
        Dim usedRam = _totalRam - availRam
        Dim ramPct = If(_totalRam > 0, usedRam / _totalRam * 100.0, 0)
        Dim diskRead = _hw.GetDiskReadMBps() : Dim diskWrite = _hw.GetDiskWriteMBps() : Dim diskTotal = diskRead + diskWrite
        Dim netSent = _hw.GetNetSentKBps() : Dim netRecv = _hw.GetNetRecvKBps() : Dim netTotal = netSent + netRecv

        _lastCpuPct = cpu : _lastRamPct = ramPct : _lastUsedRamMB = usedRam
        EnqueueVal(_cpuQ, cpu) : EnqueueVal(_ramQ, ramPct) : EnqueueVal(_diskQ, diskTotal) : EnqueueVal(_netQ, netTotal)

        ' GEÇMİŞ ORTALAMASINI KAYDET
        _cpuMinuteSum += cpu
        _cpuMinuteCount += 1
        If _tickCount Mod 60 = 0 Then
            Dim avgCpu = _cpuMinuteSum / _cpuMinuteCount
            _cpuHistory60min.Enqueue(avgCpu)
            If _cpuHistory60min.Count > 60 Then _cpuHistory60min.Dequeue()
            _cpuMinuteSum = 0
            _cpuMinuteCount = 0
            If _activeView = "history" Then UpdateHistoryPage()
        End If

        txtDateTime.Text = DateTime.Now.ToString("dd.MM.yyyy  HH:mm:ss")
        Dim up = _hw.GetUptime()
        txtUptime.Text = $"⏱  {up.Days}g {up.Hours:D2}:{up.Minutes:D2}:{up.Seconds:D2}"

        UpdateDashboard(cpu, ramPct, usedRam, availRam, diskTotal, netTotal)

        Select Case _activeView
            Case "cpu" : UpdateCpuPage(cpu)
            Case "ram" : UpdateRamPage(ramPct, usedRam, availRam)
            Case "disk" : UpdateDiskPage(diskRead, diskWrite, diskTotal)
            Case "network" : UpdateNetworkPage(netSent, netRecv, netTotal)
        End Select

        ' YENİ EKLENEN: GÜVENLİK DURUMU GÜNCELLEMESİ (Her 5 saniyede bir)
        ' Programın donmaması için Asenkron (Arka Planda) tarama yapar.
        If _tickCount Mod 5 = 0 AndAlso _activeView = "dashboard" Then
            If icSecurityStatus IsNot Nothing Then
                Task.Run(Function() _hw.GetSecurityStatus()).ContinueWith(Sub(t)
                                                                              Dispatcher.Invoke(Sub() icSecurityStatus.ItemsSource = t.Result)
                                                                          End Sub)
            End If
        End If

        If _tickCount Mod 3 = 0 Then
            _allProcesses = _hw.GetTopProcesses(50)
            If lvDashProcesses IsNot Nothing Then lvDashProcesses.ItemsSource = _hw.GetTopProcesses(10)
            If _activeView = "processes" Then RefreshProcessList()
            If _activeView = "network" Then
                Dim ping = _hw.GetPing()
                txtPing.Text = If(ping >= 0, $"{ping} ms", "Bağlantı Hatası")
            End If
        End If

        If _activeView = "cpu" Then UpdateCoreGrid()

        If _tickCount Mod 10 = 0 OrElse _tickCount = 1 Then
            Dim bat = _hw.GetBatteryInfo()
            If bat.HasBattery Then
                txtBatteryInfo.Visibility = Visibility.Visible
                txtBatteryInfo.Text = $"🔋 Pil: %{bat.ChargePercent} ({bat.Status})"
            End If
        End If

        CheckAlarms(cpu, ramPct)
    End Sub
    ' ── YENİ: PAYLAŞILAN KLASÖRLERİ YÜKLE ──
    Private Async Sub LoadSharedFolders()
        If btnRefreshShares IsNot Nothing Then btnRefreshShares.IsEnabled = False
        lvSharedFolders.ItemsSource = Await Task.Run(Function() _hw.GetSharedFolders())
        If btnRefreshShares IsNot Nothing Then btnRefreshShares.IsEnabled = True
    End Sub

    Private Sub BtnRefreshShares_Click(sender As Object, e As RoutedEventArgs)
        LoadSharedFolders()
    End Sub
    Private Sub BtnMiniWidget_Click(sender As Object, e As RoutedEventArgs)
            Dim widget As New MiniWidget()
            widget.Setup(_hw, _totalRam)
            widget.Show()
            Me.Hide()
        End Sub

        Private Function GetAccentColor() As Color
            Return CType(Me.Resources("AccentColor"), SolidColorBrush).Color
        End Function

        Private Sub UpdateDashboard(cpu As Double, ramPct As Double, usedRam As Long, availRam As Long, diskTotal As Double, netTotal As Double)
            Dim accent = GetAccentColor()
            txtCpuPct.Text = $"{cpu:F0}%" : txtRamPct.Text = $"{ramPct:F0}%"
            txtDiskIO.Text = $"{diskTotal:F1} MB/s" : txtNetIO.Text = FormatNet(netTotal)
            txtCpuSubInfo.Text = $"Ort: {_cpuQ.Average():F1}% | Maks: {_cpuQ.Max():F1}%"
            txtRamSubInfo.Text = $"{usedRam / 1024.0:F1} GB / {_totalRam / 1024.0:F1} GB"

            ChartHelper.DrawUsageBar(canvasCpuBar, cpu, accent)
            ChartHelper.DrawUsageBar(canvasRamBarCard, ramPct, Color.FromRgb(255, 107, 107))

            DrawChart(chartCpuCard, _cpuQ, 100, accent) : DrawChart(chartRamCard, _ramQ, 100, Color.FromRgb(255, 107, 107))
            DrawChart(chartDiskCard, _diskQ, 0, Color.FromRgb(255, 217, 61)) : DrawChart(chartNetCard, _netQ, 0, Color.FromRgb(107, 203, 119))
            DrawChart(chartCpuBig, _cpuQ, 100, accent, True, True) : DrawChart(chartRamBig, _ramQ, 100, Color.FromRgb(255, 107, 107), True, True)
        End Sub

        Private Sub UpdateHistoryPage()
            If _cpuHistory60min.Count > 0 Then
                txtHistoryMax.Text = $"Zirve: %{_cpuHistory60min.Max():F1}"
                DrawChart(chartHistoryCpu, _cpuHistory60min, 100, GetAccentColor(), True, True, "%")
            End If
        End Sub

        Private Sub UpdateCpuPage(cpu As Double)
            txtCpuCurrent.Text = $"{cpu:F1}%" : txtCpuAvg.Text = $"{_cpuQ.Average():F1}%" : txtCpuMax.Text = $"{_cpuQ.Max():F1}%" : txtCpuBadge.Text = $"{cpu:F0}%"
            Dim temp = _hw.GetCpuTemperature()
            txtCpuTemp.Text = If(temp > 0, $"{temp:F1} °C", "N/A")
            Dim gpuLoad = _hw.GetGpuLoad()
            txtGpuLoad.Text = If(gpuLoad >= 0, $"{gpuLoad:F0}%", "N/A")
            DrawChart(chartCpuFull, _cpuQ, 100, GetAccentColor(), True, True, "%")
        End Sub

        Private Sub BuildCoreGrid()
            Dim coreCount = Environment.ProcessorCount
            _coreCanvases.Clear() : _coreLabels.Clear()
            If pnlCoreGrid Is Nothing Then Return
            pnlCoreGrid.Children.Clear()

            For i = 0 To coreCount - 1
                Dim sp As New StackPanel With {.Width = 56, .Margin = New Thickness(4, 0, 4, 0), .HorizontalAlignment = HorizontalAlignment.Center}
                Dim bar As New Canvas With {.Height = 50}
                Dim lbl As New TextBlock With {.Text = $"Ç{i}", .Foreground = New SolidColorBrush(Color.FromArgb(160, 150, 150, 200)), .FontSize = 10, .HorizontalAlignment = HorizontalAlignment.Center, .Margin = New Thickness(0, 3, 0, 0)}
                Dim pctLbl As New TextBlock With {.Text = "0%", .Foreground = New SolidColorBrush(Color.FromRgb(200, 200, 230)), .FontSize = 10, .FontWeight = FontWeights.SemiBold, .HorizontalAlignment = HorizontalAlignment.Center}
                sp.Children.Add(bar) : sp.Children.Add(pctLbl) : sp.Children.Add(lbl)
                pnlCoreGrid.Children.Add(sp)
                _coreCanvases.Add(bar) : _coreLabels.Add(pctLbl)
            Next
        End Sub

        Private Sub UpdateCoreGrid()
            If _coreCanvases.Count = 0 Then Return
            Dim accent = GetAccentColor()
            Dim cores = _hw.GetCpuCoreUsages()
            For i = 0 To Math.Min(cores.Count - 1, _coreCanvases.Count - 1)
                Dim pct = cores(i).UsagePct
                Dim barColor = If(pct > 85, Color.FromRgb(255, 100, 100), If(pct > 60, Color.FromRgb(255, 200, 60), accent))
                ChartHelper.DrawMiniBar(_coreCanvases(i), pct, barColor)
                _coreLabels(i).Text = $"{pct:F0}%"
            Next
        End Sub

        Private Sub UpdateRamPage(ramPct As Double, usedRam As Long, availRam As Long)
            txtRamBadge.Text = $"{ramPct:F0}%" : txtRamUsed.Text = $"{usedRam / 1024.0:F2} GB" : txtRamFree.Text = $"{availRam / 1024.0:F2} GB" : txtRamTotal.Text = $"{_totalRam / 1024.0:F1} GB"
            txtRamBarLabel.Text = $"Kullanılan: {usedRam / 1024.0:F1} GB  |  Boş: {availRam / 1024.0:F1} GB  | Toplam: {_totalRam / 1024.0:F1} GB  ({ramPct:F1}%)"
            DrawChart(chartRamFull, _ramQ, 100, Color.FromRgb(255, 107, 107), True, True, "%")
            ChartHelper.DrawUsageBar(canvasRamBar, ramPct, If(ramPct > 85, Color.FromRgb(255, 100, 100), Color.FromRgb(255, 107, 107)))
        End Sub

        Private Sub UpdateDiskPage(diskRead As Double, diskWrite As Double, diskTotal As Double)
            txtDiskRead.Text = $"{diskRead:F2} MB/s" : txtDiskWrite.Text = $"{diskWrite:F2} MB/s"
            DrawChart(chartDiskFull, _diskQ, 0, Color.FromRgb(255, 217, 61), True, True, " MB/s")
        End Sub

        Private Sub UpdateNetworkPage(netSent As Double, netRecv As Double, netTotal As Double)
            txtNetSent.Text = FormatNet(netSent) : txtNetRecv.Text = FormatNet(netRecv)
            DrawChart(chartNetFull, _netQ, 0, Color.FromRgb(107, 203, 119), True, True, " KB/s")
        End Sub

        Private Sub RefreshProcessList()
            Dim filter = txtProcessSearch.Text.Trim().ToLower()
            Dim filtered = If(String.IsNullOrEmpty(filter), _allProcesses, _allProcesses.Where(Function(p) p.Name.ToLower().Contains(filter)).ToList())
            lvAllProcesses.ItemsSource = filtered : txtProcessCount.Text = $"{filtered.Count} süreç"
        End Sub

        Private Sub BtnRefreshProcesses_Click(sender As Object, e As RoutedEventArgs)
            _allProcesses = _hw.GetTopProcesses(50) : RefreshProcessList()
        End Sub
        Private Sub TxtProcessSearch_Changed(sender As Object, e As TextChangedEventArgs)
            RefreshProcessList()
        End Sub

        Private Sub SetPriorityHigh_Click(sender As Object, e As RoutedEventArgs)
            Dim sel = TryCast(lvAllProcesses.SelectedItem, SysProcessInfo)
            If sel IsNot Nothing Then
                Try
                    Process.GetProcessById(sel.PID).PriorityClass = ProcessPriorityClass.High
                    MessageBox.Show($"'{sel.Name}' YÜKSEK önceliğe alındı.", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information)
                Catch ex As Exception
                    MessageBox.Show($"Erişim reddedildi: {ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error)
                End Try
            End If
        End Sub

        Private Sub SetPriorityNormal_Click(sender As Object, e As RoutedEventArgs)
            Dim sel = TryCast(lvAllProcesses.SelectedItem, SysProcessInfo)
            If sel IsNot Nothing Then
                Try
                    Process.GetProcessById(sel.PID).PriorityClass = ProcessPriorityClass.Normal
                    MessageBox.Show($"'{sel.Name}' NORMAL önceliğe alındı.", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information)
                Catch ex As Exception
                    MessageBox.Show($"Erişim reddedildi. {ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error)
                End Try
            End If
        End Sub

        Private Sub KillProcess_Click(sender As Object, e As RoutedEventArgs)
            Dim sel = TryCast(lvAllProcesses.SelectedItem, SysProcessInfo)
            If sel IsNot Nothing AndAlso MessageBox.Show($"'{sel.Name}' sürecini sonlandır?", "Kapat", MessageBoxButton.YesNo, MessageBoxImage.Warning) = MessageBoxResult.Yes Then
                Try
                    Process.GetProcessById(sel.PID).Kill() : _allProcesses = _hw.GetTopProcesses(50) : RefreshProcessList()
                Catch ex As Exception
                    MessageBox.Show($"Kapatılamadı: {ex.Message}")
                End Try
            End If
        End Sub

        Private Async Sub CreateDump_Click(sender As Object, e As RoutedEventArgs)
            Dim sel = TryCast(lvAllProcesses.SelectedItem, SysProcessInfo)
            If sel IsNot Nothing Then
                Dim dlg As New Microsoft.Win32.SaveFileDialog() With {
                    .FileName = $"{sel.Name}_{sel.PID}.dmp", .Filter = "Dump Dosyası (*.dmp)|*.dmp", .Title = "Bellek Dökümünü Kaydet"
                }
                If dlg.ShowDialog() = True Then
                    Mouse.OverrideCursor = Cursors.Wait
                    Dim success = Await Task.Run(Function() _hw.CreateMiniDump(sel.PID, dlg.FileName))
                    Mouse.OverrideCursor = Nothing
                    If success Then MessageBox.Show($"Bellek dökümü oluşturuldu!" & vbCrLf & dlg.FileName, "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information) Else MessageBox.Show($"Döküm alınamadı! Yönetici izni gerekebilir.", "Hata", MessageBoxButton.OK, MessageBoxImage.Error)
                End If
            End If
        End Sub

        Private Async Sub BtnToolsTakeDump_Click(sender As Object, e As RoutedEventArgs)
            Dim pid As Integer
            If Not Integer.TryParse(txtDumpPid.Text.Trim(), pid) Then MessageBox.Show("Geçerli bir PID girin.", "Geçersiz", MessageBoxButton.OK, MessageBoxImage.Warning) : Return
            Dim procName As String = "Process"
            Try
                procName = Process.GetProcessById(pid).ProcessName
            Catch ex As Exception
                MessageBox.Show("Aktif süreç bulunamadı.", "Hata", MessageBoxButton.OK, MessageBoxImage.Error) : Return
            End Try
            Dim dlg As New Microsoft.Win32.SaveFileDialog() With {
                .FileName = $"{procName}_{pid}.dmp", .Filter = "Dump Dosyası (*.dmp)|*.dmp", .Title = "Bellek Dökümünü Kaydet"
            }
            If dlg.ShowDialog() = True Then
                Mouse.OverrideCursor = Cursors.Wait
                Dim success = Await Task.Run(Function() _hw.CreateMiniDump(pid, dlg.FileName))
                Mouse.OverrideCursor = Nothing
                If success Then MessageBox.Show($"Bellek dökümü oluşturuldu!", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information) Else MessageBox.Show($"Döküm alınamadı!", "Hata", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        End Sub

        Private Async Sub BtnCleanMemory_Click(sender As Object, e As RoutedEventArgs)
            Dim btn = CType(sender, Button)
            btn.Content = "⏳ Temizleniyor..." : btn.IsEnabled = False
            Dim count = Await Task.Run(Function() _hw.CleanAllMemory())
            btn.Content = "🧹 RAM Temizle" : btn.IsEnabled = True
            MessageBox.Show($"{count} sürecin bellek çalışma seti temizlendi.", "RAM Temizlendi", MessageBoxButton.OK, MessageBoxImage.Information)
        End Sub

        Private Sub RefreshDiskDrives()
            Dim drives = _hw.GetDiskDrives()
            pnlDiskDrives.Children.Clear()
            For Each d In drives
                Dim sp As New StackPanel With {.Margin = New Thickness(0, 0, 0, 16)}
                Dim header As New Grid
                header.ColumnDefinitions.Add(New ColumnDefinition With {.Width = New GridLength(1, GridUnitType.Star)})
                header.ColumnDefinitions.Add(New ColumnDefinition With {.Width = GridLength.Auto})
                Dim lblName As New TextBlock With {.Text = $"💿  {d.DisplayName}", .Foreground = New SolidColorBrush(Color.FromRgb(200, 200, 220)), .FontSize = 13, .FontWeight = FontWeights.SemiBold}
                Dim lblInfo As New TextBlock With {.Text = $"{d.UsedGB:F1} GB / {d.TotalGB:F1} GB  ({d.UsedPct:F1}% dolu)", .Foreground = New SolidColorBrush(Color.FromArgb(160, 180, 180, 200)), .FontSize = 11.5, .HorizontalAlignment = HorizontalAlignment.Right}
                Grid.SetColumn(lblInfo, 1) : header.Children.Add(lblName) : header.Children.Add(lblInfo) : sp.Children.Add(header)
                Dim barCanvas As New Canvas With {.Height = 14, .Margin = New Thickness(0, 6, 0, 0)} : sp.Children.Add(barCanvas)
                Dim barColor = If(d.UsedPct > 85, Color.FromRgb(255, 80, 80), If(d.UsedPct > 70, Color.FromRgb(255, 170, 50), Color.FromRgb(255, 217, 61)))
                ChartHelper.DrawUsageBar(barCanvas, d.UsedPct, barColor)
                sp.Children.Add(New TextBlock With {.Text = $"Boş: {d.FreeGB:F1} GB", .Foreground = New SolidColorBrush(Color.FromArgb(120, 150, 150, 180)), .FontSize = 11, .Margin = New Thickness(0, 4, 0, 0)})
                pnlDiskDrives.Children.Add(sp)
            Next
            If lvDiskHealth IsNot Nothing Then lvDiskHealth.ItemsSource = _hw.GetDiskHealth()
        End Sub

        Private Async Sub BtnAnalyzeDisk_Click(sender As Object, e As RoutedEventArgs)
            Dim btn = CType(sender, Button)
            btn.Content = "⏳ Taranıyor..." : btn.IsEnabled = False
            lvDiskSpace.ItemsSource = Nothing
            lvDiskSpace.ItemsSource = Await Task.Run(Function() _hw.GetDirectorySizes("C:\"))
            btn.Content = "🔍 Tara" : btn.IsEnabled = True
        End Sub

        Private Async Sub BtnRefreshNetstat_Click(sender As Object, e As RoutedEventArgs)
            Dim btn = CType(sender, Button)
            btn.Content = "⏳ Taranıyor..." : btn.IsEnabled = False
            lvNetConnections.ItemsSource = Await Task.Run(Function() _hw.GetActiveNetworkConnections())
            btn.Content = "🔄 Tarama Yap" : btn.IsEnabled = True
        End Sub

        Private Async Sub LoadNetworkAdapters()
            If lvNetAdapters.ItemsSource IsNot Nothing Then Return
            lvNetAdapters.ItemsSource = Await Task.Run(Function() _hw.GetNetworkAdapters())
        End Sub

        Private Async Sub BtnRefreshAdapters_Click(sender As Object, e As RoutedEventArgs)
            Dim btn = TryCast(sender, Button)
            If btn IsNot Nothing Then btn.Content = "⏳ Taranıyor..." : btn.IsEnabled = False
            lvNetAdapters.ItemsSource = Await Task.Run(Function() _hw.GetNetworkAdapters())
            If btn IsNot Nothing Then btn.Content = "🔄 Yenile" : btn.IsEnabled = True
        End Sub

        Private Async Sub LoadFirewallRules()
            btnRefreshFirewall.IsEnabled = False
            btnRefreshFirewall.Content = "⏳ Yükleniyor..."
            _allFirewallRules = Await Task.Run(Function() _hw.GetFirewallRules())
            RefreshFirewallList()
            btnRefreshFirewall.IsEnabled = True
            btnRefreshFirewall.Content = "🔄 Yenile"
        End Sub
    ' ── YENİ: ZAMANLI GÖREVLER ──
    Private Async Sub LoadTasks()
        If btnRefreshTasks IsNot Nothing Then btnRefreshTasks.IsEnabled = False : btnRefreshTasks.Content = "⏳ Taranıyor..."
        _allTasks = Await Task.Run(Function() _hw.GetScheduledTasks())
        RefreshTaskList()
        If btnRefreshTasks IsNot Nothing Then btnRefreshTasks.IsEnabled = True : btnRefreshTasks.Content = "🔄 Yenile"
    End Sub

    Private Sub RefreshTaskList()
        Dim filter = txtTaskSearch.Text.Trim().ToLower()
        Dim filtered = If(String.IsNullOrEmpty(filter), _allTasks, _allTasks.Where(Function(t) t.TaskName.ToLower().Contains(filter) OrElse t.Command.ToLower().Contains(filter)).ToList())
        lvTasks.ItemsSource = filtered : txtTaskCount.Text = $"{filtered.Count} görev"
    End Sub

    Private Sub BtnRefreshTasks_Click(sender As Object, e As RoutedEventArgs)
        LoadTasks()
    End Sub

    Private Sub TxtTaskSearch_Changed(sender As Object, e As TextChangedEventArgs)
        RefreshTaskList()
    End Sub
    Private Sub RefreshFirewallList()
            Dim filter = txtFirewallSearch.Text.Trim().ToLower()
            Dim filtered = If(String.IsNullOrEmpty(filter), _allFirewallRules, _allFirewallRules.Where(Function(r) r.RuleName.ToLower().Contains(filter)).ToList())
            lvFirewallRules.ItemsSource = filtered
            txtFirewallCount.Text = $"{filtered.Count} kural"
        End Sub

        Private Sub BtnRefreshFirewall_Click(sender As Object, e As RoutedEventArgs)
            LoadFirewallRules()
        End Sub

        Private Sub TxtFirewallSearch_Changed(sender As Object, e As TextChangedEventArgs)
            RefreshFirewallList()
        End Sub

        Private Sub LoadStartupApps()
            lvStartupApps.ItemsSource = _hw.GetStartupApps()
        End Sub

        Private Async Sub LoadCrashLogs()
            lvCrashLogs.ItemsSource = Nothing
            lvCrashLogs.ItemsSource = Await Task.Run(Function() _hw.GetCrashLogs())
        End Sub

        Private Async Sub LoadServices()
            btnRefreshServices.IsEnabled = False : btnRefreshServices.Content = "⏳ Yükleniyor..."
            _allServices = Await Task.Run(Function() _hw.GetServices())
            RefreshServiceList()
            btnRefreshServices.IsEnabled = True : btnRefreshServices.Content = "🔄 Yenile"
        End Sub

        Private Sub RefreshServiceList()
            Dim filter = txtServiceSearch.Text.Trim().ToLower()
            Dim filtered = If(String.IsNullOrEmpty(filter), _allServices, _allServices.Where(Function(s) s.DisplayName.ToLower().Contains(filter) OrElse s.ServiceName.ToLower().Contains(filter)).ToList())
            lvServices.ItemsSource = filtered : txtServiceCount.Text = $"{filtered.Count} servis"
        End Sub

        Private Sub BtnRefreshServices_Click(sender As Object, e As RoutedEventArgs)
            LoadServices()
        End Sub
        Private Sub TxtServiceSearch_Changed(sender As Object, e As TextChangedEventArgs)
            RefreshServiceList()
        End Sub

        Private Async Sub BtnStartService_Click(sender As Object, e As RoutedEventArgs)
            Dim sel = TryCast(lvServices.SelectedItem, ServiceInfo)
            If sel Is Nothing Then Return
            Dim btn = CType(sender, Button) : btn.IsEnabled = False
            Dim ok = Await Task.Run(Function() _hw.StartService(sel.ServiceName))
            btn.IsEnabled = True
            If ok Then MessageBox.Show($"'{sel.DisplayName}' başlatıldı.", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information) Else MessageBox.Show($"Başlatılamadı.", "Hata", MessageBoxButton.OK, MessageBoxImage.Warning)
            LoadServices()
        End Sub

        Private Async Sub BtnStopService_Click(sender As Object, e As RoutedEventArgs)
            Dim sel = TryCast(lvServices.SelectedItem, ServiceInfo)
            If sel Is Nothing Then Return
            If MessageBox.Show($"'{sel.DisplayName}' durdurulsun mu?", "Onay", MessageBoxButton.YesNo, MessageBoxImage.Warning) <> MessageBoxResult.Yes Then Return
            Dim btn = CType(sender, Button) : btn.IsEnabled = False
            Dim ok = Await Task.Run(Function() _hw.StopService(sel.ServiceName))
            btn.IsEnabled = True
            If ok Then MessageBox.Show($"'{sel.DisplayName}' durduruldu.", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information) Else MessageBox.Show($"Durdurulamadı.", "Hata", MessageBoxButton.OK, MessageBoxImage.Warning)
            LoadServices()
        End Sub

        Private Async Sub LoadPrograms()
            btnRefreshPrograms.IsEnabled = False : btnRefreshPrograms.Content = "⏳ Taranıyor..."
            _allPrograms = Await Task.Run(Function() _hw.GetInstalledPrograms())
            RefreshProgramList()
            btnRefreshPrograms.IsEnabled = True : btnRefreshPrograms.Content = "🔄 Yenile"
        End Sub

        Private Sub RefreshProgramList()
            Dim filter = txtProgramSearch.Text.Trim().ToLower()
            Dim filtered = If(String.IsNullOrEmpty(filter), _allPrograms, _allPrograms.Where(Function(p) p.Name.ToLower().Contains(filter) OrElse p.Publisher.ToLower().Contains(filter)).ToList())
            lvPrograms.ItemsSource = filtered : txtProgramCount.Text = $"{filtered.Count} uygulama"
        End Sub

        Private Sub BtnRefreshPrograms_Click(sender As Object, e As RoutedEventArgs)
            LoadPrograms()
        End Sub
        Private Sub TxtProgramSearch_Changed(sender As Object, e As TextChangedEventArgs)
            RefreshProgramList()
        End Sub

        Private Async Sub BtnScanJunk_Click(sender As Object, e As RoutedEventArgs)
            btnScanJunk.IsEnabled = False
            btnCleanJunk.IsEnabled = False
            btnScanJunk.Content = "⏳ Taranıyor..."

            Dim report = Await Task.Run(Function() _hw.GetJunkFilesReport())
            lvJunkFiles.ItemsSource = report

            Dim totalBytes As Long = report.Sum(Function(x) x.SizeBytes)
            If totalBytes > 1048576 Then
                txtJunkTotal.Text = $"{(totalBytes / 1048576.0):F1} MB Çöp Bulundu"
            Else
                txtJunkTotal.Text = $"{(totalBytes / 1024.0):F1} KB Çöp Bulundu"
            End If

            btnScanJunk.Content = "🔍 Analiz Et"
            btnScanJunk.IsEnabled = True
            btnCleanJunk.IsEnabled = (totalBytes > 0)
        End Sub

        Private Async Sub BtnCleanJunk_Click(sender As Object, e As RoutedEventArgs)
            If MessageBox.Show("Geri Dönüşüm Kutusu ve Sistem Temp dosyaları silinecektir. Devam?", "Sistem Temizliği", MessageBoxButton.YesNo, MessageBoxImage.Warning) <> MessageBoxResult.Yes Then Return
            btnScanJunk.IsEnabled = False : btnCleanJunk.IsEnabled = False : btnCleanJunk.Content = "⏳ Temizleniyor..."
            Mouse.OverrideCursor = Cursors.Wait
            Dim freed = Await Task.Run(Function() _hw.CleanJunkFiles())
            Mouse.OverrideCursor = Nothing
            btnCleanJunk.Content = "✨ Seçilenleri Temizle"
            Dim freedStr = If(freed > 1048576, $"{(freed / 1048576.0):F1} MB", $"{(freed / 1024.0):F1} KB")
            MessageBox.Show($"Temizlik Tamamlandı!" & vbCrLf & $"{freedStr} çöp dosya temizlendi.", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information)
            BtnScanJunk_Click(Nothing, Nothing)
        End Sub

        Private Async Sub LoadDrivers()
            If btnRefreshDrivers IsNot Nothing Then btnRefreshDrivers.IsEnabled = False : btnRefreshDrivers.Content = "⏳ Taranıyor..."
            Dim drivers = Await Task.Run(Function() _hw.GetDrivers())
            lvDrivers.ItemsSource = drivers
            If btnRefreshDrivers IsNot Nothing Then btnRefreshDrivers.IsEnabled = True : btnRefreshDrivers.Content = "🔄 Yenile"
        End Sub

        Private Sub BtnRefreshDrivers_Click(sender As Object, e As RoutedEventArgs)
            LoadDrivers()
        End Sub

        Private Async Sub LoadUsers()
            If btnRefreshUsers IsNot Nothing Then btnRefreshUsers.IsEnabled = False : btnRefreshUsers.Content = "⏳ Taranıyor..."
            Dim users = Await Task.Run(Function() _hw.GetLoggedInUsers())
            lvUsers.ItemsSource = users
            If btnRefreshUsers IsNot Nothing Then btnRefreshUsers.IsEnabled = True : btnRefreshUsers.Content = "🔄 Yenile"
        End Sub

        Private Sub BtnRefreshUsers_Click(sender As Object, e As RoutedEventArgs)
            LoadUsers()
        End Sub
    ' ── YENİ: SÜRÜCÜ GÜNCELLEME KONTROLÜ ──
    Private Async Sub LoadDriverUpdates()
        If btnRefreshDriverDates IsNot Nothing Then btnRefreshDriverDates.IsEnabled = False
        lvDriverUpdates.ItemsSource = Await Task.Run(Function() _hw.GetDriverUpdateReport())
        If btnRefreshDriverDates IsNot Nothing Then btnRefreshDriverDates.IsEnabled = True
    End Sub
    ' ── YENİ: SİSTEM GERİ YÜKLEME ──
    ' ── GÜNCELLENDİ: SİSTEM GERİ YÜKLEME ──
    Private Async Sub LoadRestorePoints()
        If btnRefreshRestore IsNot Nothing Then btnRefreshRestore.IsEnabled = False

        Dim points = Await Task.Run(Function() _hw.GetRestorePoints())
        lvRestorePoints.ItemsSource = points

        ' Veri yoksa uyar, varsa listeyi göster
        If points.Count = 0 Then
            txtNoRestore.Visibility = Visibility.Visible
            lvRestorePoints.Visibility = Visibility.Collapsed
        Else
            txtNoRestore.Visibility = Visibility.Collapsed
            lvRestorePoints.Visibility = Visibility.Visible
        End If

        If btnRefreshRestore IsNot Nothing Then btnRefreshRestore.IsEnabled = True
    End Sub

    Private Sub BtnRefreshRestore_Click(sender As Object, e As RoutedEventArgs)
        LoadRestorePoints()
    End Sub

    Private Async Sub BtnCreateRestore_Click(sender As Object, e As RoutedEventArgs)
        Dim desc = txtNewRestoreName.Text.Trim()
        If String.IsNullOrWhiteSpace(desc) Then
            MessageBox.Show("Lütfen yedek için bir açıklama girin.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        btnCreateRestore.IsEnabled = False
        btnCreateRestore.Content = "⏳ Oluşturuluyor..."

        Dim success = Await Task.Run(Function() _hw.CreateRestorePoint(desc))

        btnCreateRestore.IsEnabled = True
        btnCreateRestore.Content = "➕ Nokta Oluştur"

        If success Then
            MessageBox.Show("Sistem geri yükleme noktası başarıyla oluşturuldu!", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information)
            LoadRestorePoints()
            txtNewRestoreName.Text = "Manuel Yedek" ' Kutuyu sıfırla
        Else
            MessageBox.Show("Nokta oluşturulamadı!" & vbCrLf & "Lütfen uygulamayı Yönetici Olarak Çalıştırdığınızdan ve Windows Ayarlarından 'Sistem Koruması'nın açık olduğundan emin olun.", "Hata", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub
    Private Sub BtnRefreshDriverDates_Click(sender As Object, e As RoutedEventArgs)
        LoadDriverUpdates()
    End Sub
    Private Async Sub LoadTools()
            Dim plans = Await Task.Run(Function() _hw.GetPowerPlans())
            lvPowerPlans.ItemsSource = plans
            Dim adapters = Await Task.Run(Function() _hw.GetNetworkAdapters())
            lvToolsAdapters.ItemsSource = adapters
            LoadDrivers()
            LoadUsers()
        BtnScanJunk_Click(Nothing, Nothing)
        LoadSharedFolders()
        LoadDriverUpdates()
        LoadRestorePoints()
    End Sub

        Private Async Sub BtnSetPowerPlan_Click(sender As Object, e As RoutedEventArgs)
            Dim sel = TryCast(lvPowerPlans.SelectedItem, PowerPlanInfo)
            If sel Is Nothing Then Return
            Await Task.Run(Sub() _hw.SetPowerPlan(sel.Guid))
            Await Task.Delay(500)
            lvPowerPlans.ItemsSource = Await Task.Run(Function() _hw.GetPowerPlans())
            MessageBox.Show($"'{sel.Name}' güç planı aktif edildi.", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information)
        End Sub

        Private Sub BtnExportReport_Click(sender As Object, e As RoutedEventArgs)
            Dim dlg As New Microsoft.Win32.SaveFileDialog() With {.FileName = $"SysMonitor_Rapor_{DateTime.Now:yyyyMMdd_HHmm}.txt", .Filter = "Metin Dosyası|*.txt|Tüm Dosyalar|*.*", .DefaultExt = "txt"}
            If dlg.ShowDialog() = True Then
                _hw.ExportSystemReport(dlg.FileName, _lastCpuPct, _lastRamPct, _lastUsedRamMB)
                MessageBox.Show($"Rapor kaydedildi:{vbCrLf}{dlg.FileName}", "Rapor Dışa Aktarıldı", MessageBoxButton.OK, MessageBoxImage.Information)
            End If
        End Sub

        Private Sub BtnStressTest_Click(sender As Object, e As RoutedEventArgs)
            If _isStressTesting Then
                _isStressTesting = False : btnStressTest.Content = "🔥 Testi Başlat" : btnStressTest.Background = New SolidColorBrush(Color.FromRgb(255, 64, 96))
            Else
                If MessageBox.Show("Bu işlem CPU'yu %100 yük altına sokacak. Devam?", "Uyarı", MessageBoxButton.YesNo) = MessageBoxResult.Yes Then
                    _isStressTesting = True : btnStressTest.Content = "🛑 Testi Durdur" : btnStressTest.Background = New SolidColorBrush(Color.FromRgb(100, 100, 120))
                    For i = 1 To Environment.ProcessorCount
                        Task.Run(Sub()
                                     While _isStressTesting
                                         Dim x As Double = Math.Sqrt(Math.PI * Math.E)
                                     End While
                                 End Sub)
                    Next
                End If
            End If
        End Sub

        Private Sub CheckAlarms(cpu As Double, ramPct As Double)
            If DateTime.Now < _alarmSnoozeUntil Then Return

            Dim msg As String = Nothing
            Dim tp As String = Nothing

            If _alarmConfig.CpuEnabled AndAlso cpu > _alarmConfig.CpuThreshold Then
                msg = $"CPU kullanımı %{cpu:F0} — eşik: %{_alarmConfig.CpuThreshold}" : tp = "CPU Alarmı"
            ElseIf _alarmConfig.RamEnabled AndAlso ramPct > _alarmConfig.RamThreshold Then
                msg = $"RAM kullanımı %{ramPct:F0} — eşik: %{_alarmConfig.RamThreshold}" : tp = "RAM Alarmı"
            ElseIf _alarmConfig.DiskEnabled Then
                For Each drive In _hw.GetDiskDrives()
                    If drive.UsedPct > _alarmConfig.DiskThreshold Then
                        msg = $"{drive.DisplayName} disk doluluk %{drive.UsedPct:F0} — eşik: %{_alarmConfig.DiskThreshold}" : tp = "Disk Alarmı"
                        Exit For
                    End If
                Next
            End If

            If msg IsNot Nothing Then TriggerAlarm(tp, msg)
        End Sub

        Private Sub TriggerAlarm(type As String, msg As String)
            borderAlarmBanner.Visibility = Visibility.Visible
            txtAlarmMsg.Text = $"⚠️  {msg}"
            _alarmHistory.Insert(0, New AlarmHistoryEntry With {.Time = DateTime.Now.ToString("HH:mm:ss"), .Type = type, .Message = msg})
            If _alarmHistory.Count > 100 Then _alarmHistory.RemoveAt(_alarmHistory.Count - 1)

            Dim snoozeSec = If(_alarmConfig.SnoozeDurationSec > 0, _alarmConfig.SnoozeDurationSec, 15)
            _alarmSnoozeUntil = DateTime.Now.AddSeconds(snoozeSec)

            If _alarmConfig.AlarmSoundEnabled Then
                Try
                    System.Media.SystemSounds.Exclamation.Play()
                Catch
                End Try
            End If

            If _trayIcon IsNot Nothing AndAlso Me.Visibility <> Visibility.Visible Then
                _trayIcon.ShowBalloonTip(3000, type, msg, System.Windows.Forms.ToolTipIcon.Warning)
            End If

            Dim hideTimer As New DispatcherTimer With {.Interval = TimeSpan.FromSeconds(8)}
            AddHandler hideTimer.Tick, Sub(s, ev)
                                           borderAlarmBanner.Visibility = Visibility.Collapsed
                                           hideTimer.Stop()
                                       End Sub
            hideTimer.Start()
        End Sub

        Private Sub BtnDismissAlarm_Click(sender As Object, e As RoutedEventArgs)
            borderAlarmBanner.Visibility = Visibility.Collapsed : _alarmSnoozeUntil = DateTime.Now.AddSeconds(30)
        End Sub

        Private Sub BtnSaveAlarms_Click(sender As Object, e As RoutedEventArgs)
            _alarmConfig.CpuEnabled = chkCpuAlarm.IsChecked = True
            _alarmConfig.RamEnabled = chkRamAlarm.IsChecked = True
            _alarmConfig.DiskEnabled = chkDiskAlarm.IsChecked = True
            _alarmConfig.AlarmSoundEnabled = chkAlarmSound.IsChecked = True

            Dim v As Integer
            If Integer.TryParse(txtCpuThreshold.Text, v) Then _alarmConfig.CpuThreshold = Math.Max(10, Math.Min(100, v))
            If Integer.TryParse(txtRamThreshold.Text, v) Then _alarmConfig.RamThreshold = Math.Max(10, Math.Min(100, v))
            If Integer.TryParse(txtDiskThreshold.Text, v) Then _alarmConfig.DiskThreshold = Math.Max(10, Math.Min(100, v))
            If Integer.TryParse(txtSnoozeSec.Text, v) Then _alarmConfig.SnoozeDurationSec = Math.Max(5, Math.Min(3600, v))

            _alarmConfig.RunAtStartup = chkStartup.IsChecked = True
            _alarmConfig.MinimizeToTray = chkTray.IsChecked = True

            _hw.SetRunAtStartup(_alarmConfig.RunAtStartup)
            SaveSettings()
            MessageBox.Show("Ayarlar kaydedildi.", "Kaydedildi", MessageBoxButton.OK, MessageBoxImage.Information)
        End Sub

        Private Sub BtnClearAlarmHistory_Click(sender As Object, e As RoutedEventArgs)
            _alarmHistory.Clear()
        End Sub

        Private Sub SldrCpu_Changed(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double))
            If txtCpuThreshold IsNot Nothing Then txtCpuThreshold.Text = CInt(sldrCpu.Value).ToString()
        End Sub
        Private Sub SldrRam_Changed(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double))
            If txtRamThreshold IsNot Nothing Then txtRamThreshold.Text = CInt(sldrRam.Value).ToString()
        End Sub

        Private Sub BtnTheme_Click(sender As Object, e As RoutedEventArgs)
            Dim btn = TryCast(sender, Button)
            If btn Is Nothing Then Return
            Dim colorHex = btn.Tag.ToString()
            Dim newColor = CType(ColorConverter.ConvertFromString(colorHex), Color)
            Me.Resources("AccentColor") = New SolidColorBrush(newColor)
            Me.Resources("CpuColor") = New SolidColorBrush(newColor)
            _alarmConfig.ThemeColor = colorHex
            SaveSettings()
        End Sub

        Private Sub NavBtn_Click(sender As Object, e As RoutedEventArgs)
            NavToView(CType(sender, Button).Tag.ToString())
        End Sub

        Private Sub NavToView(tag As String)
            _activeView = tag

            For Each child In contentGrid.Children.OfType(Of UIElement)()
                child.Visibility = Visibility.Collapsed
            Next

            txtPageTitle.Text = "📊 Dashboard"
            Select Case tag
                Case "dashboard" : viewDashboard.Visibility = Visibility.Visible : txtPageTitle.Text = "📊 Dashboard"
                Case "cpu" : viewCpu.Visibility = Visibility.Visible : txtPageTitle.Text = "🔲 İşlemci (CPU)"
                Case "ram" : viewRam.Visibility = Visibility.Visible : txtPageTitle.Text = "💾 Bellek (RAM)"
                Case "disk" : viewDisk.Visibility = Visibility.Visible : txtPageTitle.Text = "💿 Disk" : RefreshDiskDrives()
                Case "network" : viewNetwork.Visibility = Visibility.Visible : txtPageTitle.Text = "🌐 Ağ İzleyici" : LoadNetworkAdapters()
                Case "firewall" : viewFirewall.Visibility = Visibility.Visible : txtPageTitle.Text = "🛡️ Güvenlik Duvarı (Firewall)" : If _allFirewallRules.Count = 0 Then LoadFirewallRules()
                Case "processes" : viewProcesses.Visibility = Visibility.Visible : txtPageTitle.Text = "⚙️ Süreçler" : _allProcesses = _hw.GetTopProcesses(50) : RefreshProcessList()
                Case "services" : viewServices.Visibility = Visibility.Visible : txtPageTitle.Text = "🔧 Windows Servisleri" : If _allServices.Count = 0 Then LoadServices()
                Case "startup" : viewStartup.Visibility = Visibility.Visible : txtPageTitle.Text = "🚀 Başlangıç" : LoadStartupApps()
                Case "programs" : viewPrograms.Visibility = Visibility.Visible : txtPageTitle.Text = "📦 Yüklü Uygulamalar" : If _allPrograms.Count = 0 Then LoadPrograms()
                Case "history" : viewHistory.Visibility = Visibility.Visible : txtPageTitle.Text = "📈 Oturum Geçmişi" : UpdateHistoryPage()
                Case "crashlogs" : viewCrashLogs.Visibility = Visibility.Visible : txtPageTitle.Text = "⚠️ Hata Raporları" : LoadCrashLogs()
                Case "tools" : viewTools.Visibility = Visibility.Visible : txtPageTitle.Text = "🛠️ Sistem Araçları" : LoadTools()
                Case "alarms" : viewAlarms.Visibility = Visibility.Visible : txtPageTitle.Text = "🔔 Ayarlar"
            Case "tasks" : viewTasks.Visibility = Visibility.Visible : txtPageTitle.Text = "⏰ Zamanlanmış Görevler" : If _allTasks.Count = 0 Then LoadTasks()
        End Select

        Dim navBtns As New Dictionary(Of String, Button) From {
                {"dashboard", btnNavDashboard}, {"cpu", btnNavCpu}, {"ram", btnNavRam},
                {"disk", btnNavDisk}, {"network", btnNavNetwork}, {"firewall", btnNavFirewall}, {"processes", btnNavProcesses},
                {"services", btnNavServices}, {"startup", btnNavStartup}, {"programs", btnNavPrograms}, {"history", btnNavHistory},
                {"crashlogs", btnNavCrashLogs}, {"tools", btnNavTools}, {"alarms", btnNavAlarms}
            }
            For Each kv In navBtns
                kv.Value.Style = If(kv.Key = tag, CType(Resources("NavBtnActive"), Style), CType(Resources("NavBtn"), Style))
            Next
        End Sub

        Private Sub EnqueueVal(q As Queue(Of Double), value As Double)
            q.Enqueue(value) : If q.Count > HISTORY_LEN Then q.Dequeue()
        End Sub

        Private Sub DrawChart(canvas As Canvas, data As Queue(Of Double), maxVal As Double, lineColor As Color, Optional showGrid As Boolean = False, Optional showLabels As Boolean = False, Optional unit As String = "%")
            ChartHelper.DrawLineChart(canvas, data, maxVal, lineColor, showGrid, showLabels, unit)
        End Sub

        Private Function FormatNet(kbps As Double) As String
            Return If(kbps >= 1024, $"{kbps / 1024.0:F2} MB/s", $"{kbps:F0} KB/s")
        End Function

        Protected Overrides Sub OnClosing(e As System.ComponentModel.CancelEventArgs)
            If _alarmConfig.MinimizeToTray Then
                e.Cancel = True
                Me.Hide()
                If _trayIcon IsNot Nothing Then
                    _trayIcon.Visible = True
                    _trayIcon.ShowBalloonTip(2000, "Sistem Monitörü", "Arka planda çalışıyor.", System.Windows.Forms.ToolTipIcon.Info)
                End If
            Else
                If _trayIcon IsNot Nothing Then
                    _trayIcon.Visible = False
                    _trayIcon.Dispose()
                End If
                _isStressTesting = False
                _timer.Stop()
                _hw.Dispose()
            End If
            MyBase.OnClosing(e)
        End Sub
    End Class
