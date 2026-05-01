Imports System.Diagnostics
Imports System.Management
Imports System.Net.NetworkInformation
Imports System.Net.Http
Imports System.Threading.Tasks
Imports Microsoft.Win32
Imports System.Runtime.InteropServices
Imports System.Linq
Imports System.Collections.Generic
Imports System.ServiceProcess
Imports System.Text
Imports System.IO
Imports Microsoft.Win32.SafeHandles
Imports System.Security.Cryptography.X509Certificates
Public Class HardwareService
    Implements IDisposable
    ' ── GERİ DÖNÜŞÜM KUTUSU İÇİN API'LER ──
    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure SHQUERYRBINFO
        Public cbSize As Integer
        Public i64Size As Long
        Public i64NumItems As Long
    End Structure

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function SHQueryRecycleBin(ByVal pszRootPath As String, ByRef pSHQueryRBInfo As SHQUERYRBINFO) As Integer
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function SHEmptyRecycleBin(ByVal hwnd As IntPtr, ByVal pszRootPath As String, ByVal dwFlags As UInteger) As Integer
    End Function

    Private Const SHERB_NOCONFIRMATION As UInteger = 1
    Private Const SHERB_NOPROGRESSUI As UInteger = 2
    Private Const SHERB_NOSOUND As UInteger = 4

    ' ── YENİ: ÇÖP DOSYA ANALİZİ (JUNK FILES) ──
    Public Function GetJunkFilesReport() As List(Of JunkReportItem)
        Dim result As New List(Of JunkReportItem)

        ' 1. Kullanıcı Temp Klasörü
        Dim userTemp = IO.Path.GetTempPath()
        result.Add(New JunkReportItem With {
                .Category = "Kullanıcı Temp (%Temp%)",
                .FolderPath = userTemp,
                .SizeBytes = CalculateFolderSize(userTemp)
            })

        ' 2. Windows Temp Klasörü
        Dim winTemp = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\Temp"
        result.Add(New JunkReportItem With {
                .Category = "Windows Temp",
                .FolderPath = winTemp,
                .SizeBytes = CalculateFolderSize(winTemp)
            })

        ' 3. Windows Prefetch (Yönetici İzni Gerektirir)
        Dim prefetch = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\Prefetch"
        result.Add(New JunkReportItem With {
                .Category = "Sistem Öngetirme (Prefetch)",
                .FolderPath = prefetch,
                .SizeBytes = CalculateFolderSize(prefetch)
            })

        ' 4. Geri Dönüşüm Kutusu (Tüm Sürücüler)
        Dim rbInfo As New SHQUERYRBINFO()
        rbInfo.cbSize = Marshal.SizeOf(GetType(SHQUERYRBINFO))
        Dim rbSize As Long = 0
        Try
            If SHQueryRecycleBin(Nothing, rbInfo) = 0 Then
                rbSize = rbInfo.i64Size
            End If
        Catch : End Try

        result.Add(New JunkReportItem With {
                .Category = "Geri Dönüşüm Kutusu",
                .FolderPath = "RecycleBin",
                .SizeBytes = rbSize
            })

        Return result
    End Function

    ' ── YENİ: ÇÖP DOSYA TEMİZLEME ──
    Public Function CleanJunkFiles() As Long
        Dim freedBytes As Long = 0
        Dim items = GetJunkFilesReport()

        For Each item In items
            If item.SizeBytes = 0 Then Continue For

            If item.FolderPath = "RecycleBin" Then
                Try
                    ' Geri dönüşüm kutusunu sessizce ve uyarısız boşalt
                    Dim flags = SHERB_NOCONFIRMATION Or SHERB_NOPROGRESSUI Or SHERB_NOSOUND
                    If SHEmptyRecycleBin(IntPtr.Zero, Nothing, flags) = 0 Then
                        freedBytes += item.SizeBytes
                    End If
                Catch : End Try
            Else
                ' Klasör içeriklerini sil (kilitli dosyalara dokunma)
                Dim bytesBefore = item.SizeBytes
                SafeDeleteFolderContents(item.FolderPath)
                Dim bytesAfter = CalculateFolderSize(item.FolderPath)
                freedBytes += (bytesBefore - bytesAfter)
            End If
        Next
        Return freedBytes
    End Function

    Private Sub SafeDeleteFolderContents(folderPath As String)
        Try
            Dim di As New IO.DirectoryInfo(folderPath)
            If Not di.Exists Then Return
            ' Dosyaları sil
            For Each file In di.GetFiles()
                Try : file.Delete() : Catch : End Try
            Next
            ' Alt klasörleri sil
            For Each Dirr In di.GetDirectories()
                Try : Dirr.Delete(True) : Catch : End Try
            Next
        Catch : End Try
    End Sub
    ' ── YENİ: MİNİ DUMP API'Sİ ──
    <DllImport("dbghelp.dll", EntryPoint:="MiniDumpWriteDump", CallingConvention:=CallingConvention.StdCall, CharSet:=CharSet.Unicode, ExactSpelling:=True, SetLastError:=True)>
    Private Shared Function MiniDumpWriteDump(
            ByVal hProcess As IntPtr,
            ByVal processId As Integer,
            ByVal hFile As SafeFileHandle,
            ByVal dumpType As UInteger,
            ByVal expParam As IntPtr,
            ByVal userStreamParam As IntPtr,
            ByVal callbackParam As IntPtr) As Boolean
    End Function

    Public Function CreateMiniDump(pid As Integer, dumpFilePath As String) As Boolean
        Try
            Dim targetProcess = Process.GetProcessById(pid)
            Using fs As New FileStream(dumpFilePath, FileMode.Create, FileAccess.ReadWrite, FileShare.Write)
                ' 2 = MiniDumpWithFullMemory (Tüm belleği içeren detaylı döküm)
                ' 0 = MiniDumpNormal (Sadece temel bilgileri içeren küçük döküm)
                Dim MINIDUMP_TYPE As UInteger = 2
                Return MiniDumpWriteDump(targetProcess.Handle, targetProcess.Id, fs.SafeFileHandle, MINIDUMP_TYPE, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero)
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function
    <DllImport("psapi.dll")>
        Private Shared Function EmptyWorkingSet(hwProc As IntPtr) As Integer
        End Function

        ' ── PERFORMANS SAYAÇLARI ──
        Private _cpuCounter As PerformanceCounter
        Private _ramCounter As PerformanceCounter
        Private _diskReadCounter As PerformanceCounter
        Private _diskWriteCounter As PerformanceCounter
        Private _netSentCounter As PerformanceCounter
        Private _netRecvCounter As PerformanceCounter
        Private _coreCounters() As PerformanceCounter

        ' ── STATİK BİLGİLER ──
        Private _totalRamMB As Long = 0
        Private _cpuName As String = ""
        Private _cpuCores As Integer = 0
        Private _cpuLogical As Integer = 0
        Private _gpuName As String = ""
        Private _gpuRamMB As Long = 0
        Private _osName As String = ""

        ' ── SÜREÇ CPU TAKİBİ ──
        Private _procCpuCache As New Dictionary(Of Integer, TimeSpan)()
        Private _procCpuLastTime As DateTime = DateTime.MinValue

        Public DeepInfo As New DeepHwInfo()
        Private _initialized As Boolean = False
        Private _disposed As Boolean = False

        Public Sub Initialize()
            If _initialized Then Return
            Try
                _cpuCounter = New PerformanceCounter("Processor", "% Processor Time", "_Total", True)
                _ramCounter = New PerformanceCounter("Memory", "Available MBytes", True)

                ' Çekirdek bazlı sayaçlar
                Try
                    Dim coreCount = Environment.ProcessorCount
                    ReDim _coreCounters(coreCount - 1)
                    For i = 0 To coreCount - 1
                        _coreCounters(i) = New PerformanceCounter("Processor", "% Processor Time", i.ToString(), True)
                    Next
                Catch
                End Try

                Try
                    Dim diskCat As New PerformanceCounterCategory("PhysicalDisk")
                    Dim diskInst = If(diskCat.GetInstanceNames().FirstOrDefault(Function(x) x <> "_Total"), "_Total")
                    _diskReadCounter = New PerformanceCounter("PhysicalDisk", "Disk Read Bytes/sec", diskInst, True)
                    _diskWriteCounter = New PerformanceCounter("PhysicalDisk", "Disk Write Bytes/sec", diskInst, True)
                Catch
                End Try

                Try
                    Dim netCat As New PerformanceCounterCategory("Network Interface")
                    Dim netInst = netCat.GetInstanceNames().FirstOrDefault(Function(x) Not x.ToLower().Contains("loopback") AndAlso Not x.ToLower().Contains("teredo"))
                    If netInst IsNot Nothing Then
                        _netSentCounter = New PerformanceCounter("Network Interface", "Bytes Sent/sec", netInst, True)
                        _netRecvCounter = New PerformanceCounter("Network Interface", "Bytes Received/sec", netInst, True)
                    End If
                Catch
                End Try

                LoadStaticInfo()
                LoadDeepInfo()
                _initialized = True
            Catch ex As Exception
            End Try
        End Sub

        Private Sub LoadStaticInfo()
            Try
                Using searcher As New ManagementObjectSearcher("SELECT Name, NumberOfCores, NumberOfLogicalProcessors FROM Win32_Processor")
                    For Each obj In searcher.Get()
                        _cpuName = If(obj("Name")?.ToString()?.Trim(), "Bilinmiyor")
                        _cpuCores = Convert.ToInt32(obj("NumberOfCores"))
                        _cpuLogical = Convert.ToInt32(obj("NumberOfLogicalProcessors"))
                        Exit For
                    Next
                End Using
            Catch
            End Try
            Try
                Using searcher As New ManagementObjectSearcher("SELECT Caption, TotalVisibleMemorySize FROM Win32_OperatingSystem")
                    For Each obj In searcher.Get()
                        _osName = If(obj("Caption")?.ToString(), "Windows")
                        _totalRamMB = Convert.ToInt64(obj("TotalVisibleMemorySize")) \ 1024
                        Exit For
                    Next
                End Using
            Catch
            End Try
            Try
                Using searcher As New ManagementObjectSearcher("SELECT Name, AdapterRAM FROM Win32_VideoController")
                    For Each obj In searcher.Get()
                        _gpuName = If(obj("Name")?.ToString()?.Trim(), "Bilinmeyen GPU")
                        _gpuRamMB = Convert.ToInt64(If(obj("AdapterRAM"), 0)) \ 1048576
                        Exit For
                    Next
                End Using
            Catch
            End Try
        End Sub

        Private Sub LoadDeepInfo()
            Try
                Using searcher As New ManagementObjectSearcher("SELECT Manufacturer, Product FROM Win32_BaseBoard")
                    For Each obj In searcher.Get()
                        DeepInfo.Motherboard = $"{obj("Manufacturer")} {obj("Product")}"
                        Exit For
                    Next
                End Using
                Using searcher As New ManagementObjectSearcher("SELECT SMBIOSBIOSVersion FROM Win32_BIOS")
                    For Each obj In searcher.Get()
                        DeepInfo.BIOS = obj("SMBIOSBIOSVersion")?.ToString()
                        Exit For
                    Next
                End Using
                Using searcher As New ManagementObjectSearcher("SELECT Speed FROM Win32_PhysicalMemory")
                    For Each obj In searcher.Get()
                        DeepInfo.RamSpeed = $"{obj("Speed")} MHz"
                        Exit For
                    Next
                End Using
            Catch
            End Try
        End Sub

        ' ── ANLAK DEĞERLER ──

        Public Function GetCpuUsage() As Single
            Try
                Return If(_cpuCounter IsNot Nothing, _cpuCounter.NextValue(), 0.0F)
            Catch
                Return 0
            End Try
        End Function

        Public Function GetCpuCoreUsages() As List(Of CpuCoreInfo)
            Dim result As New List(Of CpuCoreInfo)
            If _coreCounters Is Nothing Then Return result
            For i = 0 To _coreCounters.Length - 1
                Try
                    Dim pct = Math.Round(_coreCounters(i).NextValue(), 1)
                    result.Add(New CpuCoreInfo With {.CoreIndex = i, .UsagePct = Math.Min(100, Math.Max(0, pct))})
                Catch
                    result.Add(New CpuCoreInfo With {.CoreIndex = i, .UsagePct = 0})
                End Try
            Next
            Return result
        End Function

        Public Function GetAvailableRamMB() As Long
            Try
                Return CLng(If(_ramCounter IsNot Nothing, _ramCounter.NextValue(), 0.0F))
            Catch
                Return 0
            End Try
        End Function

        Public Function GetTotalRamMB() As Long
            If _totalRamMB <= 0 Then LoadStaticInfo()
            Return If(_totalRamMB > 0, _totalRamMB, 8192)
        End Function

        Public Function GetDiskReadMBps() As Double
            Try
                Return If(_diskReadCounter IsNot Nothing, _diskReadCounter.NextValue(), 0.0F) / 1048576.0
            Catch
                Return 0
            End Try
        End Function

        Public Function GetDiskWriteMBps() As Double
            Try
                Return If(_diskWriteCounter IsNot Nothing, _diskWriteCounter.NextValue(), 0.0F) / 1048576.0
            Catch
                Return 0
            End Try
        End Function

        Public Function GetNetSentKBps() As Double
            Try
                Return If(_netSentCounter IsNot Nothing, _netSentCounter.NextValue(), 0.0F) / 1024.0
            Catch
                Return 0
            End Try
        End Function

        Public Function GetNetRecvKBps() As Double
            Try
                Return If(_netRecvCounter IsNot Nothing, _netRecvCounter.NextValue(), 0.0F) / 1024.0
            Catch
                Return 0
            End Try
        End Function

    ' ── GPU YÜKÜ (Win10+ WMI) ──
    ' Sınıfın içine (fonksiyonların üstüne) bu iki değişkeni ekleyin
    Private _gpuCounters As List(Of PerformanceCounter)
    Private _lastGpuCheck As DateTime = DateTime.MinValue

    ' Eski GetGpuLoad fonksiyonunu tamamen silip bunu yapıştırın:
    Public Function GetGpuLoad() As Double
        Try
            ' Açılıp kapanan uygulamaları yakalamak için listeyi her 3 saniyede bir güncelleriz
            If _gpuCounters Is Nothing OrElse (DateTime.Now - _lastGpuCheck).TotalSeconds > 3 Then
                _lastGpuCheck = DateTime.Now
                If _gpuCounters IsNot Nothing Then
                    For Each c In _gpuCounters
                        Try : c.Dispose() : Catch : End Try
                    Next
                End If
                _gpuCounters = New List(Of PerformanceCounter)()

                Dim cat As New PerformanceCounterCategory("GPU Engine")
                Dim instances = cat.GetInstanceNames()
                For Each inst In instances
                    ' Görev Yöneticisi gibi sadece "3D" motorlarını dinliyoruz
                    If inst.Contains("engtype_3D") Then
                        Dim pc As New PerformanceCounter("GPU Engine", "Utilization Percentage", inst, True)
                        pc.NextValue() ' Sayacı uyandırmak için ilk okuma
                        _gpuCounters.Add(pc)
                    End If
                Next
            End If

            If _gpuCounters.Count > 0 Then
                Dim totalLoad As Double = 0
                For Each pc In _gpuCounters
                    Try
                        totalLoad += pc.NextValue()
                    Catch
                    End Try
                Next
                ' Toplam değerin %100'ü geçmemesi için sınır koyuyoruz
                Return Math.Min(100.0, Math.Round(totalLoad, 1))
            End If
        Catch
        End Try

        Return -1 ' Desteklenmeyen donanımlarda N/A basması için
    End Function

    ' ── SÜREÇLER (CPU% + RAM) ──
    Public Function GetTopProcesses(count As Integer) As List(Of SysProcessInfo)
            Dim result As New List(Of SysProcessInfo)
            Dim now = DateTime.Now
            Dim elapsedSec = If(_procCpuLastTime = DateTime.MinValue, 1.0,
                             Math.Max(0.5, (now - _procCpuLastTime).TotalSeconds))
            _procCpuLastTime = now
            Dim coreCount = Math.Max(1, Environment.ProcessorCount)
            Dim newCache As New Dictionary(Of Integer, TimeSpan)()

            Try
                Dim procs = Process.GetProcesses()
                For Each p In procs
                    Try
                        Dim cpu = 0.0
                        Try
                            Dim totalTime = p.TotalProcessorTime
                            newCache(p.Id) = totalTime
                            If _procCpuCache.ContainsKey(p.Id) Then
                                Dim delta = (totalTime - _procCpuCache(p.Id)).TotalSeconds
                                cpu = Math.Round(delta / (elapsedSec * coreCount) * 100, 1)
                                cpu = Math.Min(100, Math.Max(0, cpu))
                            End If
                        Catch
                        End Try
                        result.Add(New SysProcessInfo With {
                            .Name = p.ProcessName,
                            .PID = p.Id,
                            .MemoryMB = Math.Round(p.WorkingSet64 / 1048576.0, 1),
                            .Threads = p.Threads.Count,
                            .CpuPct = cpu
                        })
                    Catch
                    End Try
                Next
            Catch
            End Try

            _procCpuCache = newCache
            Return result.OrderByDescending(Function(x) x.MemoryMB).Take(count).ToList()
        End Function

        ' ── BELLEK TEMİZLEME ──
        Public Function CleanAllMemory() As Integer
            Dim cleaned = 0
            Try
                For Each p In Process.GetProcesses()
                    Try
                        EmptyWorkingSet(p.Handle)
                        cleaned += 1
                    Catch
                    End Try
                Next
            Catch
            End Try
            Return cleaned
        End Function

        ' ── DİSK İŞLEMLERİ ──
        Public Function GetDirectorySizes(drivePath As String) As List(Of DiskSpaceItem)
            Dim result As New List(Of DiskSpaceItem)
            Try
                Dim di As New IO.DirectoryInfo(drivePath)
                For Each d In di.GetDirectories()
                    Try
                        Dim size = CalculateFolderSize(d.FullName)
                        If size > 0 Then result.Add(New DiskSpaceItem With {.FolderName = d.Name, .SizeBytes = size})
                    Catch
                    End Try
                Next
            Catch
            End Try
            Return result.OrderByDescending(Function(x) x.SizeBytes).ToList()
        End Function

        Private Function CalculateFolderSize(folderPath As String) As Long
            Dim size As Long = 0
            Try
                Dim dInfo As New IO.DirectoryInfo(folderPath)
                For Each fi In dInfo.GetFiles()
                    size += fi.Length
                Next
                For Each di In dInfo.GetDirectories()
                    size += CalculateFolderSize(di.FullName)
                Next
            Catch
            End Try
            Return size
        End Function

        Public Function GetDiskDrives() As List(Of DiskDriveInfo)
            Dim result As New List(Of DiskDriveInfo)
            Try
                For Each drive As IO.DriveInfo In IO.DriveInfo.GetDrives()
                    If Not drive.IsReady Then Continue For
                    If drive.DriveType <> IO.DriveType.Fixed AndAlso drive.DriveType <> IO.DriveType.Removable Then Continue For
                    result.Add(New DiskDriveInfo With {
                        .Letter = drive.Name,
                        .Label = drive.VolumeLabel,
                        .TotalGB = Math.Round(drive.TotalSize / 1073741824.0, 1),
                        .FreeGB = Math.Round(drive.AvailableFreeSpace / 1073741824.0, 1)
                    })
                Next
            Catch
            End Try
            Return result
        End Function

        Public Function GetDiskHealth() As List(Of DiskHealth)
            Dim result As New List(Of DiskHealth)
            Try
                Using searcher As New ManagementObjectSearcher("SELECT Model, Status FROM Win32_DiskDrive")
                    For Each obj In searcher.Get()
                        result.Add(New DiskHealth With {
                            .Model = If(obj("Model")?.ToString(), "Bilinmeyen Disk"),
                            .Status = If(obj("Status")?.ToString(), "Bilinmiyor")
                        })
                    Next
                End Using
            Catch
            End Try
            Return result
        End Function

        ' ── AĞ ──
        Public Function GetPing() As Long
            Try
                Using p As New Ping()
                    Dim reply = p.Send("8.8.8.8", 1000)
                    If reply.Status = IPStatus.Success Then Return reply.RoundtripTime
                End Using
            Catch
            End Try
            Return -1
        End Function

        Public Async Function GetPublicIpAsync() As Task(Of String)
            Try
                Using client As New HttpClient()
                    client.Timeout = TimeSpan.FromSeconds(5)
                    Return Await client.GetStringAsync("https://api.ipify.org")
                End Using
            Catch
                Return "Bağlantı Yok"
            End Try
        End Function

        Public Function GetActiveNetworkConnections() As List(Of NetConnection)
            Dim result As New List(Of NetConnection)
            Try
                Dim pInfo As New ProcessStartInfo("netstat", "-ano") With {
                    .RedirectStandardOutput = True,
                    .UseShellExecute = False,
                    .CreateNoWindow = True
                }
                Using proc = Process.Start(pInfo)
                    Dim lines = proc.StandardOutput.ReadToEnd().Split(
                        New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                    For Each line In lines
                        Dim parts = line.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                        If parts.Length >= 4 AndAlso (parts(0) = "TCP" OrElse parts(0) = "UDP") Then
                            Dim pid As Integer
                            If Integer.TryParse(parts(parts.Length - 1), pid) AndAlso pid > 0 Then
                                Dim procName = "Bilinmiyor"
                                Try
                                    procName = Process.GetProcessById(pid).ProcessName
                                Catch
                                End Try
                                If Not parts(2).StartsWith("*") AndAlso Not parts(2).StartsWith("0.0.0.0") AndAlso Not parts(2).StartsWith("[::]") Then
                                    result.Add(New NetConnection With {
                                        .Protocol = parts(0),
                                        .LocalAddress = parts(1),
                                        .RemoteAddress = parts(2),
                                        .State = If(parts.Length = 5, parts(3), "N/A"),
                                        .ProcessName = procName
                                    })
                                End If
                            End If
                        End If
                    Next
                End Using
            Catch
            End Try
            Return result.OrderBy(Function(x) x.ProcessName).Take(100).ToList()
        End Function

        Public Function GetNetworkAdapters() As List(Of NetworkAdapterInfo)
            Dim result As New List(Of NetworkAdapterInfo)
            Try
                For Each adapter In NetworkInterface.GetAllNetworkInterfaces()
                    If adapter.OperationalStatus <> OperationalStatus.Up Then Continue For
                    If adapter.NetworkInterfaceType = NetworkInterfaceType.Loopback Then Continue For
                    Dim ip = ""
                    Try
                        For Each addr In adapter.GetIPProperties().UnicastAddresses
                            If addr.Address.AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then
                                ip = addr.Address.ToString()
                                Exit For
                            End If
                        Next
                    Catch
                    End Try
                    Dim mac = ""
                    Try
                        Dim bytes = adapter.GetPhysicalAddress().GetAddressBytes()
                        mac = String.Join(":", bytes.Select(Function(b) b.ToString("X2")))
                    Catch
                    End Try
                    Dim speedMbps = ""
                    Try
                        speedMbps = $"{adapter.Speed / 1000000.0:F0} Mbps"
                    Catch
                    End Try
                    result.Add(New NetworkAdapterInfo With {
                        .AdapterName = adapter.Name,
                        .AdapterType = adapter.NetworkInterfaceType.ToString(),
                        .IpAddress = ip,
                        .MacAddress = mac,
                        .SpeedMbps = speedMbps
                    })
                Next
            Catch
            End Try
            Return result
        End Function

        ' ── SİSTEM BİLGİSİ ──
        Public Function GetSystemInfoText() As String
            Dim sb As New StringBuilder()
            sb.AppendLine($"🖥  {Environment.MachineName}")
            sb.AppendLine($"🔲  {_cpuName} ({_cpuCores} Çekirdek)")
            Dim gpuRamStr = If(_gpuRamMB > 0, $" ({Math.Round(_gpuRamMB / 1024.0, 1)} GB VRAM)", "")
            sb.AppendLine($"🎮  {_gpuName}{gpuRamStr}")
            sb.AppendLine($"💾  {Math.Round(GetTotalRamMB() / 1024.0, 1)} GB RAM")
            sb.AppendLine($"🪟  {_osName}")
            Return sb.ToString().TrimEnd()
        End Function

        Public Function GetUptime() As TimeSpan
            Try
                Using searcher As New ManagementObjectSearcher("SELECT LastBootUpTime FROM Win32_OperatingSystem")
                    For Each obj In searcher.Get()
                        Return DateTime.Now - ManagementDateTimeConverter.ToDateTime(obj("LastBootUpTime").ToString())
                    Next
                End Using
            Catch
            End Try
            Return TimeSpan.Zero
        End Function

        Public Function GetCpuTemperature() As Double
            Try
                Using searcher As New ManagementObjectSearcher("root\WMI", "SELECT CurrentTemperature FROM MSAcpi_ThermalZoneTemperature")
                    For Each obj In searcher.Get()
                        Return Math.Round((Convert.ToDouble(obj("CurrentTemperature")) / 10.0) - 273.15, 1)
                    Next
                End Using
            Catch
            End Try
            Return 0
        End Function

        Public Function GetBatteryInfo() As BatteryInfo
            Dim bInfo As New BatteryInfo()
            Try
                Using searcher As New ManagementObjectSearcher("SELECT EstimatedChargeRemaining, BatteryStatus FROM Win32_Battery")
                    For Each obj In searcher.Get()
                        bInfo.HasBattery = True
                        bInfo.ChargePercent = Convert.ToInt32(obj("EstimatedChargeRemaining"))
                        Select Case Convert.ToInt32(obj("BatteryStatus"))
                            Case 1 : bInfo.Status = "Deşarj"
                            Case 2 : bInfo.Status = "Şarj Oluyor"
                            Case 3 : bInfo.Status = "Tam Dolu"
                            Case Else : bInfo.Status = "Bilinmiyor"
                        End Select
                        Exit For
                    Next
                End Using
            Catch
            End Try
            Return bInfo
        End Function

    ' ── BAŞLANGIÇ UYGULAMALARI ──
    ' ── GÜNCELLENDİ: BAŞLANGIÇ UYGULAMALARI (DİJİTAL İMZA KONTROLÜ) ──
    Public Function GetStartupApps() As List(Of StartupApp)
        Dim result As New List(Of StartupApp)
        Try
            ' 1. Kullanıcı (HKCU) Başlangıç Kayıtları
            Using key = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", False)
                If key IsNot Nothing Then
                    For Each valName In key.GetValueNames()
                        Dim cmd = key.GetValue(valName).ToString()
                        Dim app As New StartupApp With {.Name = valName, .Command = cmd, .RegistryLocation = "Kullanıcı (HKCU)"}
                        CheckDigitalSignature(app)
                        result.Add(app)
                    Next
                End If
            End Using

            ' 2. Sistem (HKLM) Başlangıç Kayıtları
            Using key = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", False)
                If key IsNot Nothing Then
                    For Each valName In key.GetValueNames()
                        Dim cmd = key.GetValue(valName).ToString()
                        Dim app As New StartupApp With {.Name = valName, .Command = cmd, .RegistryLocation = "Sistem (HKLM)"}
                        CheckDigitalSignature(app)
                        result.Add(app)
                    Next
                End If
            End Using
        Catch
        End Try

        Return result.OrderBy(Function(x) x.Name).ToList()
    End Function

    ' ── YARDIMCI: KARMAŞIK KOMUTLARDAN DOSYA YOLUNU ÇIKARIR VE İMZAYI DENETLER ──
    Private Sub CheckDigitalSignature(ByRef app As StartupApp)
        Try
            Dim cmd = app.Command.Trim()
            Dim filePath As String = ""

            ' Tırnak içindeyse yolu doğrudan al ("C:\app.exe" -minimized)
            If cmd.StartsWith("""") Then
                Dim endQuote = cmd.IndexOf("""", 1)
                If endQuote > 1 Then filePath = cmd.Substring(1, endQuote - 1)
            Else
                ' Tırnak yoksa .exe'ye kadar olan kısmı al
                Dim exeIndex = cmd.IndexOf(".exe", StringComparison.OrdinalIgnoreCase)
                If exeIndex > 0 Then
                    filePath = cmd.Substring(0, exeIndex + 4)
                Else
                    filePath = cmd.Split(" "c)(0)
                End If
            End If

            ' Dosya fiziksel olarak diskte bulunuyorsa imzasını oku
            If IO.File.Exists(filePath) Then
                Dim cert = X509Certificate.CreateFromSignedFile(filePath)
                Dim cert2 = New X509Certificate2(cert)
                app.Publisher = cert2.GetNameInfo(X509NameType.SimpleName, False)
                app.IsSigned = True
            Else
                app.Publisher = "Dosya Bulunamadı"
            End If
        Catch ex As Exception
            ' Dosya var ama imzasızsa bu bloğa düşer
            app.Publisher = "İmzasız / Şüpheli"
            app.IsSigned = False
        End Try
    End Sub
    ' ── YENİ: GÜVENLİK ÖZETİ SORGUSU (WMI) ──
    Public Function GetSecurityStatus() As List(Of SecurityStatusItem)
        Dim result As New List(Of SecurityStatusItem)
        ' Kontrol edilecek servisler: MpsSvc (Firewall), WinDefend (Defender), wuauserv (Updates)
        Dim servicesToCheck As String() = {"MpsSvc", "WinDefend", "wuauserv"}
        Dim displayNames As String() = {"Güvenlik Duvarı", "Windows Defender", "Otomatik Güncellemeler"}

        Try
            For i As Integer = 0 To servicesToCheck.Length - 1
                Dim svcName = servicesToCheck(i)
                Dim isRunning = False

                Using searcher As New ManagementObjectSearcher("SELECT State FROM Win32_Service WHERE Name='" & svcName & "'")
                    For Each obj In searcher.Get()
                        If obj("State").ToString().ToLower() = "running" Then
                            isRunning = True
                        End If
                    Next
                End Using

                result.Add(New SecurityStatusItem With {
                        .FeatureName = displayNames(i),
                        .StatusText = If(isRunning, "AKTİF", "KAPALI"),
                        .IsSecure = isRunning
                    })
            Next
        Catch
        End Try
        Return result
    End Function
    Public Sub SetRunAtStartup(enable As Boolean)
            Try
                Using key = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
                    Dim appName = "SysMonitorPro"
                    Dim appPath = """" & Reflection.Assembly.GetExecutingAssembly().Location & """"
                    If enable Then
                        key.SetValue(appName, appPath)
                    Else
                        If key.GetValue(appName) IsNot Nothing Then key.DeleteValue(appName)
                    End If
                End Using
            Catch
            End Try
        End Sub

        ' ── ÇÖKME LOGLARI ──
        Public Function GetCrashLogs() As List(Of CrashLogEntry)
            Dim logs As New List(Of CrashLogEntry)
            Try
                For Each logName In {"System", "Application"}
                    Using ev As New EventLog(logName)
                        For i As Integer = ev.Entries.Count - 1 To Math.Max(0, ev.Entries.Count - 150) Step -1
                            Dim entry = ev.Entries(i)
                            If entry.TimeGenerated < DateTime.Now.AddDays(-1) Then Exit For
                            If entry.EntryType = EventLogEntryType.Error OrElse entry.EntryType = EventLogEntryType.Warning Then
                                logs.Add(New CrashLogEntry With {
                                    .Time = entry.TimeGenerated.ToString("dd.MM.yyyy HH:mm"),
                                    .Level = entry.EntryType.ToString(),
                                    .Source = entry.Source,
                                    .Message = entry.Message.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
                                })
                            End If
                        Next
                    End Using
                Next
            Catch
            End Try
            Try
                Return logs.OrderByDescending(Function(x) DateTime.ParseExact(x.Time, "dd.MM.yyyy HH:mm", Nothing)).ToList()
            Catch
                Return logs
            End Try
        End Function

        ' ── YENİ: WINDOWS SERVİSLERİ ──
        Public Function GetServices() As List(Of ServiceInfo)
            Dim result As New List(Of ServiceInfo)
            Try
                For Each svc In ServiceController.GetServices()
                    Try
                        result.Add(New ServiceInfo With {
                            .ServiceName = svc.ServiceName,
                            .DisplayName = svc.DisplayName,
                            .Status = svc.Status.ToString(),
                            .StartType = svc.StartType.ToString()
                        })
                    Catch
                    End Try
                Next
            Catch
            End Try
            Return result.OrderBy(Function(x) x.DisplayName).ToList()
        End Function

        Public Function StartService(serviceName As String) As Boolean
            Try
                Dim svc As New ServiceController(serviceName)
                If svc.Status = ServiceControllerStatus.Stopped Then
                    svc.Start()
                    svc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10))
                    Return True
                End If
            Catch
            End Try
            Return False
        End Function

        Public Function StopService(serviceName As String) As Boolean
            Try
                Dim svc As New ServiceController(serviceName)
                If svc.Status = ServiceControllerStatus.Running Then
                    svc.Stop()
                    svc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(10))
                    Return True
                End If
            Catch
            End Try
            Return False
        End Function

        ' ── YENİ: YÜKLÜ PROGRAMLAR ──
        Public Function GetInstalledPrograms() As List(Of InstalledProgram)
            Dim result As New List(Of InstalledProgram)
            Dim regPaths As String() = {
                "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            }
            For Each regPath In regPaths
                Try
                    Using key = Registry.LocalMachine.OpenSubKey(regPath)
                        If key IsNot Nothing Then
                            For Each subKeyName In key.GetSubKeyNames()
                                Try
                                    Using subKey = key.OpenSubKey(subKeyName)
                                        Dim name = subKey?.GetValue("DisplayName")?.ToString()
                                        If Not String.IsNullOrWhiteSpace(name) Then
                                            result.Add(New InstalledProgram With {
                                                .Name = name,
                                                .Publisher = If(subKey.GetValue("Publisher")?.ToString(), ""),
                                                .Version = If(subKey.GetValue("DisplayVersion")?.ToString(), ""),
                                                .InstallDate = If(subKey.GetValue("InstallDate")?.ToString(), "")
                                            })
                                        End If
                                    End Using
                                Catch
                                End Try
                            Next
                        End If
                    End Using
                Catch
                End Try
            Next
            Dim uniq = result.GroupBy(Function(x) x.Name).Select(Function(g) g.First()).ToList()
            Return uniq.OrderBy(Function(x) x.Name).ToList()
        End Function

        ' ── YENİ: GÜÇ PLANLARI ──
        Public Function GetPowerPlans() As List(Of PowerPlanInfo)
            Dim result As New List(Of PowerPlanInfo)
            Try
                Dim psi As New ProcessStartInfo("powercfg", "/list") With {
                    .RedirectStandardOutput = True,
                    .UseShellExecute = False,
                    .CreateNoWindow = True
                }
                Using proc = Process.Start(psi)
                    Dim output = proc.StandardOutput.ReadToEnd()
                    For Each line In output.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                        If line.Contains("Güç Şeması GUID") OrElse line.Contains("Power Scheme GUID") OrElse (line.Trim().StartsWith("GUID") AndAlso line.Contains(":")) Then
                            Try
                                Dim guidStart = line.IndexOf(":"c) + 1
                                Dim rawGuid = line.Substring(guidStart).Trim().Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)(0).Trim()
                                Dim planName = ""
                                Dim parenStart = line.IndexOf("("c)
                                Dim parenEnd = line.IndexOf(")"c)
                                If parenStart >= 0 AndAlso parenEnd > parenStart Then
                                    planName = line.Substring(parenStart + 1, parenEnd - parenStart - 1)
                                End If
                                Dim isActive = line.ToLower().Contains("*") OrElse line.ToLower().Contains("active")
                                result.Add(New PowerPlanInfo With {
                                    .Guid = rawGuid,
                                    .Name = If(planName, rawGuid),
                                    .IsActive = isActive
                                })
                            Catch
                            End Try
                        End If
                    Next
                End Using
            Catch
            End Try
            ' Fallback: bilinen GUID'ler
            If result.Count = 0 Then
                result.Add(New PowerPlanInfo With {.Name = "Güç Tasarrufu", .Guid = "a1841308-3541-4fab-bc81-f71556f20b4a", .IsActive = False})
                result.Add(New PowerPlanInfo With {.Name = "Dengeli", .Guid = "381b4222-f694-41f0-9685-ff5bb260df2e", .IsActive = False})
                result.Add(New PowerPlanInfo With {.Name = "Yüksek Performans", .Guid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c", .IsActive = False})
            End If
            Return result
        End Function

        Public Sub SetPowerPlan(guid As String)
            Try
                Dim psi As New ProcessStartInfo("powercfg", $"/setactive {guid}") With {
                    .UseShellExecute = False,
                    .CreateNoWindow = True
                }
                Process.Start(psi)?.WaitForExit(3000)
            Catch
            End Try
        End Sub

        ' ── YENİ: RAPOR DIŞA AKTARMA ──
        Public Sub ExportSystemReport(filePath As String, cpuPct As Double, ramPct As Double, usedRamMB As Long)
            Try
                Dim sb As New StringBuilder()
                sb.AppendLine("═══════════════════════════════════════════════════════")
                sb.AppendLine("         SistemMonitörü — Sistem Raporu")
                sb.AppendLine($"        Oluşturma: {DateTime.Now:dd.MM.yyyy HH:mm:ss}")
                sb.AppendLine("═══════════════════════════════════════════════════════")
                sb.AppendLine()
                sb.AppendLine("── DONANIM BİLGİSİ ────────────────────────────────")
                sb.AppendLine(GetSystemInfoText())
                sb.AppendLine($"Anakart  : {DeepInfo.Motherboard}")
                sb.AppendLine($"BIOS     : {DeepInfo.BIOS}")
                sb.AppendLine($"RAM Hızı : {DeepInfo.RamSpeed}")
                sb.AppendLine()
                sb.AppendLine("── ANLIK DURUM ─────────────────────────────────────")
                sb.AppendLine($"CPU Kullanımı : %{cpuPct:F1}")
                sb.AppendLine($"RAM Kullanımı : %{ramPct:F1}  ({usedRamMB / 1024.0:F1} GB / {GetTotalRamMB() / 1024.0:F1} GB)")
                Dim up = GetUptime()
                sb.AppendLine($"Çalışma Süresi: {up.Days}g {up.Hours:D2}:{up.Minutes:D2}:{up.Seconds:D2}")
                sb.AppendLine()
                sb.AppendLine("── DİSK DURUMLARI ──────────────────────────────────")
                For Each d In GetDiskDrives()
                    sb.AppendLine($"{d.DisplayName,-20} {d.UsedGB:F1}/{d.TotalGB:F1} GB  (%{d.UsedPct:F1} dolu)")
                Next
                sb.AppendLine()
                sb.AppendLine("── AĞ ADAPTÖRLERİ ──────────────────────────────────")
                For Each a In GetNetworkAdapters()
                    sb.AppendLine($"{a.AdapterName,-25} IP:{a.IpAddress,-16} {a.SpeedMbps}")
                Next
                sb.AppendLine()
                sb.AppendLine("═══════════════════════════════════════════════════════")
                IO.File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8)
            Catch
            End Try
        End Sub

        ' ── YENİ: FIREWALL KURALLARI ──
        Public Function GetFirewallRules() As List(Of FirewallRule)
            Dim result As New List(Of FirewallRule)
            Try
                Dim pInfo As New ProcessStartInfo("netsh", "advfirewall firewall show rule name=all") With {
                    .RedirectStandardOutput = True,
                    .UseShellExecute = False,
                    .CreateNoWindow = True,
                    .StandardOutputEncoding = System.Text.Encoding.GetEncoding(857) ' Türkçe karakterler için
                }

                Using proc = Process.Start(pInfo)
                    Dim output = proc.StandardOutput.ReadToEnd()
                    Dim lines = output.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                    Dim currentRule As FirewallRule = Nothing

                    For Each line In lines
                        Dim text = line.Trim()
                        If text.StartsWith("Kural Adı:") OrElse text.StartsWith("Rule Name:") Then
                            If currentRule IsNot Nothing Then result.Add(currentRule)
                            currentRule = New FirewallRule()
                            currentRule.RuleName = text.Substring(text.IndexOf(":") + 1).Trim()
                        ElseIf currentRule IsNot Nothing Then
                            If text.StartsWith("Yön:") OrElse text.StartsWith("Direction:") Then
                                currentRule.Direction = text.Substring(text.IndexOf(":") + 1).Trim()
                            ElseIf text.StartsWith("Eylem:") OrElse text.StartsWith("Action:") Then
                                currentRule.Action = text.Substring(text.IndexOf(":") + 1).Trim()
                            ElseIf text.StartsWith("Profil:") OrElse text.StartsWith("Profiles:") Then
                                currentRule.Profile = text.Substring(text.IndexOf(":") + 1).Trim()
                            End If
                        End If
                    Next
                    If currentRule IsNot Nothing Then result.Add(currentRule) ' Son kuralı ekle
                End Using
            Catch
            End Try

            Return result.OrderBy(Function(x) x.RuleName).ToList()
        End Function

        ' ── DISPOSE ──
        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            If _cpuCounter IsNot Nothing Then _cpuCounter.Dispose()
            If _ramCounter IsNot Nothing Then _ramCounter.Dispose()
            If _diskReadCounter IsNot Nothing Then _diskReadCounter.Dispose()
            If _diskWriteCounter IsNot Nothing Then _diskWriteCounter.Dispose()
            If _netSentCounter IsNot Nothing Then _netSentCounter.Dispose()
            If _netRecvCounter IsNot Nothing Then _netRecvCounter.Dispose()
            If _coreCounters IsNot Nothing Then
                For Each c In _coreCounters
                    If c IsNot Nothing Then c.Dispose()
                Next
            End If
            _disposed = True
        End Sub
    ' ── YENİ: SÜRÜCÜ YÖNETİCİSİ (Win32_SystemDriver) ──
    Public Function GetDrivers() As List(Of DriverInfo)
        Dim result As New List(Of DriverInfo)
        Try
            Using searcher As New ManagementObjectSearcher("SELECT Name, DisplayName, State, Status FROM Win32_SystemDriver")
                For Each obj In searcher.Get()
                    Try
                        result.Add(New DriverInfo With {
                            .Name = If(obj("Name")?.ToString(), "Bilinmiyor"),
                            .DisplayName = If(obj("DisplayName")?.ToString(), "Bilinmiyor"),
                            .State = If(obj("State")?.ToString(), "Bilinmiyor"),
                            .Status = If(obj("Status")?.ToString(), "OK") ' Varsayılan OK
                        })
                    Catch
                    End Try
                Next
            End Using
        Catch
        End Try
        ' Adına göre alfabetik sırala
        Return result.OrderBy(Function(x) x.DisplayName).ToList()
    End Function
    Public Class JunkReportItem
        Public Property Category As String = ""
        Public Property FolderPath As String = ""
        Public Property SizeBytes As Long = 0
        Public ReadOnly Property SizeStr As String
            Get
                If SizeBytes = 0 Then Return "0 KB"
                If SizeBytes > 1048576 Then Return $"{(SizeBytes / 1048576.0):F2} MB"
                Return $"{(SizeBytes / 1024.0):F1} KB"
            End Get
        End Property
    End Class
    ' ── YENİ: PAYLAŞIMLI BİLGİSAYAR / AÇIK OTURUMLAR (quser) ──
    Public Function GetLoggedInUsers() As List(Of UserSessionInfo)
        Dim result As New List(Of UserSessionInfo)
        Dim sessionRam As New Dictionary(Of Integer, Double)

        ' 1. Önce her oturum ID'si (SessionId) için tüketilen toplam RAM'i hesapla
        Try
            For Each p In Process.GetProcesses()
                Try
                    Dim sId = p.SessionId
                    Dim ram = p.WorkingSet64 / 1048576.0 ' MB'a çevir
                    If sessionRam.ContainsKey(sId) Then
                        sessionRam(sId) += ram
                    Else
                        sessionRam(sId) = ram
                    End If
                Catch
                End Try
            Next
        Catch
        End Try

        ' 2. quser komutu ile açık kullanıcıları çek
        Try
            Dim pInfo As New ProcessStartInfo("quser") With {
                .RedirectStandardOutput = True,
                .UseShellExecute = False,
                .CreateNoWindow = True
            }
            Using proc = Process.Start(pInfo)
                Dim output = proc.StandardOutput.ReadToEnd()
                Dim lines = output.Split(New String() {Environment.NewLine, vbLf}, StringSplitOptions.RemoveEmptyEntries)

                ' İlk satır başlıktır, onu atlayıp (i=1) döngüye giriyoruz
                For i = 1 To lines.Length - 1
                    Dim line = lines(i).Trim()
                    If line.StartsWith(">") Then line = line.Substring(1).Trim() ' Kendi oturumumuzun başındaki > işaretini at

                    Dim parts = line.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                    If parts.Length >= 5 Then
                        Dim user = parts(0)

                        ' quser çıktısı kayabilir (SessionName boş olabilir). ID'nin numara olduğunu biliyoruz.
                        Dim idIndex = -1
                        For j = 1 To parts.Length - 1
                            If Integer.TryParse(parts(j), Nothing) Then
                                idIndex = j
                                Exit For
                            End If
                        Next

                        If idIndex <> -1 AndAlso parts.Length > idIndex + 2 Then
                            Dim sId As Integer = Integer.Parse(parts(idIndex))
                            Dim state = parts(idIndex + 1)
                            Dim idle = parts(idIndex + 2)
                            Dim logon = String.Join(" ", parts.Skip(idIndex + 3))
                            Dim ramUsed = If(sessionRam.ContainsKey(sId), sessionRam(sId), 0)

                            result.Add(New UserSessionInfo With {
                                .Username = user,
                                .SessionId = sId,
                                .Status = state,
                                .IdleTime = idle,
                                .LogonTime = logon,
                                .RamUsageMB = ramUsed
                            })
                        End If
                    End If
                Next
            End Using
        Catch
        End Try

        ' 3. GÜVENLİK (Fallback): Eğer Windows Home kullanılıyorsa quser yoktur. Kendi oturumumuzu manuel ekleyelim.
        If result.Count = 0 Then
            Dim currentSessionId = Process.GetCurrentProcess().SessionId
            Dim ramUsed = If(sessionRam.ContainsKey(currentSessionId), sessionRam(currentSessionId), 0)
            result.Add(New UserSessionInfo With {
                .Username = Environment.UserName,
                .SessionId = currentSessionId,
                .Status = "Active",
                .IdleTime = "none",
                .LogonTime = "Bilinmiyor",
                .RamUsageMB = ramUsed
            })
        End If

        Return result.OrderByDescending(Function(x) x.RamUsageMB).ToList()
    End Function
    ' ── YENİ: ZAMANLANMIŞ GÖREVLER (schtasks) ──
    Public Function GetScheduledTasks() As List(Of ScheduledTaskInfo)
        Dim result As New List(Of ScheduledTaskInfo)
        Try
            ' /v parametresi ile "Son Çalışma Zamanı" ve "Çalışacak Komut" detaylarını çekiyoruz
            Dim pInfo As New ProcessStartInfo("schtasks", "/query /fo CSV /v") With {
                .RedirectStandardOutput = True,
                .UseShellExecute = False,
                .CreateNoWindow = True,
                .StandardOutputEncoding = System.Text.Encoding.GetEncoding(857) ' Türkçe karakter desteği
            }
            Using proc = Process.Start(pInfo)
                Dim output = proc.StandardOutput.ReadToEnd()
                Dim lines = output.Split(New String() {Environment.NewLine, vbLf}, StringSplitOptions.RemoveEmptyEntries)

                For i = 1 To lines.Length - 1
                    Dim line = lines(i).Trim()
                    If String.IsNullOrWhiteSpace(line) Then Continue For

                    ' CSV formatını (",") ayırıcıya göre böl
                    Dim parts = line.Split(New String() {""","""}, StringSplitOptions.None)
                    If parts.Length > 8 Then
                        Dim tName = parts(1).TrimStart(""""c)
                        Dim nextRun = parts(2)
                        Dim status = parts(3)
                        Dim lastRun = parts(5)
                        Dim taskToRun = parts(8)

                        ' Microsoft'un kendi sistem görevlerini gizle ki liste temiz ve okunabilir olsun
                        If Not tName.StartsWith("\Microsoft\Windows\") Then
                            result.Add(New ScheduledTaskInfo With {
                                .TaskName = tName,
                                .NextRunTime = nextRun,
                                .Status = status,
                                .LastRunTime = lastRun,
                                .Command = taskToRun
                            })
                        End If
                    End If
                Next
            End Using
        Catch
        End Try
        Return result.OrderBy(Function(x) x.TaskName).ToList()
    End Function
    ' ── YENİ: PAYLAŞILAN KLASÖRLER (Win32_Share) ──
    ' ── GÜNCELLENDİ: PAYLAŞILAN KLASÖRLER (WMI + CMD Fallback) ──
    Public Function GetSharedFolders() As List(Of SharedFolderInfo)
        Dim result As New List(Of SharedFolderInfo)

        ' 1. YÖNTEM: WMI SORGUSU (Yönetici İzni Gerektirebilir)
        Try
            Using searcher As New ManagementObjectSearcher("SELECT Name, Path, Description, Type FROM Win32_Share")
                For Each obj In searcher.Get()
                    Dim typeVal = Convert.ToUInt32(obj("Type"))
                    Dim typeStr = "Bilinmiyor"

                    Select Case typeVal
                        Case 0 : typeStr = "Disk Paylaşımı"
                        Case 1 : typeStr = "Yazıcı"
                        Case 3 : typeStr = "IPC (Sistemler Arası)"
                        Case 2147483648UI : typeStr = "Admin Paylaşımı (Gizli)"
                        Case Else : typeStr = "Özel"
                    End Select

                    result.Add(New SharedFolderInfo With {
                            .Name = If(obj("Name")?.ToString(), "-"),
                            .Path = If(obj("Path")?.ToString(), "Ağ Kaynağı"),
                            .Description = If(obj("Description")?.ToString(), "-"),
                            .ShareType = typeStr
                        })
                Next
            End Using
        Catch : End Try

        ' 2. YÖNTEM: WMI VERİ VERMEZSE "NET SHARE" KOMUTUNU PARÇALA (B Planı)
        If result.Count = 0 Then
            Try
                Dim pInfo As New ProcessStartInfo("net", "share") With {
                        .RedirectStandardOutput = True,
                        .UseShellExecute = False,
                        .CreateNoWindow = True,
                        .StandardOutputEncoding = System.Text.Encoding.GetEncoding(857) ' Türkçe karakter desteği
                    }
                Using proc = Process.Start(pInfo)
                    Dim output = proc.StandardOutput.ReadToEnd()
                    Dim lines = output.Split(New String() {Environment.NewLine, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                    Dim isDataLine = False

                    For Each line In lines
                        ' net share çıktısında başlıklar "----" çizgisinden sonra başlar
                        If line.StartsWith("----") Then
                            isDataLine = True
                            Continue For
                        End If

                        If isDataLine AndAlso Not line.Contains("başarıyla") AndAlso Not line.Contains("successfully") Then
                            ' Sabit boşlukları tek boşluğa çevir ve böl
                            Dim parts = line.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                            If parts.Length > 0 Then
                                Dim sName = parts(0)
                                ' İkinci parça genelde C:\ gibi bir dizindir, değilse IPC/Ağ kaynağıdır
                                Dim sPath = If(parts.Length > 1 AndAlso parts(1).Contains(":\"), parts(1), "Ağ Kaynağı / Sistem")
                                Dim sType = If(sName.EndsWith("$"), "Admin Paylaşımı (Gizli)", "Klasör Paylaşımı")

                                result.Add(New SharedFolderInfo With {
                                        .Name = sName,
                                        .Path = sPath,
                                        .Description = "Komut Satırından Okundu",
                                        .ShareType = sType
                                    })
                            End If
                        End If
                    Next
                End Using
            Catch : End Try
        End If

        Return result.OrderBy(Function(x) x.Name).ToList()
    End Function
    ' ── GÜNCELLENDİ: SÜRÜCÜ GÜNCELLİK ANALİZİ (1970 ve 2006 Filtreli) ──
    Public Function GetDriverUpdateReport() As List(Of DriverUpdateInfo)
        Dim result As New List(Of DriverUpdateInfo)
        Try
            Using searcher As New ManagementObjectSearcher("SELECT DeviceName, DriverVersion, DriverDate, Manufacturer FROM Win32_PnPSignedDriver")
                For Each obj In searcher.Get()
                    Try
                        Dim rawDate = obj("DriverDate")?.ToString()
                        If Not String.IsNullOrEmpty(rawDate) Then
                            ' WMI tarihini DateTime nesnesine çeviriyoruz
                            Dim dt = ManagementDateTimeConverter.ToDateTime(rawDate)

                            ' -------------------------------------------------------------
                            ' DÜZELTME: 1970 (Hatalı tarih) ve 2006 (Microsoft Jenerik) 
                            ' tarihli temel Windows sistem sürücülerini listeden atla.
                            ' -------------------------------------------------------------
                            If dt.Year < 2010 Then Continue For

                            result.Add(New DriverUpdateInfo With {
                                    .DeviceName = If(obj("DeviceName")?.ToString(), "Bilinmiyor"),
                                    .DriverVersion = If(obj("DriverVersion")?.ToString(), "-"),
                                    .DriverDate = dt,
                                    .Manufacturer = If(obj("Manufacturer")?.ToString(), "Bilinmiyor")
                                })
                        End If
                    Catch
                    End Try
                Next
            End Using
        Catch : End Try
        ' En eski sürücü en üstte olacak şekilde sırala
        Return result.OrderByDescending(Function(x) x.AgeYears).ToList()
    End Function
    ' ── YENİ: SİSTEM GERİ YÜKLEME NOKTALARI (root\default) ──
    Public Function GetRestorePoints() As List(Of RestorePointInfo)
        Dim result As New List(Of RestorePointInfo)
        Try
            ' SystemRestore sınıfı root\default altındadır
            Using searcher As New ManagementObjectSearcher("root\default", "SELECT * FROM SystemRestore")
                For Each obj In searcher.Get()
                    Try
                        Dim rawDate = obj("CreationTime")?.ToString()
                        Dim dt As DateTime = ManagementDateTimeConverter.ToDateTime(rawDate)
                        Dim typeInt = Convert.ToInt32(obj("RestorePointType"))
                        Dim typeStr = "Bilinmiyor"

                        ' WMI Tür kodlarını insan diline çeviriyoruz
                        Select Case typeInt
                            Case 0 : typeStr = "Uygulama Yükleme"
                            Case 1 : typeStr = "Uygulama Kaldırma"
                            Case 7 : typeStr = "Sistem Kontrol Noktası"
                            Case 10 : typeStr = "Cihaz Sürücüsü Yükleme"
                            Case 11 : typeStr = "Windows Update"
                            Case Else : typeStr = "Manuel Nokta"
                        End Select

                        result.Add(New RestorePointInfo With {
                            .SequenceNumber = Convert.ToUInt32(obj("SequenceNumber")),
                            .Description = If(obj("Description")?.ToString(), "İsimsiz Nokta"),
                            .CreationTime = dt,
                            .RestorePointType = typeStr
                        })
                    Catch
                    End Try
                Next
            End Using
        Catch : End Try
        Return result.OrderByDescending(Function(x) x.CreationTime).ToList()
    End Function

    Public Function CreateRestorePoint(description As String) As Boolean
        Try
            ' Geri yükleme noktası oluşturmak için Yönetici izni gerekir
            Dim scope As New ManagementScope("root\default")
            Dim sysRestore As New ManagementClass(scope, New ManagementPath("SystemRestore"), Nothing)
            Dim inParams As ManagementBaseObject = sysRestore.GetMethodParameters("CreateRestorePoint")

            inParams("Description") = description
            inParams("RestorePointType") = 0 ' 0 = APPLICATION_INSTALL (En genel tip)
            inParams("EventType") = 100 ' 100 = BEGIN_SYSTEM_CHANGE

            sysRestore.InvokeMethod("CreateRestorePoint", inParams, Nothing)
            Return True
        Catch
            Return False
        End Try
    End Function
End Class