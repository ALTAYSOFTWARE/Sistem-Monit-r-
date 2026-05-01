

' ── MEVCUT MODELLER (GENİŞLETİLMİŞ) ──

Public Class SysProcessInfo
        Public Property Name As String = ""
        Public Property PID As Integer
        Public Property MemoryMB As Double
        Public Property Threads As Integer
        Public Property CpuPct As Double = 0
        Public ReadOnly Property MemoryMBStr As String
            Get
                Return $"{MemoryMB:F1} MB"
            End Get
        End Property
        Public ReadOnly Property CpuPctStr As String
            Get
                Return $"{CpuPct:F1}%"
            End Get
        End Property
    End Class

    Public Class DiskDriveInfo
        Public Property Letter As String = ""
        Public Property Label As String = ""
        Public Property TotalGB As Double
        Public Property FreeGB As Double
        Public ReadOnly Property UsedGB As Double
            Get
                Return TotalGB - FreeGB
            End Get
        End Property
        Public ReadOnly Property UsedPct As Double
            Get
                If TotalGB <= 0 Then Return 0
                Return Math.Round((UsedGB / TotalGB) * 100, 1)
            End Get
        End Property
        Public ReadOnly Property DisplayName As String
            Get
                If String.IsNullOrEmpty(Label) Then Return Letter
                Return $"{Letter} ({Label})"
            End Get
        End Property
    End Class

    Public Class DiskSpaceItem
        Public Property FolderName As String = ""
        Public Property SizeBytes As Long = 0
        Public ReadOnly Property SizeStr As String
            Get
                If SizeBytes > 1073741824 Then Return $"{(SizeBytes / 1073741824.0):F2} GB"
                Return $"{(SizeBytes / 1048576.0):F1} MB"
            End Get
        End Property
    End Class

    Public Class NetConnection
        Public Property ProcessName As String = ""
        Public Property Protocol As String = ""
        Public Property LocalAddress As String = ""
        Public Property RemoteAddress As String = ""
        Public Property State As String = ""
    End Class

' GÜNCELLENDİ: BAŞLANGIÇ UYGULAMASI (DİJİTAL İMZA DESTEĞİ İLE)
Public Class StartupApp
    Public Property Name As String = ""
    Public Property Command As String = ""
    Public Property RegistryLocation As String = ""
    Public Property Publisher As String = "Bilinmiyor"
    Public Property IsSigned As Boolean = False

    Public ReadOnly Property SignatureStatus As String
        Get
            Return If(IsSigned, "✅ Güvenilir", "⚠️ İmzasız")
        End Get
    End Property

    Public ReadOnly Property StatusColor As String
        Get
            Return If(IsSigned, "#6BCB77", "#FFD93D") ' İmzalıysa Yeşil, Değilse Sarı/Turuncu
        End Get
    End Property
End Class

Public Class DiskHealth
        Public Property Model As String = ""
        Public Property Status As String = ""
        Public ReadOnly Property IsHealthy As Boolean
            Get
                Return Status.Trim().ToUpper() = "OK"
            End Get
        End Property
        Public ReadOnly Property StatusColor As String
            Get
                Return If(IsHealthy, "#6BCB77", "#FF6B6B")
            End Get
        End Property
    End Class

    Public Class CrashLogEntry
        Public Property Time As String = ""
        Public Property Source As String = ""
        Public Property Message As String = ""
        Public Property Level As String = ""
        Public ReadOnly Property LevelColor As String
            Get
                Return If(Level.ToLower().Contains("error"), "#FF4060", "#FFD93D")
            End Get
        End Property
    End Class

    Public Class BatteryInfo
        Public Property HasBattery As Boolean = False
        Public Property ChargePercent As Integer = 0
        Public Property Status As String = ""
    End Class

    Public Class DeepHwInfo
        Public Property Motherboard As String = "Bilinmiyor"
        Public Property BIOS As String = "Bilinmiyor"
        Public Property RamSpeed As String = "Bilinmiyor"
    End Class

    Public Class AlarmConfig
        Public Property CpuEnabled As Boolean = True
        Public Property CpuThreshold As Integer = 90
        Public Property RamEnabled As Boolean = True
        Public Property RamThreshold As Integer = 85
        Public Property DiskEnabled As Boolean = True
        Public Property DiskThreshold As Integer = 90
        Public Property ThemeColor As String = "#00D4FF"
        Public Property RunAtStartup As Boolean = False
        Public Property MinimizeToTray As Boolean = True
        Public Property AlarmSoundEnabled As Boolean = True
        Public Property SnoozeDurationSec As Integer = 15
    End Class

    Public Class AlarmHistoryEntry
        Public Property Time As String
        Public Property Type As String
        Public Property Message As String
    End Class

    ' ── YENİ MODELLER ──

    Public Class CpuCoreInfo
        Public Property CoreIndex As Integer
        Public Property UsagePct As Double
        Public ReadOnly Property Label As String
            Get
                Return $"Ç{CoreIndex}"
            End Get
        End Property
        Public ReadOnly Property UsagePctStr As String
            Get
                Return $"{UsagePct:F0}%"
            End Get
        End Property
    End Class

    Public Class InstalledProgram
        Public Property Name As String = ""
        Public Property Publisher As String = ""
        Public Property Version As String = ""
        Public Property InstallDate As String = ""
    End Class

    Public Class ServiceInfo
        Public Property ServiceName As String = ""
        Public Property DisplayName As String = ""
        Public Property Status As String = ""
        Public Property StartType As String = ""
        Public ReadOnly Property StatusColor As String
            Get
                Select Case Status.ToLower()
                    Case "running" : Return "#6BCB77"
                    Case "stopped" : Return "#FF6B6B"
                    Case Else : Return "#FFD93D"
                End Select
            End Get
        End Property
        Public ReadOnly Property StatusTR As String
            Get
                Select Case Status.ToLower()
                    Case "running" : Return "Çalışıyor"
                    Case "stopped" : Return "Durduruldu"
                    Case "paused" : Return "Duraklatıldı"
                    Case Else : Return Status
                End Select
            End Get
        End Property
    End Class

    Public Class PowerPlanInfo
        Public Property Name As String = ""
        Public Property Guid As String = ""
        Public Property IsActive As Boolean = False
        Public ReadOnly Property ActiveStr As String
            Get
                Return If(IsActive, "✅ Aktif", "")
            End Get
        End Property
    End Class

    Public Class NetworkAdapterInfo
        Public Property AdapterName As String = ""
        Public Property AdapterType As String = ""
        Public Property IpAddress As String = ""
        Public Property MacAddress As String = ""
        Public Property SpeedMbps As String = ""
    End Class

    ' YENİ: FIREWALL KURALI MODELİ
    Public Class FirewallRule
        Public Property RuleName As String = ""
        Public Property Direction As String = ""  ' Gelen (In) / Giden (Out)
        Public Property Action As String = ""     ' İzin Ver (Allow) / Engelle (Block)
        Public Property Profile As String = ""
        Public ReadOnly Property ActionColor As String
            Get
                Return If(Action.ToLower().Contains("allow") OrElse Action.ToLower().Contains("izin"), "#6BCB77", "#FF6B6B")
            End Get
        End Property
    End Class
' YENİ: SÜRÜCÜ (DRIVER) BİLGİSİ
Public Class DriverInfo
    Public Property Name As String = ""
    Public Property DisplayName As String = ""
    Public Property State As String = ""
    Public Property Status As String = ""

    ' Status "OK" değilse (Örn: Error, Degraded) Kırmızı yap
    Public ReadOnly Property StatusColor As String
        Get
            Return If(Status.Trim().ToUpper() = "OK", "#6BCB77", "#FF4060")
        End Get
    End Property

    ' Çalışıyorsa Yeşil, Durmuşsa Sarımsı
    Public ReadOnly Property StateColor As String
        Get
            Return If(State.Trim().ToLower() = "running", "#C0C0E0", "#FFD93D")
        End Get
    End Property
End Class
' YENİ: OTURUM AÇMIŞ KULLANICI BİLGİSİ
Public Class UserSessionInfo
    Public Property Username As String = ""
    Public Property SessionId As Integer = 0
    Public Property Status As String = ""
    Public Property IdleTime As String = ""
    Public Property LogonTime As String = ""
    Public Property RamUsageMB As Double = 0
    Public ReadOnly Property RamUsageStr As String
        Get
            Return $"{RamUsageMB:F1} MB"
        End Get
    End Property
    Public ReadOnly Property StatusColor As String
        Get
            Return If(Status.Trim().ToLower().Contains("active") OrElse Status.Trim().ToLower().Contains("aktif"), "#6BCB77", "#FFD93D")
        End Get
    End Property
End Class
' YENİ: ZAMANLANMIŞ GÖREV MODELİ
Public Class ScheduledTaskInfo
    Public Property TaskName As String = ""
    Public Property NextRunTime As String = ""
    Public Property Status As String = ""
    Public Property LastRunTime As String = ""
    Public Property Command As String = ""

    Public ReadOnly Property StatusColor As String
        Get
            ' Görev "Ready" veya "Hazır" ise yeşil, diğer durumlarda gri/sarımsı yap
            Return If(Status.Trim().ToLower().Contains("ready") OrElse Status.Trim().ToLower().Contains("hazır"), "#6BCB77", "#FFD93D")
        End Get
    End Property
End Class
' YENİ: GÜVENLİK DURUM MODELİ
Public Class SecurityStatusItem
    Public Property FeatureName As String = ""
    Public Property StatusText As String = ""
    Public Property IsSecure As Boolean = False
    Public ReadOnly Property StatusColor As String
        Get
            Return If(IsSecure, "#6BCB77", "#FF6B6B") ' Yeşil veya Kırmızı
        End Get
    End Property
End Class
' YENİ: PAYLAŞILAN KLASÖR MODELİ
Public Class SharedFolderInfo
    Public Property Name As String = ""
    Public Property Path As String = ""
    Public Property Description As String = ""
    Public Property ShareType As String = ""

    ' C$, ADMIN$ gibi gizli paylaşımları ayırt etmek için
    Public ReadOnly Property IsHidden As Boolean
        Get
            Return Name.EndsWith("$")
        End Get
    End Property

    Public ReadOnly Property StatusColor As String
        Get
            Return If(IsHidden, "#FFD93D", "#6BCB77") ' Gizli: Sarı, Açık: Yeşil
        End Get
    End Property
End Class
' YENİ: SÜRÜCÜ GÜNCELLEME ANALİZ MODELİ
Public Class DriverUpdateInfo
    Public Property DeviceName As String = ""
    Public Property DriverVersion As String = ""
    Public Property DriverDate As DateTime
    Public Property Manufacturer As String = ""

    Public ReadOnly Property DateStr As String
        Get
            Return DriverDate.ToString("dd.MM.yyyy")
        End Get
    End Property

    Public ReadOnly Property AgeYears As Double
        Get
            Return Math.Round((DateTime.Now - DriverDate).TotalDays / 365.25, 1)
        End Get
    End Property

    Public ReadOnly Property StatusColor As String
        Get
            ' 2 yıldan eskiyse sarı/turuncu, 5 yıldan eskiyse kırmızı
            If AgeYears > 5 Then Return "#FF4060"
            If AgeYears > 2 Then Return "#FFD93D"
            Return "#6BCB77" ' Güncel
        End Get
    End Property
End Class
' YENİ: SİSTEM GERİ YÜKLEME NOKTASI MODELİ
Public Class RestorePointInfo
    Public Property SequenceNumber As UInteger
    Public Property Description As String = ""
    Public Property CreationTime As DateTime
    Public Property RestorePointType As String = ""

    Public ReadOnly Property DateStr As String
        Get
            Return CreationTime.ToString("dd.MM.yyyy HH:mm")
        End Get
    End Property
End Class