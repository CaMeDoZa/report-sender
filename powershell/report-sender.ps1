# ============================================================================
# Скрипт автоматической рассылки отчетов
# SMTP: mail.ru
# Дата: 29.11.2025
# Автор: Дан (telegram: @camedoza, e-mail: camedoza.dan@yandex.ru)
# ============================================================================

#region Конфигурация
$Config = @{
    # SMTP настройки Mail.ru
    SmtpServer = "smtp.mail.ru"
    Port = 587
    Username = "*E-MAIL*"
    Password = "*PASSWORD*"
    FromEmail = "*FROM EMAIL*"
    FromName = "Система отчётов"
    
    # Директории
    LogPath = "d:\Otchety\Logs"
    AnalyticPath = "d:\Otchety\Analitika"
    DZPath = "d:\Otchety\DZ"
    
    # Настройки надежности
    MaxFileWaitMinutes = 20      # Ожидание файлов отправки до 20 минут
    RetryAttempts = 5            # 5 попыток отправки
    RetryDelaySeconds = 45       # 45 секунд между попытками
    SendDelaySeconds = 3         # Пауза между письмами
    SmtpTimeout = 60000          # Таймаут SMTP 60 секунд
}

# Список получателей
$Recipients = @(
    @{
        Region = "TP1"
        To = "*TO*"
        CC = @("*CC TO*")
        FilePrefix = "*FILE PREFIX*"
    }
    #@{
    #    ...
    #}
)
#endregion

#region Инициализация логирования
# Создание папки для логов если не существует
if (-not (Test-Path $Config.LogPath)) {
    New-Item -ItemType Directory -Path $Config.LogPath -Force | Out-Null
}

$DateStamp = Get-Date -Format "yyyy-MM-dd"
$LogFile = Join-Path $Config.LogPath "mail_$DateStamp.log"
$ErrorLogFile = Join-Path $Config.LogPath "errors_$DateStamp.log"
$SummaryFile = Join-Path $Config.LogPath "summary_$DateStamp.txt"

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR', 'CRITICAL')]
        [string]$Level = 'INFO'
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    # Цветной вывод в консоль
    $Color = switch ($Level) {
        'SUCCESS' { 'Green' }
        'WARNING' { 'Yellow' }
        'ERROR' { 'Red' }
        'CRITICAL' { 'Magenta' }
        default { 'White' }
    }
    Write-Host $LogEntry -ForegroundColor $Color
    
    # Запись в основной лог
    Add-Content -Path $LogFile -Value $LogEntry -Encoding UTF8
    
    # Запись в лог ошибок
    if ($Level -in @('ERROR', 'CRITICAL')) {
        Add-Content -Path $ErrorLogFile -Value $LogEntry -Encoding UTF8
    }
}
#endregion

#region Функции
function Test-FileExists {
    param(
        [string]$FilePath,
        [int]$TimeoutMinutes
    )
    
    $Timeout = [DateTime]::Now.AddMinutes($TimeoutMinutes)
    $CheckInterval = 30 # секунд
    
    while ([DateTime]::Now -lt $Timeout) {
        if (Test-Path $FilePath) {
            $FileInfo = Get-Item $FilePath
            # Проверка что файл не пустой и не заблокирован
            if ($FileInfo.Length -gt 0) {
                try {
                    $Stream = [System.IO.File]::Open($FilePath, 'Open', 'Read', 'None')
                    $Stream.Close()
                    Write-Log "Файл готов: $FilePath (размер: $([math]::Round($FileInfo.Length/1KB, 2)) KB)" "SUCCESS"
                    return $true
                } catch {
                    Write-Log "Файл заблокирован, ожидание: $FilePath" "WARNING"
                }
            } else {
                Write-Log "Файл пустой, ожидание: $FilePath" "WARNING"
            }
        }
        
        Start-Sleep -Seconds $CheckInterval
    }
    
    Write-Log "TIMEOUT: Файл не готов после $TimeoutMinutes минут: $FilePath" "ERROR"
    return $false
}

function Send-MailRuEmail {
    param(
        [string]$To,
        [string[]]$CC,
        [string]$Subject,
        [string]$AnalyticFile,
        [string]$DZFile,
        [string]$RegionName
    )
    
    for ($Attempt = 1; $Attempt -le $Config.RetryAttempts; $Attempt++) {
        $smtp = $null
        $message = $null
        $att1 = $null
        $att2 = $null
        
        try {
            Write-Log "[$RegionName] Попытка $Attempt из $($Config.RetryAttempts)" "INFO"
            
            # Проверка файлов
            if (-not (Test-Path $AnalyticFile)) {
                throw "Файл аналитики не найден: $AnalyticFile"
            }
            if (-not (Test-Path $DZFile)) {
                throw "Файл ДЗ не найден: $DZFile"
            }
            
            # Создание сообщения
            $message = New-Object System.Net.Mail.MailMessage
            $message.From = New-Object System.Net.Mail.MailAddress($Config.FromEmail, $Config.FromName, [System.Text.Encoding]::UTF8)
            $message.To.Add($To)
            
            # Добавление копий
            foreach ($CopyAddr in $CC) {
                if ($CopyAddr) {
                    $message.CC.Add($CopyAddr)
                }
            }
            
            $message.Subject = $Subject
            $message.SubjectEncoding = [System.Text.Encoding]::UTF8
            $message.IsBodyHtml = $true
            $message.BodyEncoding = [System.Text.Encoding]::UTF8
            
            # Тело письма
            $message.Body = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; }
        .header { background: linear-gradient(135deg, #c73c3c 0%, #a02a2a 100%); color: white; padding: 20px; border-radius: 5px; }
        .content { padding: 20px; background: #f9f9f9; border-radius: 5px; margin-top: 20px; }
        .footer { margin-top: 20px; padding: 15px; background: #e9ecef; border-radius: 5px; font-size: 12px; color: #666; }
        .highlight { color: #667eea; font-weight: bold; }
    </style>
</head>
<body>
    <div class="header">
        <h2>🍵 Автоматическая рассылка отчетов</h2>
    </div>
    <div class="content">
        <p>Добрый день!</p>
        <p>Во вложении находятся ежедневные отчеты:</p>
        <ul>
            <li>📊 <span class="highlight">Аналитика продаж</span></li>
            <li>💰 <span class="highlight">Дебиторская задолженность</span></li>
        </ul>
        <p>Регион: <strong>$RegionName</strong></p>
        <p>Дата формирования: <strong>$(Get-Date -Format 'dd.MM.yyyy HH:mm')</strong></p>
    </div>
    <div class="footer">
        <p>Это автоматическое письмо. Не отвечайте на него.</p>
        <p>По вопросам работы системы обращайтесь на e-mail: camedoza.dan@yandex.ru</p>
    </div>
</body>
</html>
"@
            
            # Прикрепление файлов
            $att1 = New-Object System.Net.Mail.Attachment($AnalyticFile)
            $att2 = New-Object System.Net.Mail.Attachment($DZFile)
            $message.Attachments.Add($att1)
            $message.Attachments.Add($att2)
            
            # Настройка SMTP для Mail.ru (SSL на порту 465)
            $smtp = New-Object System.Net.Mail.SmtpClient($Config.SmtpServer, $Config.Port)
            $smtp.EnableSsl = $true
            $smtp.Timeout = $Config.SmtpTimeout
            $smtp.Credentials = New-Object System.Net.NetworkCredential($Config.Username, $Config.Password)
            
            # Отправка
            $smtp.Send($message)
            
            Write-Log "[$RegionName] ✓ УСПЕШНО отправлено на $To" "SUCCESS"
            Write-Log "[$RegionName] Файлы: $(Split-Path $AnalyticFile -Leaf), $(Split-Path $DZFile -Leaf)" "INFO"
            
            return @{
                Success = $true
                Attempt = $Attempt
                Error = $null
            }
            
        } catch {
            $ErrorMsg = $_.Exception.Message
            Write-Log "[$RegionName] ✗ Ошибка на попытке $Attempt : $ErrorMsg" "ERROR"
            
            if ($Attempt -lt $Config.RetryAttempts) {
                Write-Log "[$RegionName] Повтор через $($Config.RetryDelaySeconds) сек..." "WARNING"
                Start-Sleep -Seconds $Config.RetryDelaySeconds
            } else {
                Write-Log "[$RegionName] ✗✗✗ КРИТИЧНО: Все попытки исчерпаны ✗✗✗" "CRITICAL"
                return @{
                    Success = $false
                    Attempt = $Attempt
                    Error = $ErrorMsg
                }
            }
            
        } finally {
            # Освобождение ресурсов
            if ($att1) { $att1.Dispose() }
            if ($att2) { $att2.Dispose() }
            if ($message) { $message.Dispose() }
            if ($smtp) { $smtp.Dispose() }
        }
    }
    
    return @{
        Success = $false
        Attempt = $Config.RetryAttempts
        Error = "Неизвестная ошибка"
    }
}
#endregion

#region Основной блок выполнения
try {
    Write-Log "═══════════════════════════════════════════════════════════════" "INFO"
    Write-Log "Запуск скрипта отправки отчётов" "INFO"
    Write-Log "═══════════════════════════════════════════════════════════════" "INFO"
    Write-Log "SMTP: $($Config.SmtpServer):$($Config.Port)" "INFO"
    Write-Log "От кого: $($Config.FromEmail)" "INFO"
    Write-Log "Регионов для обработки: $($Recipients.Count)" "INFO"
    Write-Log "═══════════════════════════════════════════════════════════════" "INFO"
    
    $StartTime = Get-Date
    $Statistics = @{
        Total = $Recipients.Count
        Success = 0
        Failed = 0
        FailedRegions = @()
    }
    
    $DateFormat = Get-Date -Format "dd.MM.yyyy"
    
    foreach ($Recipient in $Recipients) {
        Write-Log " " "INFO"
        Write-Log "───────────────────────────────────────────────────────────────" "INFO"
        Write-Log "Обработка региона: $($Recipient.Region)" "INFO"
        Write-Log "───────────────────────────────────────────────────────────────" "INFO"
        
        # Формирование путей к файлам
        $AnalyticFile = Join-Path $Config.AnalyticPath "Analitika_$($Recipient.FilePrefix)_$DateFormat.xlsx"
        $DZFile = Join-Path $Config.DZPath "DZ_$($Recipient.FilePrefix)_$DateFormat.xlsx"
        
        Write-Log "Ожидание файлов для $($Recipient.Region)..." "INFO"
        
        # Проверка наличия файлов с ожиданием
        $AnalyticReady = Test-FileExists -FilePath $AnalyticFile -TimeoutMinutes $Config.MaxFileWaitMinutes
        $DZReady = Test-FileExists -FilePath $DZFile -TimeoutMinutes $Config.MaxFileWaitMinutes
        
        if (-not $AnalyticReady -or -not $DZReady) {
            Write-Log "[$($Recipient.Region)] Пропуск: Файлы не готовы" "ERROR"
            $Statistics.Failed++
            $Statistics.FailedRegions += $Recipient.Region
            continue
        }
        
        # Отправка письма
        $Result = Send-MailRuEmail `
            -To $Recipient.To `
            -CC $Recipient.CC `
            -Subject "Отчёты аналитики и ДЗ $($Recipient.Region)" `
            -AnalyticFile $AnalyticFile `
            -DZFile $DZFile `
            -RegionName $Recipient.Region
        
        if ($Result.Success) {
            $Statistics.Success++
        } else {
            $Statistics.Failed++
            $Statistics.FailedRegions += "$($Recipient.Region) (Причина: $($Result.Error))"
        }
        
        # Пауза между отправками (на всякий)
        if ($Recipient -ne $Recipients[-1]) {
            Write-Log "Пауза $($Config.SendDelaySeconds) сек. перед следующей отправкой..." "INFO"
            Start-Sleep -Seconds $Config.SendDelaySeconds
        }
    }
    
    # Итоговая статистика
    $EndTime = Get-Date
    $Duration = $EndTime - $StartTime
    
    Write-Log " " "INFO"
    Write-Log "═══════════════════════════════════════════════════════════════" "INFO"
    Write-Log "ЗАВЕРШЕНИЕ: Рассылка завершена" "INFO"
    Write-Log "═══════════════════════════════════════════════════════════════" "INFO"
    Write-Log "Время выполнения: $($Duration.Minutes) мин $($Duration.Seconds) сек" "INFO"
    Write-Log "Всего регионов: $($Statistics.Total)" "INFO"
    Write-Log "Успешно отправлено: $($Statistics.Success)" "SUCCESS"
    Write-Log "Ошибок: $($Statistics.Failed)" $(if ($Statistics.Failed -gt 0) { "ERROR" } else { "INFO" })
    
    if ($Statistics.Failed -gt 0) {
        Write-Log "Проблемные регионы:" "ERROR"
        foreach ($FailedRegion in $Statistics.FailedRegions) {
            Write-Log "  - $FailedRegion" "ERROR"
        }
    }
    Write-Log "═══════════════════════════════════════════════════════════════" "INFO"
    
    # Сохранение сводки в файл
    $Summary = @"
═══════════════════════════════════════════════════════════════
ОТЧЕТ О РАССЫЛКЕ
Дата: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')
═══════════════════════════════════════════════════════════════

Время выполнения: $($Duration.Minutes) мин $($Duration.Seconds) сек
Всего регионов: $($Statistics.Total)
Успешно: $($Statistics.Success)
Ошибок: $($Statistics.Failed)
Процент успеха: $([math]::Round(($Statistics.Success / $Statistics.Total) * 100, 2))%

$(if ($Statistics.Failed -gt 0) {
"ПРОБЛЕМНЫЕ РЕГИОНЫ:
$(($Statistics.FailedRegions | ForEach-Object { "  - $_" }) -join "`n")
"
} else {
"✓ Все регионы обработаны успешно!"
})

═══════════════════════════════════════════════════════════════
Лог файлы:
  - Основной: $LogFile
  - Ошибки: $ErrorLogFile
═══════════════════════════════════════════════════════════════
"@
    
    Set-Content -Path $SummaryFile -Value $Summary -Encoding UTF8
    
    # Код завершения
    if ($Statistics.Failed -gt 0) {
        exit 1
    } else {
        exit 0
    }
    
} catch {
    Write-Log "КРИТИЧЕСКАЯ ОШИБКА СКРИПТА: $_" "CRITICAL"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "CRITICAL"
    exit 2
}
#endregion