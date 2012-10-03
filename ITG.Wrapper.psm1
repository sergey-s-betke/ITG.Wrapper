# включаем обработку и регистрацию ошибок в журнале

[System.Reflection.Assembly]::LoadWithPartialName('System.Diagnostics') 
[System.Reflection.Assembly]::LoadWithPartialName('System.Management.Automation.PsEngineEvent') 

[bool]$script:isExitedITGScript = $false;
[string]$global:scriptActivity;
[int]$step = 0 # текущий этап в сценарии, используется для прогресс-баров
[int]$steps = 0 # общее количество этапов в сценарии, используется для прогресс-баров
[int]$progressBarId = 10

function Enter-ITGScript { 
	<#
		.Synopsis
            Инициализация "обёртки", фиксация факта начала работы сценария в журнале событий.
		.Description
            Инициализация "обёртки", фиксация факта начала работы сценария в журнале событий.
		.Parameter activity
		    Текстовое пояснение назначения сценария. Выводим в журнал событий и в прогресс-бары.
		.Example
			Enter-ITGScript `
                -activity 'Анализ журналов SMTP сервера' 
	#>
  
    param (
    	[Parameter(
            Mandatory=$true,
            Position=0,
            ValueFromPipeline=$false,
            HelpMessage="Текстовое пояснение назначения сценария. Выводим в журнал событий и в прогресс-бары."
		)]
        [string]$activity
    )

    $global:scriptActivity = $activity;
    $script:isExitedITGScript = $false;
   
    Write-EventLogITG -message 'Скрипт успешно приступил к работе...';

    write-progress `
        -id $progressBarId `
        -activity $scriptActivity `
        -currentOperation 'Инициализация сценария' `
        -status 'Инициализация...' `
        -percentcomplete 0
}

function Exit-ITGScript { 
	<#
		.Synopsis
            Финализация "обёртки", фиксация факта завершения работы сценария в журнале событий.
		.Description
            Финализация "обёртки", фиксация факта завершения работы сценария в журнале событий.
            Допускается явный вызов. Если явного вызова нет в коде сценария, вызвана будет как
            реакция на событие [System.Management.Automation.PsEngineEvent]::Exiting.
		.Example
			Exit-ITGScript
	#>

    if (-not $script:isExitedITGScript) {
        write-progress `
            -id $progressBarId `
            -activity $scriptActivity `
            -status "Завершение работы..." `
            -completed

        Write-Successfull 
    };
    $script:isExitedITGScript = $true;
}

# регистрируем обработчик на завершение сессии / сценария с тем, чтобы без явного вызова
# обработчиков этого модуля писать необходимые записи в журналы событий и выполнять 
# прочие действия при завершении сценария
Register-EngineEvent `
    -sourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) `
    -supportEvent `
    -action {
        Exit-ITGScript;
    };

<#
write-progress `
    -id $progressBarId `
    -activity $script:scriptActivity `
    -currentOperation "Готовим отчёт по всем проблемным исходящим SMTP сессиям ($step из $steps)" `
    -status "Анализ журнала SMTP сервера, сбор сессий" `
    -percentcomplete ($step/$steps*100)
#>

function Write-EventLogITG { 
	<#
		.Synopsis
            Регистрирует событие в журнале событий, при этом корректно регистрирует источник событий.
		.Description
            Регистрирует событие в журнале событий, при этом корректно регистрирует источник событий.
		.Parameter message
		    собственно текст сообщения.
		.Parameter eventLog
			Идентификатор журнала для регистрации событий. При необходимости - будет создан.
		.Parameter EntryType
			Тип сообщения
		.Parameter eventId
			Идентификатор сообщения
		.Parameter source
			Источник сообщения. В случае отсутствия - имя файла сценария.
		.Example
			Write-EventLog `
                -msg "такое вот событие" `
                -type [System.Diagnostics.EventLogEntryType]::Error `
	#>
  
    param (
    	[Parameter(
            Mandatory=$true,
            Position=0,
            ValueFromPipeline=$true,
            HelpMessage="Текстовое сообщение."
		)]
        [string]$message,

		[Parameter(
			Mandatory=$false,
			Position=1,
			ValueFromPipeline=$false,
			HelpMessage="Журнал, в который планируем регистрировать событие."
		)]
        [string]$eventLog = 'Application',

		[Parameter(
			Mandatory=$false,
			Position=2,
			ValueFromPipeline=$false,
			HelpMessage="Тип сообщения."
		)]
        [System.Diagnostics.EventLogEntryType]$EntryType = [System.Diagnostics.EventLogEntryType]::Information,

		[Parameter(
			Mandatory=$false,
			Position=3,
			ValueFromPipeline=$false,
			HelpMessage="Идентификатор сообщения."
		)]
        [int]$eventId = 0,

		[Parameter(
			Mandatory=$false,
			Position=4,
			ValueFromPipeline=$false,
			HelpMessage="Источник сообщения. По умолчанию - имя файла сценария."
		)]
        [string]$source = $([System.IO.Path]::GetFileNameWithoutExtension($global:myinvocation.mycommand.name))
    )

    #Create the source, if it does not already exist.
    if (![System.Diagnostics.EventLog]::SourceExists($source)) {
        [System.Diagnostics.EventLog]::CreateEventSource($source, $eventLog)
    } else {
        $eventLog = [System.Diagnostics.EventLog]::LogNameFromSourceName($source,'.')
    }
    #Check if Event Type is correct
    $log = New-Object System.Diagnostics.EventLog($eventLog)
    $log.Source=$source
    if ($message.length -ge 32767) {$message = $message.Substring(1,32766)};
    $log.WriteEntry(`
        $message,`
        $EntryType,`
        $eventid
    )
}
 
function Write-CurrentException { 
	<#
		.Synopsis
            Регистрирует в журнале событий текущую ошибку (исключение).
		.Description
            Регистрирует в журнале событий текущую ошибку (исключение).
		.Parameter exception
			Объект - описатель ошибки ($_ в trap).
		.Example
            trap {
                Write-CurrentException
            }
	#>
  
    param (
		[Parameter(
			Mandatory=$true,
			Position=0,
			ValueFromPipeline=$false,
			HelpMessage='Объект - описатель ошибки ($_ в trap).'
		)]
        $exception
    )

	$errorDescription = @"
При выполнении сценария $($global:myinvocation.mycommand.name) $((Get-Date).ToShortDateString()) в $((Get-Date).ToShortTimeString()) возникла ошибка:
$($exception.InvocationInfo.PositionMessage)
$($exception.Exception)

$($exception.Exception.Data["extraInfo"])
"@
    Write-EventLogITG `
        -message $errorDescription `
        -entryType Error `
    ;
	write-error $errorDescription;
}

function Write-Successfull { 
	<#
		.Synopsis
            Регистрирует в журнале событий информацию о успешном завершении сценария.
		.Description
            Регистрирует в журнале событий информацию о успешном завершении сценария.
		.Example
            Write-Successfull
	#>
  
    param (
    )

    Write-EventLogITG `
        -message "Скрипт успешно завершил работу." `
        -entryType Information
}

$ErrorActionPreference = 'Stop'
trap {
    Write-CurrentException($_);
    break;
}

Export-ModuleMember `
    Enter-ITGScript, `
    Exit-ITGScript, `
    Write-EventLogITG, `
    Write-CurrentException, `
    Write-Successfull
