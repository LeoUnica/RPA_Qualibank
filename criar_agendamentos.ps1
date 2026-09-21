# Cria no Agendador de Tarefas do Windows uma tarefa para CADA execucao dentro
# do horario comercial (de 10 em 10 minutos): QualiBank_SEG_0800, ... QualiBank_SAB_1150.
# Segunda a sexta: 08:00 as 18:00. Sabado: 08:00 as 12:00.
# Todas ficam na pasta \QualiBank do Agendador. Rodam com o usuario atual, so
# com ele logado (o Outlook desktop e o navegador precisam da sessao interativa).
#
# Uso:    powershell -ExecutionPolicy Bypass -File criar_agendamentos.ps1
# Remover: powershell -ExecutionPolicy Bypass -File criar_agendamentos.ps1 -Remover

param([switch]$Remover)

$pasta = "\QualiBank\"
$projeto = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = (& python -c "import sys;print(sys.executable)").Trim()

if ($Remover) {
    Get-ScheduledTask -TaskPath $pasta -ErrorAction SilentlyContinue |
        Unregister-ScheduledTask -Confirm:$false
    Write-Host "Tarefas removidas."
    return
}

$acao = New-ScheduledTaskAction -Execute $python -Argument "propostas.py" -WorkingDirectory $projeto
$config = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited

# Dias uteis (seg-sex): 08:00 as 18:00 | Sabado: 08:00 as 12:00
$diasSemana = @(
    @{ Sigla = "SEG"; Dia = "Monday";    HoraIni = 8; HoraFim = 18 },
    @{ Sigla = "TER"; Dia = "Tuesday";   HoraIni = 8; HoraFim = 18 },
    @{ Sigla = "QUA"; Dia = "Wednesday"; HoraIni = 8; HoraFim = 18 },
    @{ Sigla = "QUI"; Dia = "Thursday";  HoraIni = 8; HoraFim = 18 },
    @{ Sigla = "SEX"; Dia = "Friday";    HoraIni = 8; HoraFim = 18 },
    @{ Sigla = "SAB"; Dia = "Saturday";  HoraIni = 8; HoraFim = 12 }
)

$criadas = 0
foreach ($d in $diasSemana) {
    foreach ($h in $d.HoraIni..($d.HoraFim - 1)) {
        foreach ($m in 0, 10, 20, 30, 40, 50) {
            $hora = "{0:D2}{1:D2}" -f $h, $m
            $gatilho = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $d.Dia -At ([datetime]::Today.AddHours($h).AddMinutes($m))
            Register-ScheduledTask -TaskName "QualiBank_$($d.Sigla)_$hora" -TaskPath $pasta -Action $acao `
                -Trigger $gatilho -Settings $config -Principal $principal -Force | Out-Null
            $criadas++
        }
    }
}
Write-Host "$criadas tarefas criadas em $pasta"
