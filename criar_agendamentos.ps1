# Cria no Agendador de Tarefas do Windows uma tarefa para CADA execucao do dia
# (de 10 em 10 minutos, 24h): QualiBank_0000, QualiBank_0010, ... QualiBank_2350.
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

$criadas = 0
foreach ($h in 0..23) {
    foreach ($m in 0, 10, 20, 30, 40, 50) {
        $hora = "{0:D2}{1:D2}" -f $h, $m
        $gatilho = New-ScheduledTaskTrigger -Daily -At ([datetime]::Today.AddHours($h).AddMinutes($m))
        Register-ScheduledTask -TaskName "QualiBank_$hora" -TaskPath $pasta -Action $acao `
            -Trigger $gatilho -Settings $config -Principal $principal -Force | Out-Null
        $criadas++
    }
}
Write-Host "$criadas tarefas criadas em $pasta"
