# Crée les tâches planifiées Windows (du lundi au vendredi, heure de Paris) :
#  - 18h15 : régénère le site après la clôture d'Euronext Paris
#  - 22h30 : régénère après la clôture de Wall Street (favoris américains) 
# Pour les supprimer :
#   Unregister-ScheduledTask -TaskName 'Compte-rendu Bourse' -Confirm:$false
#   Unregister-ScheduledTask -TaskName 'Compte-rendu Bourse (22h30)' -Confirm:$false
$script  = Join-Path $PSScriptRoot 'generer.ps1'
$jours   = 'Monday','Tuesday','Wednesday','Thursday','Friday'
$options = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 20)

$action = New-ScheduledTaskAction -Execute 'powershell.exe' -WorkingDirectory $PSScriptRoot `
          -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`""
Register-ScheduledTask -TaskName 'Compte-rendu Bourse' -Action $action -Settings $options -Force `
  -Trigger (New-ScheduledTaskTrigger -Weekly -DaysOfWeek $jours -At '18:15') `
  -Description 'Génère le compte-rendu de clôture des actions' | Out-Null

$actionMail = New-ScheduledTaskAction -Execute 'powershell.exe' -WorkingDirectory $PSScriptRoot `
              -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`""
Register-ScheduledTask -TaskName 'Compte-rendu Bourse (22h30)' -Action $actionMail -Settings $options -Force `
  -Trigger (New-ScheduledTaskTrigger -Weekly -DaysOfWeek $jours -At '22:30') `
  -Description 'Met à jour le compte-rendu après la clôture de Wall Street' | Out-Null

Write-Host "Tâches planifiées créées : 18h15 (site) et 22h30 (site), du lundi au vendredi."
