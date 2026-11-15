# Configure l'envoi du compte-rendu du soir par e-mail (à lancer une seule fois).
# Le mot de passe est chiffré avec votre session Windows : il n'est lisible que par vous, sur ce PC.
$ErrorActionPreference = 'Stop'
$defaut = ''

Write-Host ""
Write-Host "=== Envoi du compte-rendu par e-mail ===" -ForegroundColor Cyan
Write-Host "Il faut un « mot de passe d'application » Gmail (pas votre mot de passe habituel) :"
Write-Host "  1. Ouvrez https://myaccount.google.com/apppasswords (la validation en 2 étapes doit être activée)"
Write-Host "  2. Créez un mot de passe nommé par exemple « Compte-rendu Bourse »"
Write-Host "  3. Copiez les 16 lettres affichées"
Write-Host ""

$exp = ''
while (-not $exp) { $exp = Read-Host "Adresse Gmail qui envoie" }
$dest = Read-Host "Adresse qui reçoit le compte-rendu [$exp]"
if (-not $dest) { $dest = $exp }
$mdp = Read-Host "Mot de passe d'application (16 lettres, les espaces sont acceptés)" -AsSecureString

# Retire les espaces éventuels du mot de passe d'application
$clair = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($mdp))
$mdp = ConvertTo-SecureString ($clair -replace '\s', '') -AsPlainText -Force
$clair = $null

$cfg = [pscustomobject]@{ expediteur = $exp; destinataires = @($dest); smtp = 'smtp.gmail.com'; port = 587 }
[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'notification.json'), ($cfg | ConvertTo-Json), (New-Object Text.UTF8Encoding($false)))
$mdp | ConvertFrom-SecureString | Set-Content -Path (Join-Path $PSScriptRoot 'mail.secret')

Write-Host ""
Write-Host "Configuration enregistrée. Envoi d'un e-mail de test…" -ForegroundColor Cyan
& (Join-Path $PSScriptRoot 'generer.ps1') -Mail -ForcerMail
Write-Host ""
Write-Host "Si le message n'arrive pas, regardez journal-mail.txt et vos spams."
