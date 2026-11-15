<#
  Génère le compte-rendu de clôture de chaque action suivie.
  - Données : Yahoo Finance (historique quotidien sur 1 an), aucune clé d'API.
  - Sortie  : site/index.html (dernier compte-rendu), site/rapports/AAAA-MM-JJ.html (archives),
              site/archives.html, site/data/AAAA-MM-JJ.json
  Fonctionne avec Windows PowerShell 5.1 et PowerShell 7 (Windows, Linux, macOS).
#>
param(
  [string]$OutDir     = (Join-Path $PSScriptRoot 'site'),
  [string]$ConfigFile = (Join-Path $PSScriptRoot 'actions.json'),
  [switch]$Mail,        # envoie le compte-rendu par e-mail (voir configurer-mail.ps1)
  [switch]$ForcerMail   # envoie même si la Bourse était fermée aujourd'hui (test)
)

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$fr  = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')
$inv = [Globalization.CultureInfo]::InvariantCulture
$UA  = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'

$config = Get-Content $ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json

# ---------------------------------------------------------------- utilitaires
function Enc([string]$s) { [System.Net.WebUtility]::HtmlEncode($s) }
function Num($v, [string]$fmt = 'N2') { ([double]$v).ToString($fmt, $fr) }
function Pct($v) {
  if ($null -eq $v) { return '—' }
  $s = if ($v -gt 0) { '+' } else { '' }
  $s + ([double]$v).ToString('N2', $fr) + [char]0xA0 + '%'
}
function Prix($v) { $x = [math]::Abs([double]$v); if ($x -lt 1) { Num $v 'N4' } elseif ($x -lt 10) { Num $v 'N3' } else { Num $v } }
function Initiales([string]$s) {
  $m = @(($s -replace "[^\p{L}\s\-]", '') -split '[\s\-]+' | Where-Object { $_ })
  if ($m.Count -ge 2) { "$($m[0][0])$($m[1][0])".ToUpper() } else { "$($m[0][0])".ToUpper() }
}
function Teinte([string]$s) { $t = 0; foreach ($c in $s.ToCharArray()) { $t = ($t * 31 + [int]$c) % 360 }; $t }
function Inv($v) { ([double]$v).ToString('0.####', $inv) }
function Moyenne($vals) { if (@($vals).Count -eq 0) { return $null }; ($vals | Measure-Object -Average).Average }
function Signe($v) { if ($v -gt 0.05) { 'up' } elseif ($v -lt -0.05) { 'down' } else { 'flat' } }
function Majuscule([string]$s) { if ($s) { $s.Substring(0,1).ToUpper() + $s.Substring(1) } else { $s } }
function Devise($meta) { switch ([string]$meta.currency) { 'EUR' { '€' } 'USD' { '$' } 'GBP' { '£' } 'GBp' { 'p' } '' { '€' } default { [string]$meta.currency } } }
# Vrai si la séance de ce titre n'est pas encore terminée (ex. Wall Street ferme à 22h, heure de Paris)
function Est-Provisoire($meta, $a) {
  try {
    $fin = [DateTimeOffset]::FromUnixTimeSeconds([long]$meta.currentTradingPeriod.regular.end)
    return ([DateTimeOffset]::UtcNow -lt $fin.AddMinutes(10)) -and ($fin.AddSeconds([int]$meta.gmtoffset).Date -eq $a.Date)
  } catch { return $false }
}

function Get-Historique([string]$symbole) {
  $url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + [uri]::EscapeDataString($symbole) + '?range=1y&interval=1d'
  $r = $null
  for ($essai = 1; $essai -le 3; $essai++) {
    try { $r = Invoke-RestMethod -Uri $url -UserAgent $UA -TimeoutSec 30; break }
    catch { if ($essai -eq 3) { throw }; Start-Sleep -Seconds (2 * $essai) }
  }
  $res = $r.chart.result[0]
  if (-not $res -or -not $res.timestamp) { throw "Pas de données pour $symbole" }
  $q   = $res.indicators.quote[0]
  $off = [int]$res.meta.gmtoffset
  $lignes = New-Object System.Collections.Generic.List[object]
  for ($i = 0; $i -lt $res.timestamp.Count; $i++) {
    if ($null -eq $q.close[$i]) { continue }
    $c = [double]$q.close[$i]
    $lignes.Add([pscustomobject]@{
      Date   = [DateTimeOffset]::FromUnixTimeSeconds([long]$res.timestamp[$i]).AddSeconds($off).Date
      Open   = if ($null -ne $q.open[$i]) { [double]$q.open[$i] } else { $c }
      High   = if ($null -ne $q.high[$i]) { [double]$q.high[$i] } else { $c }
      Low    = if ($null -ne $q.low[$i])  { [double]$q.low[$i] }  else { $c }
      Close  = $c
      Volume = if ($null -ne $q.volume[$i]) { [double]$q.volume[$i] } else { 0 }
    })
  }
  # Yahoo renvoie parfois deux barres pour la même date : on garde la dernière
  $parDate = [ordered]@{}
  foreach ($l in $lignes) { $parDate[$l.Date.ToString('yyyy-MM-dd')] = $l }
  [pscustomobject]@{ Meta = $res.meta; Lignes = @($parDate.Values) }
}

function Get-Analyse($h, [int]$sem = 5, [int]$mois = 21) {
  $r = $h.Lignes; $n = $r.Count
  if ($n -lt 2) { return $null }
  $j = $r[$n - 1]; $p = $r[$n - 2]

  $perf = { param($k) if ($n -gt $k) { ($j.Close / $r[$n - 1 - $k].Close - 1) * 100 } else { $null } }

  $avant = @($r | Where-Object { $_.Date.Year -lt $j.Date.Year })
  $baseYtd = if ($avant.Count) { $avant[-1].Close } else { $r[0].Open }

  $vols20  = @($r[[math]::Max(0, $n - 21)..($n - 2)] | ForEach-Object Volume | Where-Object { $_ -gt 0 })
  $volMoy  = Moyenne $vols20
  $mm20    = Moyenne @($r[[math]::Max(0, $n - 20)..($n - 1)] | ForEach-Object Close)
  $mm50    = Moyenne @($r[[math]::Max(0, $n - 50)..($n - 1)] | ForEach-Object Close)
  $mm50Vei = Moyenne @($r[[math]::Max(0, $n - 51)..($n - 2)] | ForEach-Object Close)

  $hautsAvant = @($r[0..($n - 2)] | ForEach-Object High)
  $basAvant   = @($r[0..($n - 2)] | ForEach-Object Low)
  $plusHaut   = ($hautsAvant | Measure-Object -Maximum).Maximum
  $plusBas    = ($basAvant   | Measure-Object -Minimum).Minimum

  # série de hausses / baisses consécutives
  $serie = 0
  for ($i = $n - 1; $i -ge 1; $i--) {
    $d = $r[$i].Close - $r[$i - 1].Close
    if ($serie -ge 0 -and $d -gt 0) { $serie++ }
    elseif ($serie -le 0 -and $d -lt 0) { $serie-- }
    else { break }
  }

  $amp = $j.High - $j.Low
  [pscustomobject]@{
    Date       = $j.Date
    Cloture    = $j.Close
    Veille     = $p.Close
    Ouverture  = $j.Open
    Haut       = $j.High
    Bas        = $j.Low
    Volume     = $j.Volume
    Var        = ($j.Close / $p.Close - 1) * 100
    Perf5      = & $perf $sem
    Perf21     = & $perf $mois
    VolMoy     = $volMoy
    PerfYtd    = ($j.Close / $baseYtd - 1) * 100
    Perf1an    = ($j.Close / $r[0].Close - 1) * 100
    RatioVol   = if ($volMoy -and $j.Volume -gt 0) { $j.Volume / $volMoy } else { $null }
    MM20       = $mm20
    MM50       = $mm50
    MM50Veille = $mm50Vei
    Haut52     = [math]::Max($plusHaut, $j.High)
    Bas52      = [math]::Min($plusBas, $j.Low)
    NouvHaut   = $j.High -gt $plusHaut -and $n -gt 200
    NouvBas    = $j.Low  -lt $plusBas  -and $n -gt 200
    PosJour    = if ($amp -gt 0) { ($j.Close - $j.Low) / $amp } else { 0.5 }
    Amplitude  = $amp / $p.Close * 100
    Serie      = $serie
    Spark      = @($r[[math]::Max(0, $n - 66)..($n - 1)] | ForEach-Object Close)
  }
}

function Get-Commentaire($nom, $a, [int]$rang, [int]$total, $varIndice, $nomIndice, [string]$dev = '€', [switch]$Crypto) {
  $phrases = Get-CommentaireBrut $nom $a $rang $total $varIndice $nomIndice $dev
  if (-not $Crypto) { return $phrases }
  # Les cryptos cotent en continu : on parle de 24 h glissantes plutôt que de séance
  $rempl = [ordered]@{
    'termine quasiment stable' = 'reste quasiment stable'; ", à $(Prix $a.Cloture)" = " sur 24 h, à $(Prix $a.Cloture)"
    '20 dernières séances' = '20 derniers jours'; 'Séance agitée, avec une amplitude' = 'Forte volatilité, avec un écart'
    'Clôture au plus haut de la journée' = 'Évolue au plus haut des dernières 24 h'; 'Clôture au plus bas de la journée' = 'Évolue au plus bas des dernières 24 h'
    'Le titre a toutefois' = 'Le cours a toutefois'; 'en fin de séance' = 'au cours des dernières heures'
    'séance consécutive' = 'jour consécutif'
  }
  foreach ($p in $phrases) { $s = $p; foreach ($k in $rempl.Keys) { $s = $s.Replace($k, $rempl[$k]) }; $s }
}

function Get-CommentaireBrut($nom, $a, [int]$rang, [int]$total, $varIndice, $nomIndice, [string]$dev = '€') {
  $v = $a.Var
  $verbe =
    if     ([math]::Abs($v) -lt 0.1) { 'termine quasiment stable (' + (Pct $v) + ')' }
    elseif ($v -ge 3)    { "s'envole de " + (Pct $v) }
    elseif ($v -ge 1.5)  { 'progresse nettement de ' + (Pct $v) }
    elseif ($v -gt 0)    { 'gagne ' + (Pct $v) }
    elseif ($v -le -3)   { 'chute de ' + (Pct $v) }
    elseif ($v -le -1.5) { 'recule nettement de ' + (Pct $v) }
    else                 { 'cède ' + (Pct $v) }
  $phrases = New-Object System.Collections.Generic.List[string]
  $phrases.Add("$nom $verbe, à $(Prix $a.Cloture) $dev.")

  if ($total -gt 0) {
    if ($rang -eq 1 -and $v -gt 0)        { $phrases.Add('Meilleure performance de la séance parmi les valeurs suivies.') }
    elseif ($rang -eq $total -and $v -lt 0) { $phrases.Add('Plus forte baisse de la séance parmi les valeurs suivies.') }
    elseif ($rang -le 5 -and $v -gt 0)    { $phrases.Add("Dans le top 5 des hausses ($($rang)e).") }
    elseif ($rang -gt $total - 5 -and $v -lt 0) { $phrases.Add('Parmi les 5 plus fortes baisses.') }
  }

  if ($null -ne $varIndice) {
    $ecart = $v - $varIndice
    if ($ecart -ge 1)      { $phrases.Add("Surperforme le $nomIndice de $(Num $ecart) point$(if ($ecart -ge 2) {'s'}).") }
    elseif ($ecart -le -1) { $phrases.Add("Sous-performe le $nomIndice de $(Num ([math]::Abs($ecart))) point$(if ($ecart -le -2) {'s'}).") }
  }

  if ($null -ne $a.RatioVol) {
    if ($a.RatioVol -ge 2)      { $phrases.Add("Volumes très étoffés : $(Num $a.RatioVol 'N1') fois la moyenne des 20 dernières séances.") }
    elseif ($a.RatioVol -ge 1.4) { $phrases.Add("Volumes soutenus ($(Num $a.RatioVol 'N1') × la moyenne 20 jours).") }
    elseif ($a.RatioVol -le 0.6) { $phrases.Add('Échanges peu fournis par rapport à la moyenne.') }
  }

  if ($a.Amplitude -ge 4) { $phrases.Add("Séance agitée, avec une amplitude de $(Num $a.Amplitude 'N1') % entre plus bas et plus haut.") }
  if ($a.PosJour -ge 0.9 -and $v -gt 0)     { $phrases.Add('Clôture au plus haut de la journée.') }
  elseif ($a.PosJour -le 0.1 -and $v -lt 0) { $phrases.Add('Clôture au plus bas de la journée.') }
  elseif ($a.PosJour -le 0.25 -and $v -gt 0.5) { $phrases.Add('Le titre a toutefois rendu une bonne partie de ses gains en fin de séance.') }
  elseif ($a.PosJour -ge 0.75 -and $v -lt -0.5) { $phrases.Add('Il a néanmoins réduit ses pertes en fin de séance.') }

  if ($a.Serie -ge 3)      { $phrases.Add("$($a.Serie)e séance consécutive de hausse.") }
  elseif ($a.Serie -le -3) { $phrases.Add("$(-$a.Serie)e séance consécutive de baisse.") }

  if ($a.NouvHaut)    { $phrases.Add('Nouveau plus haut sur 52 semaines.') }
  elseif ($a.NouvBas) { $phrases.Add('Nouveau plus bas sur 52 semaines.') }
  else {
    $dHaut = ($a.Haut52 - $a.Cloture) / $a.Haut52 * 100
    $dBas  = ($a.Cloture - $a.Bas52) / $a.Bas52 * 100
    if ($dHaut -le 3)   { $phrases.Add("Évolue à moins de 3 % de son plus haut annuel ($(Prix $a.Haut52) $dev).") }
    elseif ($dBas -le 3) { $phrases.Add("Évolue à moins de 3 % de son plus bas annuel ($(Prix $a.Bas52) $dev).") }
  }

  if ($a.MM50 -and $a.MM50Veille) {
    if ($a.Veille -lt $a.MM50Veille -and $a.Cloture -gt $a.MM50)     { $phrases.Add('Repasse au-dessus de sa moyenne mobile 50 jours.') }
    elseif ($a.Veille -gt $a.MM50Veille -and $a.Cloture -lt $a.MM50) { $phrases.Add('Passe sous sa moyenne mobile 50 jours.') }
    elseif ($a.Cloture -gt $a.MM20 -and $a.Cloture -gt $a.MM50 -and $a.MM20 -gt $a.MM50) { $phrases.Add('Tendance de fond haussière (cours au-dessus des moyennes 20 et 50 jours).') }
    elseif ($a.Cloture -lt $a.MM20 -and $a.Cloture -lt $a.MM50 -and $a.MM20 -lt $a.MM50) { $phrases.Add('Tendance de fond baissière (cours sous les moyennes 20 et 50 jours).') }
  }
  $phrases
}

function Sparkline($vals, [int]$w = 160, [int]$h = 40) {
  $min = ($vals | Measure-Object -Minimum).Minimum
  $max = ($vals | Measure-Object -Maximum).Maximum
  $rng = if ($max -eq $min) { 1 } else { $max - $min }
  $pas = $w / [math]::Max(1, $vals.Count - 1)
  $pts = for ($i = 0; $i -lt $vals.Count; $i++) {
    (Inv ($i * $pas)) + ',' + (Inv ($h - 3 - (($vals[$i] - $min) / $rng) * ($h - 6)))
  }
  $cls = if ($vals[-1] -ge $vals[0]) { 'up' } else { 'down' }
  $ligne = $pts -join ' '
  "<svg class='spark $cls' viewBox='0 0 $w $h' preserveAspectRatio='none' aria-hidden='true'><polygon class='area' points='0,$h $ligne $w,$h'/><polyline points='$ligne'/></svg>"
}

function Classe-Tuile($v) {
  if ($v -ge 2) { 'u3' } elseif ($v -ge 0.75) { 'u2' } elseif ($v -gt 0.05) { 'u1' }
  elseif ($v -le -2) { 'd3' } elseif ($v -le -0.75) { 'd2' } elseif ($v -lt -0.05) { 'd1' } else { 'z' }
}

# ---------------------------------------------------------------- collecte
Write-Host "Indice $($config.indice.nom)…"
$indice = $null; $aIndice = $null
try { $indice = Get-Historique $config.indice.symbole; $aIndice = Get-Analyse $indice }
catch { Write-Warning "Indice indisponible : $($_.Exception.Message)" }

$resultats = New-Object System.Collections.Generic.List[object]
foreach ($act in $config.actions) {
  try {
    $h = Get-Historique $act.symbole
    $a = Get-Analyse $h
    if (-not $a) { throw 'historique trop court' }
    $resultats.Add([pscustomobject]@{ Symbole = $act.symbole; Nom = $act.nom; Secteur = $act.secteur; A = $a; Meta = $h.Meta; Hist = $h; Dev = (Devise $h.Meta); Provisoire = $false })
    Write-Host ("  {0,-26} {1,10}  {2}" -f $act.nom, (Prix $a.Cloture), (Pct $a.Var))
  } catch {
    Write-Warning "$($act.symbole) ignorée : $($_.Exception.Message)"
  }
  Start-Sleep -Milliseconds 250
}
if ($resultats.Count -eq 0) { throw 'Aucune donnée récupérée.' }

# Favoris : suivis à part, comparés à leur propre indice de référence
$favoris = New-Object System.Collections.Generic.List[object]
$cacheRef = @{}
foreach ($f in @($config.favoris)) {
  if (-not $f) { continue }
  try {
    $h = Get-Historique $f.symbole
    $a = Get-Analyse $h
    if (-not $a) { throw 'historique trop court' }
    $refVar = $null; $refNom = $null
    if ($f.reference) {
      $refNom = $f.reference.nom
      if (-not $cacheRef.ContainsKey($f.reference.symbole)) {
        try { $cacheRef[$f.reference.symbole] = Get-Analyse (Get-Historique $f.reference.symbole) } catch { $cacheRef[$f.reference.symbole] = $null }
      }
      $ra = $cacheRef[$f.reference.symbole]
      if ($ra -and $ra.Date -eq $a.Date) { $refVar = $ra.Var }
    }
    $favoris.Add([pscustomobject]@{ Symbole = $f.symbole; Nom = $f.nom; Secteur = $f.secteur; A = $a; Meta = $h.Meta; Hist = $h
      Dev = (Devise $h.Meta); Provisoire = (Est-Provisoire $h.Meta $a); Rang = 0; RefVar = $refVar; RefNom = $refNom })
    Write-Host ("  ★ {0,-24} {1,10}  {2}" -f $f.nom, (Prix $a.Cloture), (Pct $a.Var))
  } catch {
    Write-Warning "Favori $($f.symbole) ignoré : $($_.Exception.Message)"
  }
}

# Cryptomonnaies : marché ouvert 24 h/24, variation calculée sur 24 h glissantes (données horaires)
function Get-Crypto24h([string]$symbole) {
  $url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + [uri]::EscapeDataString($symbole) + '?range=5d&interval=1h'
  $r = Invoke-RestMethod -Uri $url -UserAgent $UA -TimeoutSec 30
  $res = $r.chart.result[0]; $q = $res.indicators.quote[0]
  $barres = for ($i = 0; $i -lt $res.timestamp.Count; $i++) {
    if ($null -eq $q.close[$i]) { continue }
    [pscustomobject]@{ T = [long]$res.timestamp[$i]; C = [double]$q.close[$i]
      H = if ($null -ne $q.high[$i]) { [double]$q.high[$i] } else { [double]$q.close[$i] }
      L = if ($null -ne $q.low[$i])  { [double]$q.low[$i] }  else { [double]$q.close[$i] }
      V = if ($null -ne $q.volume[$i]) { [double]$q.volume[$i] } else { 0 } }
  }
  $prix = if ($res.meta.regularMarketPrice) { [double]$res.meta.regularMarketPrice } else { $barres[-1].C }
  $tFin = if ($res.meta.regularMarketTime) { [long]$res.meta.regularMarketTime } else { $barres[-1].T }
  $avant = @($barres | Where-Object { $_.T -le $tFin - 86400 })
  $ref = if ($avant.Count) { $avant[-1] } else { $barres[0] }
  $fenetre = @($barres | Where-Object { $_.T -gt $ref.T })
  [pscustomobject]@{
    Prix = $prix; Ref = $ref.C
    Haut = [math]::Max($prix, ($fenetre | ForEach-Object H | Measure-Object -Maximum).Maximum)
    Bas  = [math]::Min($prix, ($fenetre | ForEach-Object L | Measure-Object -Minimum).Minimum)
    Volume = ($fenetre | ForEach-Object V | Measure-Object -Sum).Sum
    Heure = [DateTimeOffset]::FromUnixTimeSeconds($tFin)
  }
}

$cryptos = New-Object System.Collections.Generic.List[object]
$heureCrypto = $null
foreach ($c in @($config.cryptos)) {
  if (-not $c) { continue }
  try {
    $h = Get-Historique $c.symbole
    $a = Get-Analyse $h 7 30
    if (-not $a) { throw 'historique trop court' }
    $d = Get-Crypto24h $c.symbole
    $amp = $d.Haut - $d.Bas
    $a.Cloture = $d.Prix; $a.Veille = $d.Ref; $a.Ouverture = $d.Ref
    $a.Var = ($d.Prix / $d.Ref - 1) * 100
    $a.Haut = $d.Haut; $a.Bas = $d.Bas; $a.Volume = $d.Volume
    $a.Amplitude = $amp / $d.Ref * 100
    $a.PosJour = if ($amp -gt 0) { ($d.Prix - $d.Bas) / $amp } else { 0.5 }
    $a.RatioVol = if ($a.VolMoy -and $d.Volume -gt 0) { $d.Volume / $a.VolMoy } else { $null }
    $a.Spark = @($a.Spark[0..($a.Spark.Count - 2)]) + $d.Prix
    $heureCrypto = $d.Heure
    $cryptos.Add([pscustomobject]@{ Symbole = $c.symbole; Nom = $c.nom; Secteur = 'Cryptomonnaie'; Code = $c.code; Couleur = $c.couleur
      A = $a; Meta = $h.Meta; Hist = $h; Dev = (Devise $h.Meta); Provisoire = $false; Rang = 0 })
    Write-Host ("  ₿ {0,-24} {1,12}  {2}" -f $c.nom, (Prix $a.Cloture), (Pct $a.Var))
  } catch {
    Write-Warning "Crypto $($c.symbole) ignorée : $($_.Exception.Message)"
  }
}
# Les altcoins sont comparés au Bitcoin
$btc = $cryptos | Where-Object { $_.Code -eq 'BTC' } | Select-Object -First 1
foreach ($x in $cryptos) {
  $x | Add-Member RefVar $(if ($btc -and $x -ne $btc) { $btc.A.Var } else { $null })
  $x | Add-Member RefNom 'Bitcoin'
}

# Date de séance = date la plus fréquente parmi les dernières barres (évite qu'une valeur suspendue fausse la date)
$dateSeance = ($resultats | Group-Object { $_.A.Date.ToString('yyyy-MM-dd') } | Sort-Object Count -Descending | Select-Object -First 1).Name
$resultats  = @($resultats | Sort-Object { $_.A.Var } -Descending)
$total      = $resultats.Count
for ($i = 0; $i -lt $total; $i++) { $resultats[$i] | Add-Member Rang ($i + 1) }

$finSeance = $null
try { $finSeance = [DateTimeOffset]::FromUnixTimeSeconds([long]$resultats[0].Meta.currentTradingPeriod.regular.end) } catch {}
$provisoire = $finSeance -and [DateTimeOffset]::UtcNow -lt $finSeance.AddMinutes(10) -and $finSeance.ToString('yyyy-MM-dd') -eq $dateSeance

$varIndice = if ($aIndice) { $aIndice.Var } else { $null }
$nomIndice = $config.indice.nom
$hausses   = @($resultats | Where-Object { $_.A.Var -gt 0.05 }).Count
$baisses   = @($resultats | Where-Object { $_.A.Var -lt -0.05 }).Count
$stables   = $total - $hausses - $baisses
$top       = $resultats[0]
$flop      = $resultats[-1]

# ---------------------------------------------------------------- synthèse du marché
$dateObj   = [datetime]::ParseExact($dateSeance, 'yyyy-MM-dd', $inv)
$dateLongue = Majuscule ($dateObj.ToString('dddd d MMMM yyyy', $fr))

$synthese = New-Object System.Collections.Generic.List[string]
if ($aIndice) {
  $sens = if ($aIndice.Var -gt 0.05) { 'en hausse de ' + (Pct $aIndice.Var) } elseif ($aIndice.Var -lt -0.05) { 'en baisse de ' + (Pct $aIndice.Var) } else { 'à l''équilibre (' + (Pct $aIndice.Var) + ')' }
  $synthese.Add("Le $nomIndice termine $sens, à $(Num $aIndice.Cloture) points.")
  if ($aIndice.Serie -ge 3)      { $synthese.Add("C'est sa $($aIndice.Serie)e séance de hausse d'affilée.") }
  elseif ($aIndice.Serie -le -3) { $synthese.Add("C'est sa $(-$aIndice.Serie)e séance de baisse d'affilée.") }
  $synthese.Add("Sur la semaine glissante l'indice affiche $(Pct $aIndice.Perf5), et $(Pct $aIndice.PerfYtd) depuis le 1er janvier.")
}
$synthese.Add("$hausses valeur$(if ($hausses -gt 1) {'s'}) en hausse, $baisses en baisse$(if ($stables) { " et $stables stable$(if ($stables -gt 1) {'s'})" }).")
$synthese.Add("En tête : $($top.Nom) ($(Pct $top.A.Var)). Lanterne rouge : $($flop.Nom) ($(Pct $flop.A.Var)).")
$volFort = $resultats | Where-Object { $_.A.RatioVol -ge 2 } | Sort-Object { $_.A.RatioVol } -Descending | Select-Object -First 3
if ($volFort) { $synthese.Add('Volumes inhabituels sur ' + (($volFort | ForEach-Object { $_.Nom }) -join ', ') + '.') }

# Secteurs
$secteurs = $resultats | Group-Object Secteur | ForEach-Object {
  [pscustomobject]@{ Nom = $_.Name; Var = Moyenne @($_.Group | ForEach-Object { $_.A.Var }); N = $_.Count }
} | Sort-Object Var -Descending

# ---------------------------------------------------------------- HTML
# --- Historique détaillé des cryptos (pour la fenêtre de suivi)
function Get-Serie([string]$symbole, [string]$params) {
  $url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + [uri]::EscapeDataString($symbole) + '?' + $params
  $r = Invoke-RestMethod -Uri $url -UserAgent $UA -TimeoutSec 30
  $res = $r.chart.result[0]; $q = $res.indicators.quote[0]
  $t = New-Object System.Collections.Generic.List[long]; $c = New-Object System.Collections.Generic.List[double]
  for ($i = 0; $i -lt $res.timestamp.Count; $i++) {
    if ($null -ne $q.close[$i]) { $t.Add([long]$res.timestamp[$i]); $c.Add([double]$q.close[$i]) }
  }
  [pscustomobject]@{ T = $t; C = $c }
}
function Arrondi([double]$v) { if ($v -ge 1000) { [math]::Round($v, 2) } elseif ($v -ge 1) { [math]::Round($v, 4) } else { [math]::Round($v, 7) } }
function SerieJson($s, [int]$max, [long]$tFin, [double]$prix) {
  $n = $s.T.Count
  $pas = [math]::Max(1, [math]::Ceiling($n / $max))
  $ts = New-Object System.Collections.Generic.List[string]; $cs = New-Object System.Collections.Generic.List[string]
  for ($i = 0; $i -lt $n; $i += $pas) { $ts.Add([string]$s.T[$i]); $cs.Add((Arrondi $s.C[$i]).ToString('R', $inv)) }
  if ($n -and (($n - 1) % $pas) -ne 0) { $ts.Add([string]$s.T[$n - 1]); $cs.Add((Arrondi $s.C[$n - 1]).ToString('R', $inv)) }
  if ($tFin -and $n -and $tFin -gt $s.T[$n - 1]) { $ts.Add([string]$tFin); $cs.Add((Arrondi $prix).ToString('R', $inv)) }
  '{"t":[' + ($ts -join ',') + '],"c":[' + ($cs -join ',') + ']}'
}


$css = @'
:root{color-scheme:dark;
--bg:#04050b;--ink:#eaf0ff;--muted:#8b95b4;--dim:#5d6684;--line:rgba(255,255,255,.08);--line2:rgba(255,255,255,.14);
--cyan:#22d3ee;--violet:#8b5cf6;--pink:#ec4899;--amber:#fbbf24;
--up:#2ef5a2;--down:#ff4d6d;--flat:#9aa3bf;
--upa:rgba(46,245,162,.12);--downa:rgba(255,77,109,.12);--flata:rgba(255,255,255,.06);
--glass:linear-gradient(155deg,rgba(255,255,255,.085),rgba(255,255,255,.018) 60%);
--sh:0 18px 40px -12px rgba(0,0,0,.7),0 2px 6px rgba(0,0,0,.4),inset 0 1px 0 rgba(255,255,255,.07);
--r:20px;--mono:"JetBrains Mono",ui-monospace,monospace;--disp:"Space Grotesk","Inter",sans-serif}
*{box-sizing:border-box}html{-webkit-text-size-adjust:100%;scroll-behavior:smooth;background:var(--bg)}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.6 "Inter",system-ui,-apple-system,"Segoe UI",sans-serif;-webkit-font-smoothing:antialiased;overflow-x:hidden}
body.noscroll{overflow:hidden}
a{color:inherit}
::selection{background:rgba(34,211,238,.35)}

/* Fond : aurore animée + grille */
.bgfx{position:fixed;inset:0;z-index:0;pointer-events:none;overflow:hidden}
.bgfx i{position:absolute;border-radius:50%;filter:blur(90px);opacity:.55;animation:drift 26s ease-in-out infinite alternate}
.bgfx i:nth-child(1){width:620px;height:620px;left:-160px;top:-180px;background:#5b21b6}
.bgfx i:nth-child(2){width:560px;height:560px;right:-160px;top:8%;background:#0e7490;animation-duration:32s;animation-delay:-8s}
.bgfx i:nth-child(3){width:700px;height:520px;left:30%;bottom:-260px;background:#9d174d;opacity:.35;animation-duration:38s}
.bgfx::after{content:"";position:absolute;inset:0;background-image:linear-gradient(rgba(255,255,255,.035) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.035) 1px,transparent 1px);background-size:56px 56px;mask-image:radial-gradient(ellipse 80% 70% at 50% 0%,#000 30%,transparent 80%);-webkit-mask-image:radial-gradient(ellipse 80% 70% at 50% 0%,#000 30%,transparent 80%)}
@keyframes drift{to{transform:translate(80px,60px) scale(1.15)}}
.page{position:relative;z-index:1}
.wrap{max-width:1240px;margin:0 auto;padding:0 16px}
.num,.mono{font-family:var(--mono);font-variant-numeric:tabular-nums;letter-spacing:-.02em}

/* En-tête */
header.top{position:sticky;top:0;z-index:40;background:rgba(4,5,11,.6);backdrop-filter:blur(18px) saturate(1.6);-webkit-backdrop-filter:blur(18px) saturate(1.6);border-bottom:1px solid var(--line)}
header.top .wrap{display:flex;align-items:center;gap:20px;height:64px}
.brand{display:flex;align-items:center;gap:11px;font-family:var(--disp);font-weight:700;font-size:17px;letter-spacing:-.01em;text-decoration:none;white-space:nowrap}
.brand svg{filter:drop-shadow(0 0 10px rgba(46,245,162,.55))}
nav.sections{display:flex;gap:2px;overflow-x:auto;scrollbar-width:none;margin-left:auto}
nav.sections::-webkit-scrollbar{display:none}
nav.sections a{color:var(--muted);text-decoration:none;font-size:14px;font-weight:500;padding:8px 13px;border-radius:999px;white-space:nowrap;transition:color .2s,background .2s}
nav.sections a:hover{color:#fff;background:rgba(255,255,255,.07)}
@media (max-width:760px){header.top .wrap{height:auto;flex-direction:column;align-items:stretch;gap:4px;padding-top:10px;padding-bottom:6px}nav.sections{margin-left:-10px}}

/* Bandeau défilant */
.ticker{border-bottom:1px solid var(--line);background:rgba(255,255,255,.02);overflow:hidden;white-space:nowrap;font-size:13px;mask-image:linear-gradient(90deg,transparent,#000 6%,#000 94%,transparent);-webkit-mask-image:linear-gradient(90deg,transparent,#000 6%,#000 94%,transparent)}
.ticker .track{display:inline-flex;gap:30px;padding:9px 0;animation:defile 90s linear infinite}
.ticker:hover .track{animation-play-state:paused}
.ticker span{display:inline-flex;gap:8px;align-items:center}.ticker b{font-weight:600;color:#cfd6ee}
@keyframes defile{to{transform:translateX(-50%)}}

/* Verre */
.glass{position:relative;background:var(--glass);border:1px solid var(--line);border-radius:var(--r);box-shadow:var(--sh);backdrop-filter:blur(16px);-webkit-backdrop-filter:blur(16px);--glow:rgba(34,211,238,.55)}
.glass::before{content:"";position:absolute;inset:-1px;border-radius:inherit;padding:1px;background:linear-gradient(140deg,var(--glow),transparent 35%,transparent 65%,rgba(139,92,246,.45));-webkit-mask:linear-gradient(#000 0 0) content-box,linear-gradient(#000 0 0);-webkit-mask-composite:xor;mask:linear-gradient(#000 0 0) content-box exclude,linear-gradient(#000 0 0);opacity:.55;transition:opacity .3s;pointer-events:none}
.glass::after{content:"";position:absolute;inset:0;border-radius:inherit;background:radial-gradient(420px circle at var(--mx,50%) var(--my,-20%),rgba(140,200,255,.13),transparent 45%);opacity:0;transition:opacity .35s;pointer-events:none}
.glass:hover::before{opacity:1}.glass:hover::after{opacity:1}
.glass.up{--glow:rgba(46,245,162,.75)}.glass.down{--glow:rgba(255,77,109,.75)}.glass.gold{--glow:rgba(251,191,36,.85)}
.tilt{transform:perspective(1000px) rotateX(var(--rx,0deg)) rotateY(var(--ry,0deg)) translateY(var(--ty,0px));transition:transform .25s cubic-bezier(.2,.8,.2,1),box-shadow .35s;will-change:transform}
.tilt:hover{--ty:-6px;box-shadow:0 34px 60px -18px rgba(0,0,0,.85),0 0 0 1px rgba(255,255,255,.04),0 0 50px -12px var(--glow)}

/* Reveal au défilement */
.js .rv{opacity:0;translate:0 26px}
.js .rv.in{opacity:1;translate:0 0;transition:opacity .7s ease,translate .7s cubic-bezier(.2,.8,.2,1)}

/* Héros */
.hero{padding:56px 0 30px}
.hero .wrap{display:grid;grid-template-columns:1.2fr 1fr;gap:40px;align-items:center}
@media (max-width:900px){.hero .wrap{grid-template-columns:1fr}.hero{padding-top:36px}}
.chip{display:inline-flex;align-items:center;gap:8px;font-size:12px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#bfe9ff;background:rgba(34,211,238,.08);border:1px solid rgba(34,211,238,.25);padding:6px 12px;border-radius:999px}
.chip i{width:7px;height:7px;border-radius:50%;background:var(--up);box-shadow:0 0 0 0 rgba(46,245,162,.6);animation:pulse 2s infinite}
@keyframes pulse{70%{box-shadow:0 0 0 9px rgba(46,245,162,0)}100%{box-shadow:0 0 0 0 rgba(46,245,162,0)}}
h1{font-family:var(--disp);font-weight:700;font-size:clamp(38px,6.4vw,74px);line-height:1;letter-spacing:-.035em;margin:20px 0 20px;background:linear-gradient(100deg,#ffffff 10%,#c4b5fd 45%,#22d3ee 85%);-webkit-background-clip:text;background-clip:text;color:transparent}
.lead{font-size:17.5px;line-height:1.7;color:#b9c2dc;margin:0;max-width:620px}
.lead strong{color:#fff;font-weight:600}
.banner{display:inline-block;background:rgba(251,191,36,.1);color:#fcd34d;border:1px solid rgba(251,191,36,.35);border-radius:12px;padding:8px 12px;margin:0 0 16px;font-size:14px}
.idxcard{padding:24px}
.idxhead{display:flex;justify-content:space-between;align-items:flex-start;gap:12px}
.eyebrow{font-size:12px;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:var(--muted)}
.idxval{font-family:var(--mono);font-size:44px;font-weight:600;letter-spacing:-.04em;line-height:1.1;margin-top:6px}
.idxcard .spark{height:130px;margin:16px 0 14px}
.idxstats{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;border-top:1px solid var(--line);padding-top:16px}
.idxstats small{display:block;color:var(--muted);font-size:11.5px;margin-bottom:2px}
.idxstats b{font-size:14px;font-weight:600}

/* Pastilles */
.pill{display:inline-flex;align-items:center;gap:5px;font-family:var(--mono);font-weight:600;font-size:13px;padding:4px 10px;border-radius:999px;white-space:nowrap;letter-spacing:-.02em;border:1px solid transparent}
.pill.up{background:var(--upa);color:var(--up);border-color:rgba(46,245,162,.28);box-shadow:0 0 18px -4px rgba(46,245,162,.35)}
.pill.down{background:var(--downa);color:var(--down);border-color:rgba(255,77,109,.28);box-shadow:0 0 18px -4px rgba(255,77,109,.35)}
.pill.flat{background:var(--flata);color:var(--flat);border-color:var(--line)}
.pill.up::before{content:"▲";font-size:8px}.pill.down::before{content:"▼";font-size:8px}
.t-up{color:var(--up)}.t-down{color:var(--down)}.t-flat{color:var(--flat)}

/* Chiffres clés */
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
@media (max-width:900px){.kpis{grid-template-columns:repeat(2,1fr)}}
.kpi{padding:18px 20px;min-width:0}
.kpi small{display:block;color:var(--muted);font-size:12.5px;font-weight:500;margin-bottom:8px}
.kpi b{display:block;font-family:var(--disp);font-size:21px;letter-spacing:-.02em;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.kpi .sub{margin-top:8px}
.breadth{display:flex;height:6px;border-radius:999px;overflow:hidden;margin-top:12px;background:rgba(255,255,255,.06);gap:2px}
.breadth i{display:block;border-radius:999px}.breadth .bu{background:linear-gradient(90deg,#10b981,var(--up));box-shadow:0 0 12px var(--up)}.breadth .bd{background:linear-gradient(90deg,#f43f5e,var(--down));box-shadow:0 0 12px var(--down)}.breadth .bf{background:var(--dim)}

/* Sections */
section.block{padding-top:72px;scroll-margin-top:76px}
.shead{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-bottom:22px}
h2{font-family:var(--disp);font-weight:700;font-size:clamp(28px,3.6vw,40px);letter-spacing:-.03em;margin:8px 0 0;line-height:1.05}
h2 em{font-style:normal;background:linear-gradient(90deg,#c4b5fd,#22d3ee);-webkit-background-clip:text;background-clip:text;color:transparent}
.shead p{margin:0;color:var(--muted);font-size:14px;max-width:520px}
.panel{padding:22px}
.panel h3{font-family:var(--disp);font-size:15px;margin:0 0 14px;color:#cfd6ee;font-weight:600;letter-spacing:.01em}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
@media (max-width:900px){.grid2{grid-template-columns:1fr}}

/* Avatars */
.mono-av,.coin{flex:none;width:40px;height:40px;border-radius:12px;display:grid;place-items:center;font-family:var(--disp);font-weight:700;font-size:14px}
.mono-av{background:hsl(var(--h) 80% 60% / .13);color:hsl(var(--h) 95% 78%);border:1px solid hsl(var(--h) 90% 70% / .28);box-shadow:0 0 22px -6px hsl(var(--h) 90% 60% / .55)}
.coin{border-radius:50%;background:radial-gradient(circle at 30% 25%,rgba(255,255,255,.45),transparent 45%),var(--c);color:#fff;font-size:10.5px;letter-spacing:.02em;box-shadow:0 0 0 1px rgba(255,255,255,.15),0 0 26px -4px var(--c);text-shadow:0 1px 2px rgba(0,0,0,.4)}

/* Palmarès */
.mv{display:grid;grid-template-columns:36px minmax(0,1fr) 100px auto;align-items:center;gap:12px;padding:9px 6px;border-radius:12px;text-decoration:none;transition:background .2s}
.mv:hover{background:rgba(255,255,255,.05)}
.mv .mono-av{width:36px;height:36px;font-size:12px;border-radius:10px}
.mv .nm{font-weight:600;font-size:14.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.mvb{height:6px;background:rgba(255,255,255,.06);border-radius:999px;overflow:hidden}.mvb i{display:block;height:100%;border-radius:999px}
.mvb i.up{background:linear-gradient(90deg,#0ea5e9,var(--up));box-shadow:0 0 10px var(--up)}.mvb i.down{background:linear-gradient(90deg,#a855f7,var(--down));box-shadow:0 0 10px var(--down)}.mvb i.flat{background:var(--dim)}
@media (max-width:440px){.mv{grid-template-columns:36px minmax(0,1fr) auto}.mvb{display:none}}

/* Carte de chaleur */
.heat{display:grid;grid-template-columns:repeat(auto-fill,minmax(122px,1fr));gap:6px}
.tile{display:block;text-decoration:none;border-radius:12px;padding:10px 11px;font-size:12.5px;line-height:1.3;min-height:62px;border:1px solid var(--line);transition:transform .2s,box-shadow .2s,border-color .2s}
.tile:hover{transform:translateY(-3px) scale(1.04);z-index:2;position:relative}
.tile b{display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-weight:600;color:#e6ebfa}
.tile span{font-family:var(--mono);font-weight:600;font-size:13px}
.u1{background:rgba(46,245,162,.08);color:var(--up)}.u2{background:rgba(46,245,162,.18);color:var(--up);border-color:rgba(46,245,162,.25)}.u3{background:rgba(46,245,162,.32);color:#d6ffee;border-color:rgba(46,245,162,.5);box-shadow:0 0 24px -6px rgba(46,245,162,.6)}
.d1{background:rgba(255,77,109,.08);color:var(--down)}.d2{background:rgba(255,77,109,.18);color:var(--down);border-color:rgba(255,77,109,.25)}.d3{background:rgba(255,77,109,.32);color:#ffe0e6;border-color:rgba(255,77,109,.5);box-shadow:0 0 24px -6px rgba(255,77,109,.6)}
.z{background:rgba(255,255,255,.04);color:var(--flat)}
.tile.u1:hover,.tile.u2:hover,.tile.u3:hover{box-shadow:0 0 30px -4px rgba(46,245,162,.6)}.tile.d1:hover,.tile.d2:hover,.tile.d3:hover{box-shadow:0 0 30px -4px rgba(255,77,109,.6)}
.legend{display:flex;gap:4px;align-items:center;font-size:12px;color:var(--muted);font-family:var(--mono)}
.legend i{width:20px;height:10px;border-radius:3px;display:inline-block;border:0}

/* Secteurs */
.sec{display:grid;grid-template-columns:minmax(0,200px) 1fr 80px;gap:14px;align-items:center;padding:8px 0;border-bottom:1px solid var(--line);font-size:14px}
.sec:last-child{border-bottom:0}
.sec .nm{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.sec .nm small{color:var(--dim)}
.dbar{position:relative;height:8px}.dbar::before{content:"";position:absolute;left:50%;top:-5px;bottom:-5px;width:1px;background:var(--line2)}
.dbar i{position:absolute;top:0;bottom:0;border-radius:999px}
.dbar i.up{left:50%;background:linear-gradient(90deg,#0ea5e9,var(--up));box-shadow:0 0 12px rgba(46,245,162,.6)}
.dbar i.down{right:50%;background:linear-gradient(270deg,#a855f7,var(--down));box-shadow:0 0 12px rgba(255,77,109,.6)}
.dbar i.flat{left:50%;background:var(--dim)}
.sec .v{text-align:right;font-family:var(--mono);font-weight:600;font-size:13.5px}
@media (max-width:540px){.sec{grid-template-columns:minmax(0,1fr) 70px 70px}}

/* Fiches actions */
.tools{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
.tools input,.tools select{font:inherit;font-size:14px;padding:11px 16px;border:1px solid var(--line2);border-radius:999px;background:rgba(255,255,255,.04);color:var(--ink);backdrop-filter:blur(10px)}
.tools select option{background:#0b0d16}
.tools input{flex:1;min-width:220px}
.tools input:focus,.tools select:focus{outline:none;border-color:var(--cyan);box-shadow:0 0 0 4px rgba(34,211,238,.15),0 0 30px -6px rgba(34,211,238,.6)}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(330px,1fr));gap:16px}
@media (max-width:440px){.cards{grid-template-columns:1fr}}
.card{padding:20px;scroll-margin-top:90px;min-width:0}
.card:target{outline:2px solid var(--cyan);outline-offset:3px}
.chead{display:flex;justify-content:space-between;gap:12px;align-items:flex-start}
.cid{display:flex;gap:12px;align-items:center;min-width:0}
.cid h3{margin:0;font-family:var(--disp);font-size:17px;letter-spacing:-.02em;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tk{color:var(--muted);font-size:12.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.price{text-align:right;flex:none}.price b{font-family:var(--mono);font-size:19px;font-weight:600;display:block;letter-spacing:-.03em;margin-bottom:4px}
.tags{display:flex;gap:6px;flex-wrap:wrap;margin-top:12px}
.tag{font-size:11.5px;font-weight:600;color:var(--muted);background:rgba(255,255,255,.04);border:1px solid var(--line);border-radius:999px;padding:2px 9px}
.tag small{font-weight:500;opacity:.7}
.tag.gold{color:var(--amber);background:rgba(251,191,36,.08);border-color:rgba(251,191,36,.35);box-shadow:0 0 16px -4px rgba(251,191,36,.45)}
.tag.live{color:var(--cyan);border-color:rgba(34,211,238,.4)}
.spark{width:100%;height:50px;margin:14px 0 6px;display:block;overflow:visible}
.spark.up{color:var(--up)}.spark.down{color:var(--down)}
.spark polyline{fill:none;stroke:currentColor;stroke-width:2;vector-effect:non-scaling-stroke;stroke-linejoin:round;stroke-linecap:round;filter:drop-shadow(0 0 5px currentColor)}
.spark .area{fill:currentColor;opacity:.09;stroke:none}
.card p,.ctile p{margin:8px 0 14px;font-size:14px;line-height:1.65;color:#c3cbe2}
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;background:rgba(0,0,0,.25);border:1px solid var(--line);border-radius:14px;padding:11px 12px;font-size:12.5px}
.stats div{font-family:var(--mono);font-weight:600;letter-spacing:-.02em}.stats small{display:block;color:var(--muted);font-family:"Inter",sans-serif;font-weight:500;font-size:11px;letter-spacing:0}
.range{font-size:12.5px;color:var(--muted);margin-top:12px;font-variant-numeric:tabular-nums}
.bar{position:relative;height:5px;background:linear-gradient(90deg,rgba(255,77,109,.5),rgba(255,255,255,.1),rgba(46,245,162,.5));border-radius:999px;margin:8px 0 0}
.bar i{position:absolute;top:-5px;width:15px;height:15px;margin-left:-7px;border-radius:50%;background:#fff;box-shadow:0 0 0 3px rgba(34,211,238,.35),0 0 16px rgba(34,211,238,.9)}
.empty{display:none;text-align:center;color:var(--muted);padding:40px}

/* Cryptos : petites cases qui s'agrandissent au survol */
.cgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:14px}
.ctile{position:relative;height:148px;outline:none;cursor:pointer}
.ctile .inner{position:absolute;top:0;left:0;width:100%;min-height:100%;padding:16px;z-index:1;
  transition:width .35s cubic-bezier(.2,.8,.2,1),transform .35s cubic-bezier(.2,.8,.2,1),box-shadow .35s,background .35s}
.ctile.gauche .inner{left:auto;right:0}
.ctile:hover,.ctile:focus-visible{z-index:30}
.ctop{padding-right:38px}
/* Bouton « Suivi détaillé » */
.dbtn{position:relative;z-index:2;display:flex;width:100%;justify-content:center;align-items:center;gap:9px;margin-top:16px;font:inherit;font-weight:600;font-size:13.5px;color:#d9f6ff;cursor:pointer;padding:11px 14px;border-radius:13px;border:1px solid rgba(34,211,238,.35);background:linear-gradient(90deg,rgba(34,211,238,.10),rgba(139,92,246,.14));transition:background .25s,color .25s,box-shadow .25s,transform .2s}
.dbtn:hover,.dbtn:focus-visible{background:linear-gradient(90deg,#22d3ee,#a78bfa);color:#06070d;border-color:transparent;box-shadow:0 0 30px -4px rgba(34,211,238,.85);outline:none}
.dbtn:active{transform:scale(.97)}
.dbtn.ico{position:absolute;top:12px;right:12px;width:34px;height:34px;padding:0;margin:0;border-radius:11px;z-index:3}
.ctop{display:flex;align-items:center;gap:11px;min-width:0}
.ctop .coin{width:36px;height:36px}
.ctop b{font-family:var(--disp);font-size:16px;letter-spacing:-.01em;display:block;line-height:1.2}
.ctop span{color:var(--muted);font-size:12px;font-family:var(--mono)}
.crow{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-top:12px}
.crow b{font-family:var(--mono);font-size:17px;font-weight:600;letter-spacing:-.03em;white-space:nowrap}
.ctile .spark{height:34px;margin:10px 0 0;transition:height .35s}
.ctile .more{display:grid;grid-template-rows:0fr;opacity:0;transition:grid-template-rows .35s cubic-bezier(.2,.8,.2,1),opacity .25s}
.ctile .more>div{overflow:hidden}
.cta{display:inline-flex;align-items:center;gap:8px;margin-top:14px;font-weight:600;font-size:13.5px;color:#0b0d16;background:linear-gradient(90deg,#22d3ee,#a78bfa);padding:8px 14px;border-radius:999px;box-shadow:0 0 24px -4px rgba(34,211,238,.7)}
@media (hover:hover){
  .ctile:hover .inner,.ctile:focus-visible .inner{width:max(100%,370px);transform:translateY(-10px);z-index:30;background:linear-gradient(155deg,rgba(28,32,52,.97),rgba(12,14,26,.97));box-shadow:0 40px 80px -20px rgba(0,0,0,.9),0 0 60px -12px var(--glow)}
  .ctile:hover .more,.ctile:focus-visible .more{grid-template-rows:1fr;opacity:1}
  .ctile:hover .spark,.ctile:focus-visible .spark{height:58px}
}
.ctile .hint{display:none;font-size:11.5px;color:var(--cyan);margin-top:8px}
@media (hover:none){.ctile .hint{display:block}.ctile{height:162px}}
.ctile:focus-visible .inner{outline:2px solid var(--cyan);outline-offset:3px}

/* Fenêtre de suivi détaillé */
.modal{position:fixed;inset:0;z-index:100;display:grid;place-items:center;padding:16px}
.modal[hidden]{display:none}
.mback{position:absolute;inset:0;background:rgba(2,3,8,.72);backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);opacity:0;transition:opacity .3s}
.mpanel{position:relative;width:min(1000px,100%);max-height:calc(100vh - 32px);overflow:auto;padding:26px;background:linear-gradient(155deg,rgba(26,30,50,.96),rgba(9,11,22,.97));opacity:0;transform:translateY(24px) scale(.97);transition:opacity .3s,transform .35s cubic-bezier(.2,.8,.2,1)}
.modal.on .mback{opacity:1}.modal.on .mpanel{opacity:1;transform:none}
.mclose{position:absolute;top:16px;right:16px;width:38px;height:38px;border-radius:50%;border:1px solid var(--line2);background:rgba(255,255,255,.05);color:var(--ink);font-size:20px;cursor:pointer;z-index:2;transition:background .2s,transform .2s}
.mclose:hover{background:rgba(255,255,255,.12);transform:rotate(90deg)}
.mhead{display:flex;align-items:center;gap:16px;flex-wrap:wrap;padding-right:48px}
.mhead .coin{width:54px;height:54px;font-size:13px}
.mhead h3{margin:0;font-family:var(--disp);font-size:28px;letter-spacing:-.03em;line-height:1.1}
.mprice{margin-left:auto;text-align:right}
.mprice b{display:block;font-family:var(--mono);font-size:32px;font-weight:600;letter-spacing:-.04em;line-height:1.1}
@media (max-width:600px){.mprice{margin-left:0;text-align:left;width:100%}.mpanel{padding:20px 16px}}
.tabs{display:grid;grid-template-columns:repeat(5,1fr);gap:8px;margin:22px 0 16px}
@media (max-width:600px){.tabs{grid-template-columns:repeat(3,1fr)}}
.tabs button{font:inherit;text-align:left;cursor:pointer;color:var(--ink);background:rgba(255,255,255,.035);border:1px solid var(--line);border-radius:14px;padding:10px 12px;transition:border-color .2s,background .2s,box-shadow .2s}
.tabs button small{display:block;color:var(--muted);font-size:12px;margin-bottom:2px}
.tabs button span{font-family:var(--mono);font-weight:600;font-size:14.5px;letter-spacing:-.02em}
.tabs button:hover{background:rgba(255,255,255,.07)}
.tabs button[aria-selected="true"]{border-color:rgba(34,211,238,.6);background:rgba(34,211,238,.09);box-shadow:0 0 0 3px rgba(34,211,238,.1),0 0 30px -8px rgba(34,211,238,.8)}
.tabs button[disabled]{opacity:.35;cursor:default}
.chart{position:relative;height:340px;border-radius:16px;background:rgba(0,0,0,.28);border:1px solid var(--line);overflow:hidden;touch-action:pan-y}
@media (max-width:600px){.chart{height:250px}}
.chart svg{display:block;width:100%;height:100%}
.chart .gl{stroke:rgba(255,255,255,.06);stroke-dasharray:3 5}
.chart .lbl{fill:var(--dim);font:11px var(--mono)}
.chart .ln{fill:none;stroke-width:2.2;stroke-linejoin:round}
.chart .xl{stroke:rgba(255,255,255,.35);stroke-dasharray:3 4}
.tip{position:absolute;top:12px;pointer-events:none;background:rgba(8,10,20,.92);border:1px solid var(--line2);border-radius:10px;padding:7px 10px;font-size:12.5px;white-space:nowrap;box-shadow:0 10px 30px rgba(0,0,0,.5);opacity:0;transition:opacity .15s}
.tip b{display:block;font-family:var(--mono);font-size:14px;letter-spacing:-.02em}
.tip span{color:var(--muted)}
.mstats{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-top:16px}
@media (max-width:700px){.mstats{grid-template-columns:repeat(2,1fr)}}
.mstats div{background:rgba(255,255,255,.035);border:1px solid var(--line);border-radius:14px;padding:12px 14px}
.mstats small{display:block;color:var(--muted);font-size:12px}
.mstats b{font-family:var(--mono);font-size:15.5px;letter-spacing:-.02em}
.note{color:var(--muted);font-size:13px;margin:14px 0 0}

footer{color:var(--muted);font-size:13px;border-top:1px solid var(--line);margin-top:90px;padding:28px 0 50px;background:rgba(0,0,0,.25)}
footer .wrap{display:flex;justify-content:space-between;gap:20px;flex-wrap:wrap}

.arch{list-style:none;padding:0;margin:0}
.arch a{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:15px 8px;text-decoration:none;border-bottom:1px solid var(--line);border-radius:10px;transition:background .2s}
.arch a:hover{background:rgba(255,255,255,.05)}

@media (prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}.js .rv{opacity:1;translate:none}}
'@

$js = @'
<script>
(function(){
  var reduce=matchMedia('(prefers-reduced-motion: reduce)').matches;
  var fine=matchMedia('(hover:hover) and (pointer:fine)').matches;

  /* Apparition au défilement */
  var rv=[].slice.call(document.querySelectorAll('.rv'));
  if('IntersectionObserver' in window && !reduce){
    var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('in');io.unobserve(e.target);}});},{rootMargin:'0px 0px -6% 0px'});
    rv.forEach(function(e){io.observe(e);});
  } else rv.forEach(function(e){e.classList.add('in');});

  /* Relief 3D + reflet qui suit la souris */
  if(fine && !reduce){
    [].forEach.call(document.querySelectorAll('.glass'),function(el){
      var tilt=el.classList.contains('tilt');
      el.addEventListener('pointermove',function(ev){
        var r=el.getBoundingClientRect(),px=(ev.clientX-r.left)/r.width,py=(ev.clientY-r.top)/r.height;
        el.style.setProperty('--mx',(px*100)+'%');el.style.setProperty('--my',(py*100)+'%');
        if(tilt){el.style.setProperty('--rx',((.5-py)*7).toFixed(2)+'deg');el.style.setProperty('--ry',((px-.5)*9).toFixed(2)+'deg');}
      });
      el.addEventListener('pointerleave',function(){el.style.setProperty('--rx','0deg');el.style.setProperty('--ry','0deg');});
    });
  }

  /* Recherche et tri des actions */
  var q=document.getElementById('q'),s=document.getElementById('tri'),box=document.getElementById('cards'),vide=document.getElementById('vide');
  if(q&&box){
    var cards=[].slice.call(box.children);
    var norm=function(t){return t.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'');};
    var run=function(){
      var t=norm(q.value),k=s.value,n=0;
      cards.sort(function(a,b){
        if(k==='nom')return a.dataset.nom.localeCompare(b.dataset.nom,'fr');
        if(k==='varasc')return parseFloat(a.dataset.var)-parseFloat(b.dataset.var);
        return parseFloat(b.dataset[k])-parseFloat(a.dataset[k]);
      });
      cards.forEach(function(c){var ok=c.dataset.find.indexOf(t)>-1;c.style.display=ok?'':'none';if(ok){n++;c.classList.add('in');}box.appendChild(c);});
      if(vide)vide.style.display=n?'none':'block';
    };
    q.addEventListener('input',run);s.addEventListener('change',run);
  }

  /* Cases crypto : s'agrandissent vers la gauche si elles débordent à droite */
  var tiles=[].slice.call(document.querySelectorAll('.ctile'));
  tiles.forEach(function(t){
    t.addEventListener('mouseenter',function(){var r=t.getBoundingClientRect();t.classList.toggle('gauche',r.left+370>document.documentElement.clientWidth-8);});
    t.addEventListener('click',function(){ouvrir(t.dataset.id,t);});
    t.addEventListener('keydown',function(e){if(e.target===t&&(e.key==='Enter'||e.key===' ')){e.preventDefault();ouvrir(t.dataset.id,t);}});
  });
  /* Bouton « Suivi détaillé » présent sur chaque case */
  document.addEventListener('click',function(e){
    var b=e.target.closest&&e.target.closest('.dbtn');
    if(b){e.preventDefault();e.stopPropagation();ouvrir(b.dataset.id,b);}
  },true);

  /* Fenêtre de suivi détaillé */
  var el=document.getElementById('cdata'),modal=document.getElementById('modal');
  if(!el||!modal)return;
  var data=JSON.parse(el.textContent),cur=null,per='j',origine=null,serie=null;
  var racine=document.body.dataset.racine||'',ver=document.body.dataset.v||'';
  window.__S=window.__S||{};
  var PC=[['j','24 heures'],['s','7 jours'],['m','1 mois'],['a','1 an'],['max','Depuis le début']];
  var PA=[['j','Dernière séance'],['s','5 séances'],['m','1 mois'],['a','1 an'],['max','Depuis l’introduction']];
  var nf=function(d){return new Intl.NumberFormat('fr-FR',{minimumFractionDigits:d,maximumFractionDigits:d});};
  var prix=function(v){var a=Math.abs(v);return nf(a<1?4:a<10?3:2).format(v)+' '+(cur?cur.dev:'€');};
  var pct=function(v){return(v>0?'+':'')+nf(2).format(v)+' %';};
  var sg=function(v){return v>0.05?'up':v<-0.05?'down':'flat';};
  var dCourt=new Intl.DateTimeFormat('fr-FR',{weekday:'short',day:'numeric',month:'short',hour:'2-digit',minute:'2-digit'});
  var dLong=new Intl.DateTimeFormat('fr-FR',{day:'numeric',month:'long',year:'numeric'});
  var dMois=new Intl.DateTimeFormat('fr-FR',{month:'short',year:'numeric'});
  var $=function(id){return document.getElementById(id);};
  var variation=function(sr){return sr&&sr.c.length>1?(sr.c[sr.c.length-1]/sr.c[0]-1)*100:null;};

  function charger(id,ok,ko){
    if(window.__S[id])return ok(window.__S[id]);
    var s=document.createElement('script');
    s.src=racine+'series/'+encodeURIComponent(id)+'.js?v='+ver;
    s.onload=function(){window.__S[id]?ok(window.__S[id]):ko();};s.onerror=ko;
    document.head.appendChild(s);
  }
  function ouvrir(id,src){
    var m=data[id];if(!m)return;cur=m;origine=src||null;serie=null;pts=[];
    var co=$('mcoin');co.textContent=cur.code;
    co.className=cur.crypto?'coin':'mono-av';
    co.style.setProperty('--c',cur.couleur||'#888');co.style.setProperty('--h',cur.h);
    $('mtitle').textContent=cur.nom;$('msub').textContent=cur.sub;
    $('mprice').textContent=prix(cur.prix);
    $('mvar').className='pill '+sg(cur['var']);$('mvar').textContent=pct(cur['var'])+(cur.crypto?' sur 24 h':' sur la séance');
    $('mtabs').innerHTML='';$('mstats').innerHTML='';$('msvg').innerHTML='';
    $('mnote').textContent='Chargement de l’historique…';
    modal.hidden=false;document.body.classList.add('noscroll');
    requestAnimationFrame(function(){modal.classList.add('on');});
    $('mclose').focus();
    charger(id,function(sx){
      if(cur!==m)return;serie=sx;
      var tabs=$('mtabs');
      (cur.crypto?PC:PA).forEach(function(p){
        var sr=serie[p[0]],v=variation(sr),b=document.createElement('button');
        b.type='button';b.setAttribute('role','tab');b.dataset.p=p[0];
        b.innerHTML='<small>'+p[1]+'</small><span class="t-'+(v==null?'flat':sg(v))+'">'+(v==null?'—':pct(v))+'</span>';
        if(!sr||sr.c.length<2)b.disabled=true;
        b.addEventListener('click',function(){per=p[0];dessiner();});
        tabs.appendChild(b);
      });
      per=serie.j&&serie.j.c.length>1?'j':'a';
      requestAnimationFrame(dessiner);
    },function(){$('mnote').textContent='Historique indisponible pour le moment.';});
  }
  function fermer(){
    modal.classList.remove('on');document.body.classList.remove('noscroll');
    setTimeout(function(){modal.hidden=true;},280);
    if(origine)origine.focus();
  }
  modal.addEventListener('click',function(e){if(e.target.hasAttribute('data-close'))fermer();});
  document.addEventListener('keydown',function(e){if(e.key==='Escape'&&!modal.hidden)fermer();});
  window.addEventListener('resize',function(){if(!modal.hidden&&serie)dessiner();});

  var pts=[],sr=null;
  function dessiner(){
    [].forEach.call($('mtabs').children,function(b){b.setAttribute('aria-selected',b.dataset.p===per?'true':'false');});
    sr=serie[per];var box=$('mchart'),svg=$('msvg');
    var W=box.clientWidth,H=box.clientHeight,pl=12,pr=78,pt=22,pb=30;
    if(!sr||sr.c.length<2){svg.innerHTML='';return;}
    var c=sr.c,t=sr.t,n=c.length,mn=Math.min.apply(null,c),mx=Math.max.apply(null,c),pad=(mx-mn)*.08||mx*.01;
    var lo=mn-pad,hi=mx+pad,t0=t[0],t1=t[n-1];
    var X=function(v){return pl+(v-t0)/(t1-t0||1)*(W-pl-pr);},Y=function(v){return pt+(1-(v-lo)/(hi-lo))*(H-pt-pb);};
    pts=[];var d='';
    for(var i=0;i<n;i++){var x=X(t[i]),y=Y(c[i]);pts.push([x,y,i]);d+=(i?'L':'M')+x.toFixed(1)+' '+y.toFixed(1);}
    var v=(c[n-1]/c[0]-1)*100,col=v>=0?'#2ef5a2':'#ff4d6d';
    var g='';
    for(var k=0;k<=4;k++){var yv=lo+(hi-lo)*k/4,yy=Y(yv);g+='<line class="gl" x1="'+pl+'" x2="'+(W-pr+6)+'" y1="'+yy+'" y2="'+yy+'"/><text class="lbl" x="'+(W-pr+12)+'" y="'+(yy+4)+'">'+prix(yv).replace(' €','')+'</text>';}
    var fx=per==='j'||per==='s'?function(s){return new Intl.DateTimeFormat('fr-FR',per==='j'?{hour:'2-digit',minute:'2-digit'}:{weekday:'short',day:'numeric'}).format(new Date(s*1000));}
           :per==='m'?function(s){return new Intl.DateTimeFormat('fr-FR',{day:'numeric',month:'short'}).format(new Date(s*1000));}
           :function(s){return dMois.format(new Date(s*1000));};
    for(var k2=0;k2<=4;k2++){var tv=t0+(t1-t0)*k2/4,xx=X(tv),anc=k2===0?'start':k2===4?'end':'middle';g+='<text class="lbl" text-anchor="'+anc+'" x="'+xx+'" y="'+(H-9)+'">'+fx(tv)+'</text>';}
    var iMax=c.indexOf(mx),iMin=c.indexOf(mn);
    svg.setAttribute('viewBox','0 0 '+W+' '+H);
    svg.innerHTML='<defs><linearGradient id="ga" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="'+col+'" stop-opacity=".38"/><stop offset="1" stop-color="'+col+'" stop-opacity="0"/></linearGradient>'
      +'<filter id="gw" x="-10%" y="-30%" width="120%" height="160%"><feGaussianBlur stdDeviation="4" result="b"/><feMerge><feMergeNode in="b"/><feMergeNode in="SourceGraphic"/></feMerge></filter></defs>'+g
      +'<path d="'+d+'L'+X(t1).toFixed(1)+' '+(H-pb)+'L'+X(t0).toFixed(1)+' '+(H-pb)+'Z" fill="url(#ga)"/>'
      +'<path class="ln" d="'+d+'" stroke="'+col+'" filter="url(#gw)"/>'
      +'<circle cx="'+X(t[iMax])+'" cy="'+Y(mx)+'" r="3.5" fill="#fff" opacity=".8"/><circle cx="'+X(t[iMin])+'" cy="'+Y(mn)+'" r="3.5" fill="#fff" opacity=".8"/>'
      +'<circle cx="'+X(t1)+'" cy="'+Y(c[n-1])+'" r="5" fill="'+col+'" filter="url(#gw)"/>'
      +'<line id="mxl" class="xl" y1="'+pt+'" y2="'+(H-pb)+'" x1="-10" x2="-10"/><circle id="mdot" r="5.5" cx="-10" cy="-10" fill="#fff" stroke="'+col+'" stroke-width="3"/>';
    var amp=(mx/mn-1)*100;
    var debut=per==='max'?'<div><small>Premier cours ('+dLong.format(new Date(t0*1000))+')</small><b>'+prix(c[0])+'</b></div>':'<div><small>Début de période</small><b>'+prix(c[0])+'</b></div>';
    $('mstats').innerHTML='<div><small>Variation</small><b class="t-'+sg(v)+'">'+pct(v)+'</b></div>'+debut
      +'<div><small>Plus haut</small><b>'+prix(mx)+'</b></div><div><small>Plus bas</small><b>'+prix(mn)+'</b></div><div><small>Écart haut / bas</small><b>'+pct(amp).replace('+','')+'</b></div>';
    $('mnote').textContent=per==='max'?'Historique disponible depuis le '+dLong.format(new Date(t0*1000))+' (début des données Yahoo Finance, un point par mois environ).':'Survolez ou touchez le graphique pour lire un cours précis.';
    $('mtip').style.opacity=0;
  }
  function survol(ev){
    if(!pts.length)return;
    var r=$('mchart').getBoundingClientRect(),x=(ev.touches?ev.touches[0].clientX:ev.clientX)-r.left;
    var lo=0,hi=pts.length-1;while(hi-lo>1){var m=(lo+hi)>>1;if(pts[m][0]<x)lo=m;else hi=m;}
    var p=Math.abs(pts[lo][0]-x)<Math.abs(pts[hi][0]-x)?pts[lo]:pts[hi],i=p[2];
    $('mxl').setAttribute('x1',p[0]);$('mxl').setAttribute('x2',p[0]);
    $('mdot').setAttribute('cx',p[0]);$('mdot').setAttribute('cy',p[1]);
    var tip=$('mtip'),dt=new Date(sr.t[i]*1000);
    tip.innerHTML='<b>'+prix(sr.c[i])+'</b><span>'+(per==='j'||per==='s'||per==='m'?dCourt.format(dt):dLong.format(dt))+'</span>';
    tip.style.opacity=1;var w=tip.offsetWidth;tip.style.left=Math.max(8,Math.min(r.width-w-8,p[0]-w/2))+'px';
  }
  $('mchart').addEventListener('pointermove',survol);
  $('mchart').addEventListener('touchmove',survol,{passive:true});
  $('mchart').addEventListener('pointerleave',function(){$('mtip').style.opacity=0;$('mxl').setAttribute('x1',-10);$('mxl').setAttribute('x2',-10);$('mdot').setAttribute('cx',-10);});
})();
</script>
'@

function Sans-Accents([string]$s) {
  $d = $s.Normalize([Text.NormalizationForm]::FormD)
  (-join ($d.ToCharArray() | Where-Object { [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne 'NonSpacingMark' })).ToLower()
}
function Pastille($v, [string]$suffixe = '') { "<span class='pill $(Signe $v)'>$(Pct $v)$suffixe</span>" }
function T($v) { "<span class='t-$(Signe $v)'>$(Pct $v)</span>" }
function IdDe($x, [string]$prefixe = '') { $prefixe + ($x.Symbole -replace '[^A-Za-z0-9]', '-').ToLower() }
function Avatar($x) {
  if ($x.Code) { "<span class='coin' style='--c:$($x.Couleur)'>$(Enc $x.Code)</span>" }
  else { "<span class='mono-av' style='--h:$(Teinte $x.Nom)'>$(Enc (Initiales $x.Nom))</span>" }
}

$logo = "<svg width='30' height='30' viewBox='0 0 30 30' aria-hidden='true'><defs><linearGradient id='lg' x1='0' y1='0' x2='1' y2='1'><stop offset='0' stop-color='#22d3ee'/><stop offset='1' stop-color='#8b5cf6'/></linearGradient></defs><rect width='30' height='30' rx='9' fill='url(#lg)'/><path d='M7 20l5-5 4 3 7-8' fill='none' stroke='#fff' stroke-width='2.4' stroke-linecap='round' stroke-linejoin='round'/><circle cx='23' cy='10' r='2.3' fill='#fff'/></svg>"
$iconeSuivi = "<svg width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2.2' stroke-linecap='round' stroke-linejoin='round' aria-hidden='true'><path d='M3 3v18h18'/><path d='M7 15l4-4 3 3 6-7'/><path d='M16 7h4v4'/></svg>"
$favicon = [uri]::EscapeDataString(($logo -replace "aria-hidden='true'", "xmlns='http://www.w3.org/2000/svg'"))

$heureParis = try { [TimeZoneInfo]::ConvertTime([DateTimeOffset]::UtcNow, [TimeZoneInfo]::FindSystemTimeZoneById('Romance Standard Time')) } catch {
              try { [TimeZoneInfo]::ConvertTime([DateTimeOffset]::UtcNow, [TimeZoneInfo]::FindSystemTimeZoneById('Europe/Paris')) } catch { [DateTimeOffset]::Now } }
$majTxt = "Mis à jour le " + $heureParis.ToString("d MMMM 'à' HH'h'mm", $fr)

function Page([string]$titre, [string]$racine, [string]$corps, [bool]$accueil = $true) {
  $nav = if ($accueil) {
    "<a href='#synthese'>Synthèse</a><a href='#favoris'>Mes favoris</a><a href='#crypto'>Crypto</a><a href='#marche'>Marché</a><a href='#valeurs'>Valeurs</a><a href='${racine}archives.html'>Archives</a>"
  } else { "<a href='${racine}index.html'>Dernière séance</a><a href='${racine}archives.html'>Archives</a>" }
@"
<!doctype html>
<html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>$(Enc $titre)</title>
<meta name="description" content="Compte-rendu de clôture de chaque action, de mes favoris et des cryptomonnaies, mis à jour chaque soir.">
<meta name="theme-color" content="#04050b">
<link rel="icon" href="data:image/svg+xml,$favicon">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500;600&family=Space+Grotesk:wght@500;600;700&display=swap" rel="stylesheet">
<script>document.documentElement.classList.add('js')</script>
<style>$css</style></head><body data-racine="$racine" data-v="$version">
<div class="bgfx" aria-hidden="true"><i></i><i></i><i></i></div>
<div class="page">
<header class="top"><div class="wrap"><a class="brand" href="${racine}index.html">$logo<span>$(Enc $config.titre)</span></a>
<nav class="sections" aria-label="Sections">$nav</nav></div></header>
$corps
<footer><div class="wrap"><span>$majTxt · Données Yahoo Finance, susceptibles d'être différées.</span>
<span>Commentaires générés automatiquement à partir des cours, à titre informatif : ce n'est pas un conseil en investissement.</span></div></footer>
</div>
$js
</body></html>
"@
}

function Carte($x, [int]$total, $refVar, $refNom, [string]$prefixe = '', [string]$mode = 'action') {
  $a = $x.A; $dev = $x.Dev
  $id = IdDe $x $prefixe
  $com = Get-Commentaire $x.Nom $a $x.Rang $total $refVar $refNom $dev
  if ($x.Provisoire) { $com = @($com) + 'Séance encore en cours : chiffres provisoires.' }
  $tags = New-Object System.Collections.Generic.List[string]
  if ($mode -eq 'action') { $tags.Add("<span class='tag'>#$($x.Rang) <small>sur $total</small></span>") }
  if ($mode -eq 'fav')    { $tags.Add("<span class='tag gold'>★ Favori</span>") }
  if ($x.Provisoire)      { $tags.Add("<span class='tag live'>● Séance en cours</span>") }
  if ($a.Date.ToString('yyyy-MM-dd') -ne $dateSeance) { $tags.Add("<span class='tag'>Cours du $($a.Date.ToString('dd/MM', $fr))</span>") }
  $pos52 = if ($a.Haut52 -gt $a.Bas52) { [math]::Round(($a.Cloture - $a.Bas52) / ($a.Haut52 - $a.Bas52) * 100, 1) } else { 50 }
  $rv = if ($null -ne $a.RatioVol) { $a.RatioVol } else { 0 }
  $find = Sans-Accents "$($x.Nom) $($x.Symbole) $($x.Secteur)"
  $volTxt = if ($null -ne $a.RatioVol) { (Num $a.RatioVol 'N1') + ' ×' } else { '—' }
  $cls = if ($mode -eq 'fav') { 'gold' } else { Signe $a.Var }
@"
<article class='card glass tilt rv $cls' id='$id' data-var='$(Inv $a.Var)' data-vol='$(Inv $rv)' data-ytd='$(Inv $a.PerfYtd)' data-nom='$(Enc $x.Nom)' data-find='$(Enc $find)'>
<div class='chead'><div class='cid'>$(Avatar $x)<div style='min-width:0'><h3>$(Enc $x.Nom)</h3><div class='tk'>$(Enc $x.Symbole) · $(Enc $x.Secteur)</div></div></div>
<div class='price'><b>$(Prix $a.Cloture) $dev</b>$(Pastille $a.Var)</div></div>
<div class='tags'>$($tags -join '')</div>
$(Sparkline $a.Spark 300 50)
<p>$(Enc ($com -join ' '))</p>
<div class='stats'>
<div><small>1 semaine</small>$(T $a.Perf5)</div>
<div><small>1 mois</small>$(T $a.Perf21)</div>
<div><small>Depuis janv.</small>$(T $a.PerfYtd)</div>
<div><small>Volume / moy.</small>$volTxt</div>
</div>
<div class='range'>Séance : $(Prix $a.Bas) – $(Prix $a.Haut) $dev · ouverture $(Prix $a.Ouverture) $dev</div>
<div class='range'>52 semaines : $(Prix $a.Bas52) – $(Prix $a.Haut52) $dev<div class='bar' title='Position dans la fourchette annuelle'><i style='left:$(Inv $pos52)%'></i></div></div>
<button class='dbtn' type='button' data-id='$id'>$iconeSuivi<span>Suivi détaillé</span></button>
</article>
"@
}

function CarteCrypto($x) {
  $a = $x.A; $dev = $x.Dev
  $com = Get-Commentaire $x.Nom $a 0 0 $x.RefVar $x.RefNom $dev -Crypto
  $pos52 = if ($a.Haut52 -gt $a.Bas52) { [math]::Round(($a.Cloture - $a.Bas52) / ($a.Haut52 - $a.Bas52) * 100, 1) } else { 50 }
  $volTxt = if ($null -ne $a.RatioVol) { (Num $a.RatioVol 'N1') + ' ×' } else { '—' }
@"
<div class='ctile rv' data-id='$(IdDe $x 'cr-')' tabindex='0' role='button' aria-label='$(Enc $x.Nom) : ouvrir le suivi détaillé' id='$(IdDe $x 'cr-')'>
<div class='inner glass $(Signe $a.Var)'>
<button class='dbtn ico' type='button' data-id='$(IdDe $x 'cr-')' aria-label='Suivi détaillé de $(Enc $x.Nom)' title='Suivi détaillé'>$iconeSuivi</button>
<div class='ctop'>$(Avatar $x)<div style='min-width:0'><b>$(Enc $x.Nom)</b><span>$(Enc $x.Code)</span></div></div>
<div class='crow'><b>$(Prix $a.Cloture) $dev</b>$(Pastille $a.Var)</div>
$(Sparkline $a.Spark 300 34)
<span class='hint'>Touchez pour le suivi détaillé →</span>
<div class='more'><div>
<p>$(Enc ($com -join ' '))</p>
<div class='stats'>
<div><small>7 jours</small>$(T $a.Perf5)</div>
<div><small>30 jours</small>$(T $a.Perf21)</div>
<div><small>Depuis janv.</small>$(T $a.PerfYtd)</div>
<div><small>Vol. / moy.</small>$volTxt</div>
</div>
<div class='range'>24 h : $(Prix $a.Bas) – $(Prix $a.Haut) $dev</div>
<div class='range'>52 semaines : $(Prix $a.Bas52) – $(Prix $a.Haut52) $dev<div class='bar'><i style='left:$(Inv $pos52)%'></i></div></div>
<span class='cta'>Voir le suivi détaillé →</span>
</div></div>
</div></div>
"@
}

function Movers($items, [string]$titre) {
  $max = (@($items) | ForEach-Object { [math]::Abs($_.A.Var) } | Measure-Object -Maximum).Maximum
  if (-not $max) { $max = 1 }
  $sbm = New-Object System.Text.StringBuilder
  [void]$sbm.Append("<div class='panel glass rv'><h3>$titre</h3>")
  foreach ($x in $items) {
    $w = [math]::Max(4, [math]::Abs($x.A.Var) / $max * 100)
    [void]$sbm.Append("<a class='mv' href='#$(IdDe $x)'>$(Avatar $x)<span class='nm'>$(Enc $x.Nom)</span><span class='mvb'><i class='$(Signe $x.A.Var)' style='width:$(Inv $w)%'></i></span>$(Pastille $x.A.Var)</a>")
  }
  [void]$sbm.Append("</div>")
  $sbm.ToString()
}

# Historique détaillé de chaque valeur : un petit fichier par valeur (site/series/<id>.js),
# chargé seulement quand on ouvre son suivi détaillé. Les infos de base restent dans la page.
$dirSeries = Join-Path $OutDir 'series'
New-Item -ItemType Directory -Force -Path $dirSeries | Out-Null
Get-ChildItem $dirSeries -Filter '*.js' -ErrorAction SilentlyContinue | Remove-Item -Force
$utf8s = New-Object System.Text.UTF8Encoding($false)
$maintenant = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$version = $heureParisV = [DateTimeOffset]::UtcNow.ToString('yyyyMMddHHmm')

function SerieDepuisLignes($lignes) {
  $t = New-Object System.Collections.Generic.List[long]; $c = New-Object System.Collections.Generic.List[double]
  foreach ($l in $lignes) { $t.Add(([DateTimeOffset]::new($l.Date.Ticks, [TimeSpan]::Zero)).ToUnixTimeSeconds() + 61200); $c.Add($l.Close) }
  [pscustomobject]@{ T = $t; C = $c }
}

$aSuivre = New-Object System.Collections.Generic.List[object]
foreach ($x in $favoris)   { $aSuivre.Add(@{ X = $x; Id = (IdDe $x 'fav-') }) }
foreach ($x in $cryptos)   { $aSuivre.Add(@{ X = $x; Id = (IdDe $x 'cr-') }) }
foreach ($x in $resultats) { $aSuivre.Add(@{ X = $x; Id = (IdDe $x) }) }

$metas = New-Object System.Collections.Generic.List[string]
Write-Host "Historiques détaillés ($($aSuivre.Count) valeurs)…"
foreach ($e in $aSuivre) {
  $x = $e.X; $id = $e.Id
  $crypto = [bool]$x.Code
  $tFin = if ($crypto -and $heureCrypto) { $heureCrypto.ToUnixTimeSeconds() } else { try { [long]$x.Meta.regularMarketTime } catch { 0 } }
  $req = if ($crypto) { [ordered]@{
      j = @("period1=$($maintenant - 86400)&period2=$maintenant&interval=5m", 300)
      s = @("period1=$($maintenant - 7 * 86400)&period2=$maintenant&interval=1h", 200)
      m = @("period1=$($maintenant - 30 * 86400)&period2=$maintenant&interval=1h", 360)
      max = @('range=max&interval=1wk', 700) }
  } else { [ordered]@{
      j = @('range=1d&interval=5m', 300)
      s = @('range=5d&interval=30m', 200)
      m = @('range=1mo&interval=1h', 360)
      max = @('range=max&interval=1mo', 700) } }
  $parties = New-Object System.Collections.Generic.List[string]
  foreach ($k in $req.Keys) {
    try {
      $s = Get-Serie $x.Symbole $req[$k][0]
      # La journée part du cours de référence (clôture de la veille, ou cours d'il y a 24 h pour les cryptos)
      # pour que la variation affichée corresponde à celle de la fiche
      if ($k -eq 'j' -and $s.T.Count) { $s.T.Insert(0, $s.T[0] - 300); $s.C.Insert(0, [double]$x.A.Veille) }
      $parties.Add("`"$k`":" + (SerieJson $s $req[$k][1] $tFin $x.A.Cloture))
    }
    catch { Write-Warning "Historique $k de $($x.Symbole) indisponible : $($_.Exception.Message)" }
    Start-Sleep -Milliseconds 80
  }
  if ($x.Hist) { $parties.Add('"a":' + (SerieJson (SerieDepuisLignes $x.Hist.Lignes) 370 0 0)) }
  [IO.File]::WriteAllText((Join-Path $dirSeries "$id.js"), "window.__S=window.__S||{};window.__S[`"$id`"]={$($parties -join ',')};", $utf8s)

  $code = if ($crypto) { $x.Code } else { Initiales $x.Nom }
  $sub  = if ($crypto) { "$($x.Symbole) · cours en euros" } else { "$($x.Symbole) · $($x.Secteur)" }
  $metas.Add("`"$id`":{`"nom`":$($x.Nom | ConvertTo-Json),`"sub`":$($sub | ConvertTo-Json),`"code`":$($code | ConvertTo-Json),`"couleur`":`"$($x.Couleur)`",`"h`":$(Teinte $x.Nom),`"dev`":$($x.Dev | ConvertTo-Json),`"prix`":$((Arrondi $x.A.Cloture).ToString('R', $inv)),`"var`":$([math]::Round($x.A.Var, 3).ToString('R', $inv)),`"crypto`":$(if ($crypto) {'true'} else {'false'})}")
}
$cdataJson = ('{' + ($metas -join ',') + '}').Replace('</', '<\/')

$sb = New-Object System.Text.StringBuilder

# --- Bandeau défilant
$elements = @($resultats) + @($favoris.ToArray()) + @($cryptos.ToArray())
$tk = ($elements | ForEach-Object { "<span><b>$(Enc $_.Nom)</b><span class='num'>$(Prix $_.A.Cloture) $($_.Dev)</span>$(T $_.A.Var)</span>" }) -join ''
[void]$sb.Append("<div class='ticker' aria-hidden='true'><div class='track'>$tk$tk</div></div>")

# --- Héros
[void]$sb.Append("<section class='hero' id='synthese'><div class='wrap'><div class='rv'>")
[void]$sb.Append("<span class='chip'><i></i>Compte-rendu de clôture · Paris</span><h1>$(Enc $dateLongue)</h1>")
if ($provisoire) { [void]$sb.Append("<div class='banner'>La séance n'est pas encore terminée : ces chiffres sont provisoires.</div>") }
[void]$sb.Append("<p class='lead'>$(Enc ($synthese -join ' '))</p></div>")
if ($aIndice) {
  [void]$sb.Append("<div class='idxcard glass tilt rv $(Signe $aIndice.Var)'><div class='idxhead'><div><div class='eyebrow'>$(Enc $nomIndice) · 3 mois</div><div class='idxval'>$(Num $aIndice.Cloture)</div></div>$(Pastille $aIndice.Var)</div>")
  [void]$sb.Append((Sparkline $aIndice.Spark 400 130))
  [void]$sb.Append("<div class='idxstats num'><div><small>1 semaine</small><b>$(T $aIndice.Perf5)</b></div><div><small>1 mois</small><b>$(T $aIndice.Perf21)</b></div><div><small>Depuis janv.</small><b>$(T $aIndice.PerfYtd)</b></div><div><small>1 an</small><b>$(T $aIndice.Perf1an)</b></div></div></div>")
}
[void]$sb.Append("</div></section>")

# --- Chiffres clés
$pU = [math]::Round($hausses / $total * 100, 1); $pD = [math]::Round($baisses / $total * 100, 1); $pF = [math]::Max(0, 100 - $pU - $pD)
[void]$sb.Append("<main class='wrap'><div class='kpis'>")
[void]$sb.Append("<div class='kpi glass tilt rv'><small>Hausses / baisses</small><b class='num'><span class='t-up'>$hausses</span> <span style='color:var(--dim)'>/</span> <span class='t-down'>$baisses</span></b><div class='breadth'><i class='bu' style='width:$(Inv $pU)%'></i><i class='bf' style='width:$(Inv $pF)%'></i><i class='bd' style='width:$(Inv $pD)%'></i></div></div>")
[void]$sb.Append("<div class='kpi glass tilt rv up'><small>Meilleure valeur</small><b>$(Enc $top.Nom)</b><div class='sub'>$(Pastille $top.A.Var)</div></div>")
[void]$sb.Append("<div class='kpi glass tilt rv down'><small>Pire valeur</small><b>$(Enc $flop.Nom)</b><div class='sub'>$(Pastille $flop.A.Var)</div></div>")
if ($btc) {
  [void]$sb.Append("<div class='kpi glass tilt rv $(Signe $btc.A.Var)'><small>Bitcoin · 24 h</small><b class='num'>$(Prix $btc.A.Cloture) $($btc.Dev)</b><div class='sub'>$(Pastille $btc.A.Var)</div></div>")
} elseif ($aIndice) {
  [void]$sb.Append("<div class='kpi glass tilt rv'><small>$(Enc $nomIndice) depuis janvier</small><b>$(T $aIndice.PerfYtd)</b></div>")
}
[void]$sb.Append("</div>")

# --- Favoris
if ($favoris.Count) {
  [void]$sb.Append("<section class='block' id='favoris'><div class='shead rv'><div><div class='eyebrow'>Mon portefeuille</div><h2>★ Mes <em>favoris</em></h2></div>")
  [void]$sb.Append("<p>Comparés à leur indice de référence. Wall Street ferme à 22h : les valeurs américaines sont recalées au passage de 22h30.</p></div><div class='cards'>")
  foreach ($x in $favoris) { [void]$sb.Append((Carte $x 0 $x.RefVar $x.RefNom 'fav-' 'fav')) }
  [void]$sb.Append("</div></section>")
}

# --- Cryptos
if ($cryptos.Count) {
  $hc = if ($heureCrypto) { $heureCrypto.ToOffset($heureParis.Offset) } else { $null }
  $hcTxt = if ($hc) { " Relevé du $($hc.ToString("d MMMM 'à' HH'h'mm", $fr))." } else { '' }
  [void]$sb.Append("<section class='block' id='crypto'><div class='shead rv'><div><div class='eyebrow'>Marché ouvert 24 h/24</div><h2><em>Crypto</em>monnaies</h2></div>")
  [void]$sb.Append("<p>Variations sur 24 h glissantes, en euros. Survolez une case pour le détail, cliquez pour le suivi complet (24 h, 7 jours, 1 mois, 1 an, depuis le début).$hcTxt</p></div><div class='cgrid'>")
  foreach ($x in ($cryptos | Sort-Object { $_.A.Var } -Descending)) { [void]$sb.Append((CarteCrypto $x)) }
  [void]$sb.Append("</div></section>")
}

# --- Marché
[void]$sb.Append("<section class='block' id='marche'><div class='shead rv'><div><div class='eyebrow'>Vue d'ensemble</div><h2>Le marché en un <em>coup d'œil</em></h2></div></div>")
[void]$sb.Append("<div class='grid2'>")
[void]$sb.Append((Movers @($resultats | Select-Object -First 5) 'Plus fortes hausses'))
[void]$sb.Append((Movers @($resultats | Select-Object -Last 5 | Sort-Object { $_.A.Var }) 'Plus fortes baisses'))
[void]$sb.Append("</div>")
[void]$sb.Append("<div class='panel glass rv' style='margin-top:16px'><div class='shead' style='margin-bottom:14px'><h3 style='margin:0'>Carte de la séance</h3><div class='legend'>-2 %<i class='d3'></i><i class='d2'></i><i class='d1'></i><i class='z'></i><i class='u1'></i><i class='u2'></i><i class='u3'></i>+2 %</div></div><div class='heat'>")
foreach ($x in $resultats) {
  [void]$sb.Append("<a class='tile $(Classe-Tuile $x.A.Var)' href='#$(IdDe $x)' title='$(Enc $x.Nom)'><b>$(Enc $x.Nom)</b><span>$(Pct $x.A.Var)</span></a>")
}
[void]$sb.Append("</div></div>")
$maxSec = (@($secteurs) | ForEach-Object { [math]::Abs($_.Var) } | Measure-Object -Maximum).Maximum
if (-not $maxSec) { $maxSec = 1 }
[void]$sb.Append("<div class='panel glass rv' style='margin-top:16px'><h3>Par secteur · variation moyenne</h3>")
foreach ($s in $secteurs) {
  $w = [math]::Max(1, [math]::Abs($s.Var) / $maxSec * 50)
  [void]$sb.Append("<div class='sec'><span class='nm'>$(Enc $s.Nom) <small>· $($s.N)</small></span><span class='dbar'><i class='$(Signe $s.Var)' style='width:$(Inv $w)%'></i></span><span class='v t-$(Signe $s.Var)'>$(Pct $s.Var)</span></div>")
}
[void]$sb.Append("</div></section>")

# --- Valeur par valeur
[void]$sb.Append("<section class='block' id='valeurs'><div class='shead rv'><div><div class='eyebrow'>$total actions</div><h2>Valeur par <em>valeur</em></h2></div>")
[void]$sb.Append("<div class='tools'><input id='q' type='search' placeholder='Rechercher une action, un secteur…' aria-label='Rechercher'>")
[void]$sb.Append("<select id='tri' aria-label='Trier'><option value='var'>Plus fortes hausses</option><option value='varasc'>Plus fortes baisses</option><option value='vol'>Volumes inhabituels</option><option value='ytd'>Performance depuis janvier</option><option value='nom'>Ordre alphabétique</option></select></div></div>")
[void]$sb.Append("<div class='cards' id='cards'>")
foreach ($x in $resultats) { [void]$sb.Append((Carte $x $total $varIndice $nomIndice)) }
[void]$sb.Append("</div><div class='empty' id='vide'>Aucune action ne correspond à cette recherche.</div></section></main>")

# --- Fenêtre de suivi détaillé (toutes les valeurs)
if ($metas.Count) {
  [void]$sb.Append(@"
<div class='modal' id='modal' hidden><div class='mback' data-close></div>
<div class='mpanel glass' role='dialog' aria-modal='true' aria-labelledby='mtitle'>
<button class='mclose' id='mclose' data-close aria-label='Fermer'>×</button>
<div class='mhead'><span class='coin' id='mcoin'></span><div><h3 id='mtitle'></h3><div class='tk' id='msub'></div></div>
<div class='mprice'><b id='mprice'></b><span id='mvar' class='pill'></span></div></div>
<div class='tabs' id='mtabs' role='tablist' aria-label='Période'></div>
<div class='chart' id='mchart'><svg id='msvg' role='img' aria-label='Évolution du cours'></svg><div class='tip' id='mtip'></div></div>
<div class='mstats' id='mstats'></div>
<p class='note' id='mnote'></p>
</div></div>
<script type='application/json' id='cdata'>$cdataJson</script>
"@)
}

# ---------------------------------------------------------------- écriture
$utf8 = New-Object System.Text.UTF8Encoding($false)
$dirRapports = Join-Path $OutDir 'rapports'
$dirData     = Join-Path $OutDir 'data'
New-Item -ItemType Directory -Force -Path $OutDir, $dirRapports, $dirData | Out-Null

$titrePage = "$($config.titre) — $dateLongue"
$corps = $sb.ToString()
[IO.File]::WriteAllText((Join-Path $OutDir 'index.html'), (Page $titrePage '' $corps $true), $utf8)
if (-not $provisoire) {
  [IO.File]::WriteAllText((Join-Path $dirRapports "$dateSeance.html"), (Page $titrePage '../' $corps $false), $utf8)
  $versLigne = { param($x) [pscustomobject]@{ symbole = $x.Symbole; nom = $x.Nom; secteur = $x.Secteur; cours = $x.A.Cloture; variation = [math]::Round($x.A.Var, 4)
                 ouverture = $x.A.Ouverture; haut = $x.A.Haut; bas = $x.A.Bas; volume = $x.A.Volume } }
  $donnees = [pscustomobject]@{
    date    = $dateSeance
    indice  = if ($aIndice) { [pscustomobject]@{ nom = $nomIndice; cloture = $aIndice.Cloture; variation = [math]::Round($aIndice.Var, 4) } } else { $null }
    actions = @($resultats | ForEach-Object { & $versLigne $_ })
    favoris = @($favoris   | ForEach-Object { & $versLigne $_ })
    cryptos = @($cryptos   | ForEach-Object { & $versLigne $_ })
  }
  [IO.File]::WriteAllText((Join-Path $dirData "$dateSeance.json"), ($donnees | ConvertTo-Json -Depth 5), $utf8)
}

# Page d'archives
$liste = New-Object System.Text.StringBuilder
foreach ($f in (Get-ChildItem $dirRapports -Filter '*.html' | Sort-Object Name -Descending)) {
  $d = $f.BaseName
  $info = ''
  $jf = Join-Path $dirData "$d.json"
  if (Test-Path $jf) {
    try { $jd = Get-Content $jf -Raw -Encoding UTF8 | ConvertFrom-Json
          if ($jd.indice) { $info = "<span class='num' style='color:var(--muted);font-size:14px'>$nomIndice $(Num $jd.indice.cloture) &nbsp;$(Pastille $jd.indice.variation)</span>" } } catch {}
  }
  $lib = Majuscule ([datetime]::ParseExact($d, 'yyyy-MM-dd', $inv).ToString('dddd d MMMM yyyy', $fr))
  [void]$liste.Append("<li><a href='rapports/$d.html'><b>$(Enc $lib)</b>$info</a></li>")
}
$corpsArch = "<section class='hero' style='padding-bottom:10px'><div class='wrap' style='display:block'><span class='chip'><i></i>Historique</span><h1 style='margin-bottom:0'>Archives</h1></div></section><main class='wrap'><div class='panel glass' style='margin-top:24px'><ul class='arch'>$($liste.ToString())</ul></div></main>"
[IO.File]::WriteAllText((Join-Path $OutDir 'archives.html'), (Page "Archives — $($config.titre)" '' $corpsArch $false), $utf8)

Write-Host ""
Write-Host "Compte-rendu du $dateSeance généré ($total actions, $($favoris.Count) favoris, $($cryptos.Count) cryptos$(if ($provisoire) {', provisoire'})) dans $OutDir"

# ---------------------------------------------------------------- e-mail du soir
# Mise en page en tableaux et styles en ligne : c'est ce que les applis mail (Gmail, Outlook, iPhone) affichent correctement.
function CouleurMail($v) { switch (Signe $v) { 'up' { '#2ef5a2' } 'down' { '#ff4d6d' } default { '#9aa3bf' } } }
function PctMail($v) { "<span style='color:$(CouleurMail $v);font-weight:700;white-space:nowrap'>$(Pct $v)</span>" }
function LigneMail([string]$nom, [string]$sous, [string]$prix, $v) {
  "<tr><td style='padding:11px 0;border-bottom:1px solid #20263d'><div style='font-weight:600;color:#eaf0ff;font-size:15px'>$nom</div><div style='font-size:12px;color:#8b95b4;margin-top:2px'>$sous</div></td>" +
  "<td align='right' style='padding:11px 0;border-bottom:1px solid #20263d;white-space:nowrap'><div style='font-family:Consolas,Menlo,monospace;color:#eaf0ff;font-size:14px'>$prix</div><div style='font-family:Consolas,Menlo,monospace;font-size:14px'>$(PctMail $v)</div></td></tr>"
}
function BlocMail([string]$titre, [string]$contenu) {
  "<tr><td style='padding:0 0 18px'><table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='background:#11152a;border:1px solid #232a45;border-radius:16px'><tr><td style='padding:18px 20px'>" +
  "<div style='font-size:12px;letter-spacing:2px;text-transform:uppercase;color:#7dd3fc;font-weight:700;margin-bottom:8px'>$titre</div>$contenu</td></tr></table></td></tr>"
}

$blocs = New-Object System.Text.StringBuilder

# En-tête + indice
$entete = "<div style='font-size:26px;font-weight:700;color:#ffffff;letter-spacing:-0.5px;line-height:1.2'>$(Enc $dateLongue)</div>"
if ($aIndice) {
  $entete += "<div style='margin-top:14px'><span style='font-size:13px;color:#8b95b4;letter-spacing:1px;text-transform:uppercase'>$(Enc $nomIndice)</span><br><span style='font-family:Consolas,Menlo,monospace;font-size:30px;font-weight:700;color:#ffffff'>$(Num $aIndice.Cloture)</span>&nbsp;&nbsp;<span style='font-family:Consolas,Menlo,monospace;font-size:18px'>$(PctMail $aIndice.Var)</span></div>"
}
$entete += "<p style='margin:14px 0 0;font-size:15px;line-height:1.6;color:#c3cbe2'>$(Enc ($synthese -join ' '))</p>"
[void]$blocs.Append((BlocMail 'Compte-rendu de clôture' $entete))

# Favoris, avec leur commentaire
if ($favoris.Count) {
  $f = New-Object System.Text.StringBuilder
  foreach ($x in $favoris) {
    $com = Get-Commentaire $x.Nom $x.A 0 0 $x.RefVar $x.RefNom $x.Dev
    if ($x.Provisoire) { $com = @($com) + 'Séance encore en cours : chiffres provisoires.' }
    [void]$f.Append("<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='border-bottom:1px solid #20263d'><tr>" +
      "<td style='padding:12px 0 4px'><div style='font-weight:700;color:#fbbf24;font-size:16px'>★ $(Enc $x.Nom)</div><div style='font-size:12px;color:#8b95b4'>$(Enc $x.Symbole)</div></td>" +
      "<td align='right' style='padding:12px 0 4px;white-space:nowrap'><div style='font-family:Consolas,Menlo,monospace;color:#eaf0ff;font-size:15px'>$(Prix $x.A.Cloture) $($x.Dev)</div><div style='font-family:Consolas,Menlo,monospace;font-size:15px'>$(PctMail $x.A.Var)</div></td></tr>" +
      "<tr><td colspan='2' style='padding:2px 0 12px;font-size:13.5px;line-height:1.55;color:#aeb7d2'>$(Enc ($com -join ' '))<br><span style='color:#8b95b4'>1 semaine : </span>$(PctMail $x.A.Perf5) &nbsp; <span style='color:#8b95b4'>Depuis janv. : </span>$(PctMail $x.A.PerfYtd)</td></tr></table>")
  }
  [void]$blocs.Append((BlocMail 'Mes favoris' $f.ToString()))
}

# Cryptos
if ($cryptos.Count) {
  $rows = ($cryptos | Sort-Object { $_.A.Var } -Descending | ForEach-Object { LigneMail (Enc $_.Nom) "$($_.Code) · 7 jours : $(Pct $_.A.Perf5)" "$(Prix $_.A.Cloture) $($_.Dev)" $_.A.Var }) -join ''
  [void]$blocs.Append((BlocMail 'Cryptomonnaies · 24 h' "<table role='presentation' width='100%' cellpadding='0' cellspacing='0'>$rows</table>"))
}

# Palmarès
$rowsH = (@($resultats | Select-Object -First 5) | ForEach-Object { LigneMail (Enc $_.Nom) (Enc $_.Secteur) "$(Prix $_.A.Cloture) $($_.Dev)" $_.A.Var }) -join ''
$rowsB = (@($resultats | Select-Object -Last 5 | Sort-Object { $_.A.Var }) | ForEach-Object { LigneMail (Enc $_.Nom) (Enc $_.Secteur) "$(Prix $_.A.Cloture) $($_.Dev)" $_.A.Var }) -join ''
[void]$blocs.Append((BlocMail 'Plus fortes hausses' "<table role='presentation' width='100%' cellpadding='0' cellspacing='0'>$rowsH</table>"))
[void]$blocs.Append((BlocMail 'Plus fortes baisses' "<table role='presentation' width='100%' cellpadding='0' cellspacing='0'>$rowsB</table>"))

$mailHtml = @"
<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="color-scheme" content="dark"><title>$(Enc $titrePage)</title></head>
<body style="margin:0;padding:0;background:#05060c">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#05060c"><tr><td align="center" style="padding:20px 12px">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;font-family:-apple-system,'Segoe UI',Roboto,Helvetica,Arial,sans-serif">
<tr><td style="padding:4px 4px 18px"><table role="presentation" cellpadding="0" cellspacing="0"><tr>
<td style="width:34px;height:34px;border-radius:10px;background:#22d3ee;background-image:linear-gradient(135deg,#22d3ee,#8b5cf6);text-align:center;color:#fff;font-size:18px;font-weight:700">↗</td>
<td style="padding-left:10px;color:#eaf0ff;font-size:17px;font-weight:700">$(Enc $config.titre) <span style="color:#8b95b4;font-weight:500">· le compte-rendu du soir</span></td></tr></table></td></tr>
$($blocs.ToString())
<tr><td style="padding:6px 4px 20px;font-size:11.5px;line-height:1.5;color:#5d6684">$majTxt · Données Yahoo Finance, susceptibles d'être différées. Commentaires générés automatiquement à partir des cours, à titre informatif : ce n'est pas un conseil en investissement.<br>Le site complet (graphiques, suivi détaillé) est dans le dossier Bourse-compte-rendu\site de votre ordinateur.</td></tr>
</table></td></tr></table></body></html>
"@
[IO.File]::WriteAllText((Join-Path $OutDir 'apercu-mail.html'), $mailHtml, $utf8)

# Objet : l'essentiel lisible directement dans la notification
$morceaux = New-Object System.Collections.Generic.List[string]
if ($aIndice) { $morceaux.Add("$nomIndice $(Pct $aIndice.Var)") }
foreach ($x in $favoris) { $morceaux.Add("$($x.Nom) $(Pct $x.A.Var)") }
if ($btc) { $morceaux.Add("Bitcoin $(Pct $btc.A.Var)") }
$sujet = "Clôture du $($dateObj.ToString('dd/MM', $fr)) · " + ($morceaux -join ' · ')

if ($Mail) {
  $journal = Join-Path $PSScriptRoot 'journal-mail.txt'
  $note = { param($t) Add-Content -Path $journal -Value "$((Get-Date).ToString('yyyy-MM-dd HH:mm')) $t" -Encoding UTF8; Write-Host $t }
  $aujourdhui = $heureParis.ToString('yyyy-MM-dd')
  $cfgMail = Join-Path $PSScriptRoot 'notification.json'
  $secret  = Join-Path $PSScriptRoot 'mail.secret'
  if ($dateSeance -ne $aujourdhui -and -not $ForcerMail) {
    & $note "Pas de séance aujourd'hui (dernière : $dateSeance) : aucun e-mail envoyé."
  } elseif (-not (Test-Path $cfgMail) -or -not (Test-Path $secret)) {
    & $note "E-mail non configuré : lancez configurer-mail.ps1."
  } else {
    $cm  = Get-Content $cfgMail -Raw -Encoding UTF8 | ConvertFrom-Json
    $msg = New-Object System.Net.Mail.MailMessage
    $smtp = $null
    try {
      $mdp = Get-Content $secret -Raw | ForEach-Object { $_.Trim() } | ConvertTo-SecureString
      $msg.From = New-Object System.Net.Mail.MailAddress($cm.expediteur, $config.titre)
      foreach ($d in @($cm.destinataires)) { $msg.To.Add($d) }
      $msg.Subject = $sujet; $msg.SubjectEncoding = [Text.Encoding]::UTF8
      $msg.Body = $mailHtml; $msg.IsBodyHtml = $true; $msg.BodyEncoding = [Text.Encoding]::UTF8
      $smtp = New-Object System.Net.Mail.SmtpClient($cm.smtp, [int]$cm.port)
      $smtp.EnableSsl = $true
      $smtp.Credentials = New-Object System.Net.NetworkCredential($cm.expediteur, $mdp)
      $smtp.Send($msg)
      & $note "E-mail envoyé : $sujet"
    } catch {
      & $note "Échec de l'envoi de l'e-mail : $($_.Exception.GetBaseException().Message)"
    } finally {
      $msg.Dispose(); if ($smtp) { $smtp.Dispose() }
    }
  }
}
