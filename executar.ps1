. .\config.ps1
#$ScriptDir += "\config.ps1"

$nomeArquivo = "teste.docx"
$MYFILE = Join-Path $PSScriptRoot $nomeArquivo

Write-Host "Conteúdo da minhaVariavel: $nomeDaImpressora"
function sound {
    $nomeDoSom = "sound.wav"
    $caminhoDoSom = Join-Path $PSScriptRoot $nomeDoSom
    # Crie uma instância da classe System.Media.SoundPlayer
    $soundPlayer = New-Object System.Media.SoundPlayer

    # Defina o caminho do som
    $soundPlayer.SoundLocation = $caminhoDoSom

    # Reproduza o som
    $soundPlayer.Play()
}

function cleanerJobs {
    Write-Host "Para o serviço de spooler de impressao"
    Stop-Service -Name "Spooler" -Force

    Write-Host "Limpa o spooler de impressao"
    Remove-Item -Path "$env:SystemRoot\System32\spool\PRINTERS\*" -Force

    Write-Host "Inicia o serviço de spooler de impressao"
    Start-Service -Name "Spooler"
}

try {
    $i = 1
    for (; $i -le $numeroDeExecucoes; $i++) {   
        
        #$DEFAULTPRINTER = (Get-CimInstance -ClassName CIM_Printer | WHERE { $_.Default -eq $True }[0])
        #$PRINTERTMP = (Get-CimInstance -ClassName CIM_Printer | WHERE { $_.NAme -eq $nomeDaImpressora }[0])
        #$PRINTERTMP | Invoke-CimMethod -MethodName SetDefaultPrinter | Out-Null

        (New-Object -ComObject WScript.Network).SetDefaultPrinter($nomeDaImpressora)
        Start-Process -FilePath $MYFILE -Verb print -PassThru -Wait
        #$DEFAULTPRINTER | Invoke-CimMethod -MethodName SetDefaultPrinter | Out-Null

        Get-Printer -Name "$nomeDaImpressora"
        sound
        Write-Output "Imprimindo na impressora $nomeDaImpressora pagina: $i"
        Start-Sleep -s $tempoDeIntervaloEmSegundos
    }
}
catch {
    # Qualquer código para tratar exceções aqui
    Write-Host "Ocorreu um erro: $_"
}
finally {

    cleanerJobs
}

