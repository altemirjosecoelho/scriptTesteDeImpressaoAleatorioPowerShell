. .\config.ps1

function sound {
    $soundPlayer = New-Object System.Media.SoundPlayer
    $soundPlayer.SoundLocation = ".\sound.wav"
    $soundPlayer.Play()
}

function cleanerJobs {
    Write-Host "Para o servico de spooler de impressao"
    Stop-Service -Name "Spooler" -Force

    Write-Host "Limpa o spooler de impressao"
    Remove-Item -Path "$env:SystemRoot\System32\spool\PRINTERS\*" -Force

    Write-Host "Inicia o servico de spooler de impressao"
    Start-Service -Name "Spooler"
}

try {
    $i = 1
    for (; $i -le $numeroDeExecucoes; $i++) {   
        (New-Object -ComObject WScript.Network).SetDefaultPrinter($nomeDaImpressora)
        Start-Process -FilePath ".\teste.docx" -Verb print -PassThru -Wait
        sound
        Write-Output "Imprimindo na impressora $nomeDaImpressora pagina: $i"
        Start-Sleep -s $tempoDeIntervaloEmSegundos
    }
}
catch {
    Write-Host "Ocorreu um erro: $_"
}
finally {
    cleanerJobs
}

