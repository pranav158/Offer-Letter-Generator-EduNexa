#              $$\      $$\                       $$\     $$\           
#              $$$\    $$$ |                      $$ |    \__|          
#              $$$$\  $$$$ |$$\   $$\  $$$$$$$\ $$$$$$\   $$\  $$$$$$$\ 
#              $$\$$\$$ $$ |$$ |  $$ |$$  _____|\_$$  _|  $$ |$$  _____|
#              $$ \$$$  $$ |$$ |  $$ |\$$$$$$\    $$ |    $$ |$$ /      
#              $$ |\$  /$$ |$$ |  $$ | \____$$\   $$ |$$\ $$ |$$ |      
#              $$ | \_/ $$ |\$$$$$$$ |$$$$$$$  |  \$$$$  |$$ |\$$$$$$$\ 
#              \__|     \__| \____$$ |\_______/    \____/ \__| \_______|
#                           $$\   $$ |                                  
#                           \$$$$$$  |                                  
#                            \______/               
# 


$curr_path = Split-Path -parent $MyInvocation.MyCommand.Path
$word_app = New-Object -ComObject Word.Application

Get-ChildItem -Path $curr_path -Recurse -Filter *.doc? | ForEach-Object {
    Write-Host "Processing" $_.FullName "..."
    $document = $word_app.Documents.Open($_.FullName)
    $pdf_filename = "$($curr_path)\$($_.BaseName).pdf"
    $document.SaveAs([ref] $pdf_filename, [ref] 17)
    $document.Close()
}

$word_app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word_app)
