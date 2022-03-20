function WriteExcel($Sheet){
    $PSpath = Get-Location | ForEach-Object Path
    
    #py -3 WriteExcel.py $PSpath $Sheet
    Start-Process -FilePath C:\Users\okanfire\Desktop\code\WriteExcel\dist\WriteExcel.exe -ArgumentList $PSpath,$Sheet -Wait
}