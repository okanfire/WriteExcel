function WriteExcel{
    $PSpath = Get-Location | ForEach-Object Path
    
    py -3 WriteExcel.py  -ArgumentList $PSpath -Wait
    #Start-Process -FilePath C:\Users\okanfire\Desktop\code\dist\WriteExcel.exe -ArgumentList $PSpath -Wait
}