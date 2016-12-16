##################################################
# Copy .xlsx files from a source folder into a   #
# destination folder, turning them into .csv     #
#                                                #
# R. Gold, robing@pscleanair.org, 12/13/16       #
##################################################

$Source = read-host -prompt 'Enter source folder path'
$Dest = read-host -prompt 'Enter destination folder path'

Function ExcelCSV ($File)
{
    $excelFile = "$Source\" + $File + ".xlsx"
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $wb = $Excel.Workbooks.Open($excelFile)
    foreach ($ws in $wb.Worksheets)
    {
        $ws.SaveAs("$Dest\" + $File + ".csv", 6)
    }
    $Excel.Quit()
}

ExcelCSV -File *
