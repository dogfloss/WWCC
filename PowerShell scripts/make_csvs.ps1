##################################################
# Copy .xlsx files from a source folder into a   #
# destination folder, turning them into .csv     #
#                                                #
# R. Gold, robing@pscleanair.org, 12/9/16        #
##################################################

$Source = read-host -prompt 'Enter source folder path'
$Dest = read-host -prompt 'Enter destination folder path'
mkdir Tmp222
cp $Source\*.xlsx Tmp222\.
cd Tmp222
Dir | Rename-Item -newname { $_.name -replace ".xlsx",".csv" }
cp *.csv ..\$Dest\.
cd ..
Remove-Item -Recurse Tmp222
