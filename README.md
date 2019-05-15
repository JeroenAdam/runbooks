# PowerShell snippet

// count lines/words/chars of code for one file type
```
dir .\* -rec -include *.js | gc | out-file .\combine.txt
gc .\combine.txt | where {$_ -ne ""} > .\clean.txt
gc .\clean.txt | select-string -pattern '\/\/' -notmatch | Out-File .\result.txt
gc .\result.txt | Measure-Object -line –word -character
del .\combine.txt  
del .\clean.txt 
del .\result.txt
```

//clean by removing empty lines and lines that start with #
```
gc .\toclean.txt | where {$_ -ne ""} | select-string -pattern '^#' -notmatch | Out-File .\result.txt
```
To add: AD queries, CSV parsing, Excel COM interface
