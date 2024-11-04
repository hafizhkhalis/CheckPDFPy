Microsoft Powershell Profile

function checkpdf {
param (
[string]$scriptPath = "D:\Apps\Shell Programs\checkpdfpython\main.py"
)
python $scriptPath
}

function checkpdfpath {
param (
[string]$scriptPath = "D:\Apps\Shell Programs\checkpdfpythonwithpath\main.py"
)
python $scriptPath
}
