$pythonCandidates = @(
    "C:\Users\kumax\AppData\Local\Programs\Python\Python313\python.exe",
    "C:\Users\kumax\AppData\Local\Programs\Python\Python312\python.exe",
    "C:\Users\kumax\AppData\Local\Programs\Python\Launcher\py.exe"
)

$python = $pythonCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
$scriptPath = Join-Path $PSScriptRoot "send_mails.py"
if (-not $python) {
    throw "Python not found."
}

& $python $scriptPath @args
