$env:VENV=".venv"
$env:OLDPATH=$env:PATH
$env:PYENVPATH=PWD
$env:PYENVPATH+="\"
$env:PYENVPATH+=$env:VENV
$env:PYENVPATH+="\Scripts\"
$env:PATH="$env:PYENVPATH;$env:PATH"