$controlName = Read-Host -Prompt 'Input control name'
New-Item $controlName -ItemType directory
cd $controlName
pac.exe pcf init --namespace OPSBase.PCFControls --name $controlName --template field
npm install
cd ..