cd ..
rmdir /s /q .venv
.\setup\python-3.14.5-amd64.exe /quiet PrependPath=1 InstallAllUsers=1
python -m venv .venv
.venv\Scripts\pip install -r .\setup\requirements.txt
