if exist venv rmdir /s /q venv

python -m venv venv
call venv\Scripts\activate.bat
pip install -r requirements.txt

pause