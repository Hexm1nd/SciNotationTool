REM if exist resources_rc.py del resources_rc.py
if exist ui_MainWindow.py del ui_MainWindow.py

set root=C:\ProgramData\Miniconda3
call %root%\Scripts\activate.bat %root%
REM call pyrcc5 resources.qrc -o resources_rc.py
call pyuic5 MainWindow.ui -o ui_MainWindow.py
