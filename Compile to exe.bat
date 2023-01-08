set root=C:\ProgramData\Miniconda3
call %root%\Scripts\activate.bat %root%
call activate etosci

nuitka --onefile --follow-imports --enable-plugin=pyqt5 --windows-disable-console --remove-output -o SciNotationTool.exe SciNotationTool.py