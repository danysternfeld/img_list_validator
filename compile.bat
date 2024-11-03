pyinstaller -F img_list_validator.py
copy dist\*.exe install\
cd install
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" ./installerScript.iss
del *.exe
