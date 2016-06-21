CLS

set app_name=OrderManager
set main_dir=%~dp0
set pyinstaller_dir=S:\Everyone\Management Software\pyinstaller-2.0\utils
set build_options=
REM set python=C:\Python27\python.exe
set python=python



REM set makespec_options=--console
REM %python% "%pyinstaller_dir%\Makespec.py" %makespec_options% "%main_dir%%app_name%.py"

@echo -----------------------------------

%python% -O "%pyinstaller_dir%\Build.py" %build_options% "%main_dir%%app_name%.spec"

@echo -----------------------------------


REM MOVE /Y "%main_dir%\dist\%app_name%" "%main_dir%"

XCOPY /Y /E "%main_dir%dist\%app_name%" "%main_dir%%app_name%"

rmdir /s /q "%main_dir%\build"
rmdir /s /q "%main_dir%\dist"


@PAUSE
