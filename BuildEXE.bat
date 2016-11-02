CLS

if "%~dp0" == "\\cbssrvit\kwfiles\Management Software\OrderManager\development\" (
CLS
@echo This server does not allow you to build executables directly on the server.
@echo Copy the development folder to your local machine and then build.
@PAUSE
EXIT
)


set app_name=OrderManager
set main_dir=%~dp0
set pyinstaller_dir=C:\pyinstaller-2.0\utils
set build_options=
set python=C:\Python27\python.exe


REM set makespec_options=--console
REM %python% "%pyinstaller_dir%\Makespec.py" %makespec_options% "%main_dir%%app_name%.py"


@echo -----------------------------------

%python% -O "%pyinstaller_dir%\Build.py" %build_options% "%main_dir%%app_name%.spec"

@echo -----------------------------------

@echo Finished.


@PAUSE
