@echo off

set VENV=.winvenv
set REQUIREMENTS=requirements.txt
set PYTHON=python

REM Function to create a virtual environment and install requirements
:install
    if not exist %VENV% (
        %PYTHON% -m venv %VENV%
        call %VENV%\Scripts\activate
        pip install --upgrade pip
        pip install -r %REQUIREMENTS%
    ) else (
        call %VENV%\Scripts\activate
    )
    exit /b

REM Run all scripts related to effectivity
:run_effectivity
    call :install
    %VENV%\Scripts\python split_excel.py
    %VENV%\Scripts\python create_config.py
    %VENV%\Scripts\python mod_excel.py
    %VENV%\Scripts\python ss_uploader.py set
    %VENV%\Scripts\python ss_uploader.py update
    exit /b

REM Run the set command
:run
    call :install
    %VENV%\Scripts\python ss_uploader.py set
    exit /b

REM Run the config script
:config
    call :install
    %VENV%\Scripts\python create_config.py
    exit /b

REM Update the sheet
:update
    call :install
    %VENV%\Scripts\python ss_uploader.py update
    exit /b

REM Create a summary
:summary
    call :install
    %VENV%\Scripts\python ss_uploader.py summary
    exit /b

REM Get the sheet with a specified filepath
:get
    call :install
    if "%1"=="" (
        echo Error: No filepath specified.
    ) else (
        %VENV%\Scripts\python ss_uploader.py get %1
    )
    exit /b

REM Run tests
:test
    call :install
    %VENV%\Scripts\python ss_uploader.py test
    exit /b

REM Clean up virtual environment and directories
:clean
    rmdir /s /q %VENV%
    del /q Effectivity_Reports_Mod\*
    del /q Effectivity_Reports_Split\*
    exit /b
