@ECHO OFF

if not exist %~dp0data/init_data.json (
    echo Setting up init configuration...
    python %~dp0config.py
)

rem check if a venv exists
if exist Scripts (
    echo Scripts directory exists
    echo Running bot
    call %~dp0\Scripts\activate
    rem run bot
    python %~dp0bot.py
) else (
    echo Virtual environment not created.
    echo Creating Virtual environment now...
    python -m venv %~dp0
    call %~dp0\Scripts\activate
    pip install -r requirements.txt
    rem run bot
    python %~dp0bot.py
)


