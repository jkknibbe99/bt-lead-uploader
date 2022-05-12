@ECHO OFF

if not exist %~dp0data/init_data.json (
    echo Setting up init configuration...
    python %~dp0config.py
)

python %~dp0bot.py
