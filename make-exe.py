import PyInstaller.__main__

PyInstaller.__main__.run(
    [
        'get_bank_com_port.py',
        '--name=Get_Bank_COM_Port',
        '--onefile'
        # '--noconsole',
        # '--add-data=config-kassa.json;.',
        # '--add-data=img\kassabot.ico;img',
        # '--uac-admin',
        # '--icon=img\kassabot.ico'


    ]
)