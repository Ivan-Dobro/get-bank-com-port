import PyInstaller.__main__
'''
Creating .exe file for deployment to Windwos7
'''
PyInstaller.__main__.run(
    [
        'get-bank-com-port.py',
        '--name=Get-Bank-COM-Port',
        '--onefile'
    ]
)