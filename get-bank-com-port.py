'''
Get-Bank-COM-Port 
v. 1.0
21.08.23
'''

from argparse import ArgumentParser, Namespace
import logging
import logging.handlers
from pathlib import Path
import wmi,re,sys


# Default config
cfg = {
    'filename':'com-port.txt',
    'devicename':'USB Serial',
    'log-file':'get-bank-com-port.log',
    'CAPTION_DUMMY':'USB Serial Port (COM0)' # for dev enviroment
}


class Bank_COM_Port(): 
    ''' Main Class '''
    def __init__(self) -> None:
        self.filename : str
        self.devicename: str
        self.log : logging
        self.init_logging()
        self.cmd_parametrs()




    def get_cwd(self):
        ''' get program run path according to execution environment'''
        if getattr (sys,'frozen',False) and hasattr(sys,'_MEIPASS'):
            return str(Path(sys.executable).parent.absolute()) + '\\'
        else:
            return str(Path( __file__ ).parent.absolute()) + '\\'


        return str(Path( __file__ ).parent.absolute())+'\\'

    def init_logging(self):
        ''' logging to file and stduot'''
        self.log = logging.getLogger(__name__)
        self.log.setLevel(logging.DEBUG)
        log_formater = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s',
                                              datefmt='%d-%m-%Y %H:%M:%S')
        
        log_file_handler=logging.handlers.TimedRotatingFileHandler(f'{self.get_cwd()+ cfg["log-file"]}',
                                                         encoding='utf-8',
                                                         when='D',
                                                         interval=30,
                                                         backupCount=6
                                                         )
        log_file_handler.setLevel(logging.INFO)
        log_file_handler.setFormatter(log_formater)

        log_stream_handler = logging.StreamHandler()
        log_stream_handler.setFormatter(log_formater)
        log_stream_handler.setLevel(logging.DEBUG)

        self.log.addHandler(log_file_handler)
        self.log.addHandler(log_stream_handler)
        
        self.log.info('--- START ---')
        self.log.info('Loger set')
      



    def cmd_parametrs(self):
        ''' commad line parametrs'''
        cmd_parser = ArgumentParser(
            prog='Get_Bank_COM_Port.exe',
            description='Getting USB-COM port of connected payment terminal in Windows Device Manager v. 1.0 (c)2023 by iD',
            usage='get_bank_com_port [-f \"filename\"] [-n \"device name\"]',
            epilog='')

        cmd_parser.add_argument('-f','--file',
                                help=f'\"Filename\" to store COM number. Default - {cfg["filename"]}',
                                default=self.get_cwd()+cfg["filename"],
                                type=str)
        cmd_parser.add_argument('-d','--device',
                                help=f'Device name to start with. Default - \" {cfg["devicename"]} \"',
                                default=cfg['devicename'],
                                type=str)
        cmd_parser.add_argument('-t','--terminal',help='Output to terminal',action='store_true')


        args: Namespace = cmd_parser.parse_args()

        self.filename = args.file
        self.devicename = args.device +'%'


        self.log.debug(f'{self.filename=}')
        self.log.debug(f'{self.devicename=}')


    def get_com_port(self):
        ''' find device name in Device Manager and retrieve COM port number'''
        w = wmi.WMI()

        wql = f"SELECT * FROM Win32_PnPEntity WHERE Caption LIKE \'{self.devicename}\'"

        self.log.debug(f'{wql=}')

        try:
            devs = w.query(wql)
        except Exception as er:
            self.log.error(f'WMI Error query ( {wql} ) with {er}')
            devs = []

        device_caption = cfg['CAPTION_DUMMY'] # for dev enviroment

        self.log.debug(f'{devs=}')

        if len(devs) == 1 : 
            device_caption = devs[0].__getattr__('Caption') 
        elif len(devs) == 0:
            self.log.warning('No device found , using Dummy COM port 0')
        elif len(devs) > 1:
            self.log.warning(f'Found {len(devs)} devices, geting first one ')
            device_caption = devs[0].__getattr__('Caption') 

        # self.log.debug(f'{devs=}')
        self.log.info(f'Device Name Caption - {device_caption}')    
     

        match = re.search("(?<=\(COM)\d+(?=\))",device_caption) # get digit after (COM  and)
        if match:
            self.log.info(f'Port nuber is - {match[0]}')
            com_number = match[0]
        else:
            self.log.error(f'Getting re match from - {device_caption}')
            com_number = '0'

        with open(self.filename,mode='w') as f:
            try :
                f.write(com_number)
            except OSError as er:
                self.log.error(f'Writing to file {self.filename} error {er}')


if __name__ == "__main__":
    
    print(cfg['filename'])

    app = Bank_COM_Port()
    app.get_com_port()

