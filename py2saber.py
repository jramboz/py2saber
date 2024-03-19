# py2saber Copyright © 2023 Jason Ramboz
#
# This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License along with this program. If not, see <https://www.gnu.org/licenses/>.

'''
Python clone of Nuntis' sendtosaber command, with a few extensions.

Original sendtosaber source: https://github.com/Nuntis-Spayz/Send-To-Saber
Polaris Anima EVO Comms Protocol: https://github.com/LamaDiLuce/polaris-opencore/blob/master/Documentation/COMMS-PROTOCOL.md
'''

import serial
import serial.tools.list_ports as lp
import platform
import re
import os
import logging
import sys
import argparse
import errno
from getch import pause_exit
import glob
import time

script_version = '0.12b'
script_authors = 'Jason Ramboz'
script_repo = 'https://github.com/jramboz/py2saber'



# Custom Exceptions
class NoAnimaSaberException(Exception):
    pass

class AnimaNotReadyException(Exception):
    pass

class NotEnoughFreeSpaceException(Exception):
    pass

class AnimaFileWriteException(Exception):
    pass

class InvalidSaberResponseException(Exception):
    pass

class Saber_Controller:
    '''Controls communication with an OpenCore-based lightsaber.'''
    # Serial port communication settings
    _SERIAL_SETTINGS = {'baudrate': 115200, 
                        'bytesize': 8, 
                        'parity': 'N', 
                        'stopbits': 1, 
                        'xonxoff': False, 
                        'dsrdtr': False, 
                        'rtscts': False, 
                        'timeout': 3, 
                        'write_timeout': None, 
                        'inter_byte_timeout': None}

    def __init__(self, port: str=None, gui: bool = False, loglevel: int = logging.ERROR) -> None:
        self.log = logging.getLogger('Saber_Controller')
        self.log.setLevel(loglevel)
        if not self.log.hasHandlers():
            stream = logging.StreamHandler(sys.stdout)
            #stream.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s: %(message)s'))
            stream.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))
            self.log.addHandler(stream)
        self.log.info('Initializing saber connection.')
        self.gui = gui # Flag for whether to output signals for PySide GUI
        self.port = port

        # If a specific port is supplied, check that it is an OpenCore saber
        if self.port: 
            if not Saber_Controller.port_is_anima(port):
                self.log.error(f'No OpenCore saber found on port {port}')
                raise NoAnimaSaberException
        # Otherwise, use the first port found with an OpenCore saber connected
        else:
            ports = Saber_Controller.get_ports()
            for p in ports:
                if Saber_Controller.port_is_anima(p):
                    self.port = p
                    break

        if not self.port:
            self.log.error('No OpenCore sabers found.')
            raise NoAnimaSaberException
        
        # Initialize the serial connection
        self._ser = serial.Serial(self.port)
        self._ser.apply_settings(self._SERIAL_SETTINGS)

    def __del__(self):
        try: # Exception handling is necessary for the case that no saber was found during initialization. In this case, self._ser never gets created
            self._ser.close()
        except Exception:
            pass

    @staticmethod
    def get_ports() -> list[str]:
        '''Returns available serial ports as list of strings.'''
        serial_ports = []
        match_string = r''

        _log = logging.getLogger('Saber_Controller')

        _log.info('Searching for available serial ports.')
        port_list = lp.comports()
        _log.debug(f'Found {len(port_list)} ports before filtering: {[port.device for port in port_list]}')

        # Filter down list of ports depending on OS
        system = platform.system()
        _log.info(f'Detected OS: {system}')
        if system == 'Darwin':
            match_string = r'^/dev/cu.usb*'
        elif system == 'Windows':
            match_string = r'^COM\d+'
        elif system == 'Linux':
            match_string = r'^/dev/ttyS*'
        else: # You're on your own!
            match_string = r'*'
        
        for port in port_list:
            if re.match(match_string, port.device):
                serial_ports.append(port.device)
        
        _log.info(f'Found {len(serial_ports)} port(s).')
        _log.debug(f'Found ports: {serial_ports}')
        return serial_ports

    @staticmethod
    def port_is_anima(port: str) -> bool:
        '''Checks to see if the device attached to port is a Polaris Anima EVO.
        Returns True if an Anima is found, False otherwise.
        
        NB: This method will not throw exceptions. Any exceptions will cause result of False.'''

        try:
            log = logging.getLogger('Saber_Controller')

            log.info(f'Checking if decvice on port {port} is a Polaris Anima EVO.')
            ser = serial.Serial(port)
            ser.apply_settings({'baudrate': 115200, 
                           'bytesize': 8, 
                           'parity': 'N', 
                           'stopbits': 1, 
                           'xonxoff': False, 
                           'dsrdtr': False, 
                           'rtscts': False, 
                           'timeout': 3, 
                           'write_timeout': None, 
                           'inter_byte_timeout': None})
            
            # Checking logic (based on Nuntis' script):
            # Send 'V?'. Return false if no response or respond with a 1.x version. Otherwise continue.
            # Send 'S?'. Return false if no response or response doesn't start with 'S='. Otherwise continue.
            # Send 'WR?'. Return false if empty or response doesn't sart with 'OK, Write'.
            # Otherwise return true.
            log.debug('Sending command: V?')
            ser.write(b'V?\n')
            response = ser.readline()
            log.debug(f'Received response: {response}')
            if not response or response.startswith(b'V=1.'):
                ser.close()
                log.info(f'No Polaris Anima EVO found on port {port}')
                return False
            
            log.debug('Sending command: S?')
            ser.write(b'S?\n')
            response = ser.readline()
            log.debug(f'Received response: {response}')
            if not response or not response.startswith(b'S='):
                ser.close()
                log.info(f'No Polaris Anima EVO found on port {port}')
                return False
            
            log.debug('Sending command: WR?')
            ser.write(b'WR?\n')
            response = ser.readline()
            log.debug(f'Received response: {response}')
            if not response or not response.startswith(b'OK, Write'):
                ser.close()
                log.info(f'No Polaris Anima EVO found on port {port}')
                return False

            ser.close()
            log.info(f'Found Polaris Anima EVO on port {port}')
            return True
        except:
            log.info(f'No Polaris Anima EVO found on port {port}')
            return False

    def send_command(self, cmd: bytes) -> None:
        '''Send command string to attached saber. It will automatically add b'\n' terminator if not already present.
        
        Note: this method does not check that the saber is write-ready.'''
        if not cmd.endswith(b'\n'):
            cmd += b'\n'
        self.log.debug(f'Sending command to saber: {cmd}')
        self._ser.write(cmd)
    
    def read_line(self) -> bytes:
        '''Reads the next line (terminated by b'\n') from the serial buffer.
        
        Note: this removes the line from the buffer.'''
        response = self._ser.readline()
        self.log.debug(f'Received response: {response}')
        # Sometimes it seems the read request comes too fast, before the anima has had a chance to respond.
        # So now, if I get a blank response, I wait and try again until a response is received or timeout.
        if not response:
            tries = 1
            max_tries = 5
            while not response and tries < max_tries:
                time.sleep(0.5)
                response = self._ser.readline()
                self.log.debug(f'Received response: {response}')
                tries += 1
        return response

    def saber_is_ready(self) -> bool:
        '''Checks to see if saber is ready to receive commands.
        
        NB: This method will not throw exceptions. Any exceptions will cause result of False.'''
        try:
            self.send_command(b'WR?')
            response = self.read_line()
            if response == b'OK, Write Ready\n':
                self.log.debug('Saber is ready to receive commands.')
                return True
            self.log.error('Saber is not ready to receive commands.')
            return False
        except Exception:
            self.log.error('Saber is not ready to receive commands.')
            return False

    def get_saber_info(self) -> dict:
        '''Retrieve firmware version and serial number from saber. Returns a dict with keys 'version' and 'serial'.'''
        if self.saber_is_ready():
            self.log.info('Retrieving firmware version and serial number from saber.')
            info = {}
            
            # get firmware version
            cmd = b'V?'
            self.send_command(cmd)
            response = self.read_line()
            if not response or not response.startswith(b'V='):
                self.log.error(f'Invalid response received.\nCommand: {cmd}\nResponse: {response}')
                raise InvalidSaberResponseException
            else:
                info['version'] = response.decode().strip()[2:]

            # get serial number
            cmd = b'S?'
            self.send_command(cmd)
            response = self.read_line()
            if not response or not response.startswith(b'S='):
                self.log.error(f'Invalid response received.\nCommand: {cmd}\nResponse: {response}')
                raise InvalidSaberResponseException
            else:
                info['serial'] = response.decode().strip()[2:]
            
            self.log.info(f'Found saber info: {info}')
            return info
        else:
            raise AnimaNotReadyException

    def list_files_on_saber_as_bytes(self) -> bytes:
        '''Returns the raw byte string reported by the saber LIST? command.'''
        file_list = b''
        self.log.info(f'Retrieving file list from saber.')

        if self.saber_is_ready():
            cmd = b'LIST?'
            self.send_command(cmd)
            response = b''

            while not response.endswith(b'\x03\n'):
                response = self.read_line()
                file_list += response
        
            self.log.debug(f'Final byte string: {file_list}')
            return file_list
        else:
            raise AnimaNotReadyException
    
    def list_files_on_saber(self) -> dict[str: int]:
        '''Returns a dictionary containing the files on the saber. Key is the filename, value is the file size in bytes.'''
        file_dict = {}
        file_list = self.list_files_on_saber_as_bytes().decode()
        regex = re.compile(r"^(\w+\.RAW)\s+(\d+)$", re.MULTILINE)
        matches = regex.findall(file_list)
        for match in matches:
            file_dict[match[0]] = int(match[1])
        return file_dict

    def erase_all_files(self, progress_callback: callable = None) -> None:
        '''Erases all files on the anima. USE CAREFULLY.'''
        if self.saber_is_ready():
            self.log.info('Erasing all files on saber. This may take several minutes.')
            cmd = b'ERASE=ALL'
            self.send_command(cmd)
            
            # Listen to output stream to make sure process is ongoing/completed
            response = self.read_line() # "Erasing Serial Flash, this may take 20s to 2 minutes\n"
            self._ser.timeout = 5 # increase timeout while we wait for #s to display
            i = 0
            c = self._ser.read(1)
            while c < b'\x41': # any ascii character before 'A'
                if self.gui:
                    i += 1
                    progress_callback.emit(i*100//140)
                print(c.decode(), end='', flush=True)
                c = self._ser.read(1)
            # manually read the next line using the serial connection and combine it with the last char read.
            response = c + self._ser.readline()
            self.log.debug(f'Received response: {response}') # b'OK, Now re-load your sound files.\n'
            response = self.read_line() # b'OK, Serial Flash Erased.\n'
            response = self.read_line() # b'\n'
            self._ser.timeout = self._SERIAL_SETTINGS['timeout']
        else:
            raise AnimaNotReadyException

    def get_free_space(self) -> int:
        '''Returns the amount of free space in Anima storage in bytes.'''
        self.log.info('Getting free space on Anima.')
        cmd = b'FREE?'
        self.send_command(cmd)
        response = self.read_line()
        r = response.decode().strip()
        free_space = int(r[5:])
        self.log.info(f'Free space: {free_space} bytes')
        return free_space

    def read_config_ini(self) -> str:
        '''Read the config.ini file from saber and return as a string'''
        self.log.info('Reading config.ini from saber')
        cmd = b'RD?config.ini\n'
        self.send_command(cmd)
        # Anima doesn't seem to properly sent STX/ETX bytes. Instead just sends '2' and '3'
        config = b''
        while not config.endswith(b'}3'):
            # read one byte at a time, since last line isn't terminated with \n
            config += self._ser.read()

        self.log.debug(f'Raw config string: {config}')
        # slice off first and last char ('2' and '3')
        return config.decode().strip()[1:-1]

    def write_files_to_saber(self, files: list[str], progress_callback: callable = None) -> None:
        '''Write file(s) to saber. Expects a list of file names.

        NB: This method does no checking that files exist either on disk or saber. Please verify files before calling this method.'''
        
        if self.saber_is_ready():
            files.sort()
            self.log.info(f'Preparing to write file(s) to saber: {files}')
            for file in files:
                self.log.info(f'Writing file to saber: {file}')
                
                # Check for enough free space
                file_size = os.path.getsize(file)
                self.log.debug(f'File size: {file_size}')
                free_space = self.get_free_space()
                if free_space < file_size:
                    self.log.error(f'Not enough free space on saber for file {file}')
                    raise NotEnoughFreeSpaceException

                # Write the file
                with open(file, mode='rb') as binary_file:
                    bytes_sent = 0
                    fname = os.path.basename(file)
                    report_every_n_bytes = 512 # How often to update the bytes sent display
                    system = platform.system()

                    cmd = b'WR=' + fname.encode('utf-8') + b', ' + str(file_size).encode('utf-8') + b'\n'
                    self.send_command(cmd)
                    response = self.read_line()
                    self.log.debug(f'Beginning byte stream for file {fname}')
                    byte = binary_file.read(1)
                    print(f'{fname} - Bytes sent: {bytes_sent} - Bytes remaining: {file_size - bytes_sent}', end='', flush=True)
                    while byte:
                        self._ser.write(byte)
                        if system != 'Windows':
                            time.sleep(0.00009) # otherwise it sends too fast on mac (and linux?)
                        bytes_sent += 1
                        byte = binary_file.read(1)
                        if bytes_sent % report_every_n_bytes == 0:
                            print(f'\r{fname} - Bytes sent: {bytes_sent} - Bytes remaining: {file_size - bytes_sent}', end='', flush=True)
                            if self.gui:
                                progress_callback.emit(bytes_sent)
                    print(f'\r{fname} - Bytes sent: {bytes_sent} - Bytes remaining: {file_size - bytes_sent}     ')
                
                response = self.read_line()
                if not response == b'OK, Write Complete\n':
                    self.log.error(f'Error writing file to saber. Error message: {response}')
                    raise AnimaFileWriteException
            
                self.log.info(f'Successfully wrote file to saber: {file}')
                time.sleep(1)
        else:
            raise AnimaNotReadyException

    @staticmethod
    def rgbw_to_byte_str(r: int, g: int, b:int, w: int):
        return bytes(str(r), 'ascii') + b',' + bytes(str(g), 'ascii') + b',' + bytes(str(b), 'ascii') + b','+ bytes(str(w), 'ascii')

    def preview_color(self, r: int, g: int, b: int, w: int):
        cmd = b'P=' + self.rgbw_to_byte_str(r, g, b, w) + b'\n'
        self.send_command(cmd)
        response = self.read_line()
        if response[:2] != b'OK':
            self.log.error(f'Invalid response received.\nCommand: {cmd}\nResponse: {response}')
            raise InvalidSaberResponseException
    
    def set_color(self, bank: int, effect: str, r: int, g: int, b: int, w:int):
        '''Writes a color setting to the saber. 
        bank specifies which bank (0-7). 
        effect is one of "color", "clash", or "swing". 
        RGBW values are 0-255'''
        match effect:
            case "color":
                cmd = b'C'
            case "clash":
                cmd = b'F'
            case "swing":
                cmd = b'W'
            case _:
                self.log.error(f'Invalid effect type specified: {effect}')
                return
        
        cmd = cmd + bytes(str(bank),'ascii') + b'=' + self.rgbw_to_byte_str(r, g, b, w) + b'\n'
        self.send_command(cmd)
        response = self.read_line()
        if response[:2] != b'OK':
            self.log.error(f'Invalid response received.\nCommand: {cmd}\nResponse: {response}')
            raise InvalidSaberResponseException
    
    def set_active_bank(self, bank:int):
        '''Sets the active bank (0-7)'''
        cmd = b'B=' + bytes(str(bank), 'ascii') + b'\n'
        self.send_command(cmd)
        response = self.read_line()
        if response[:2] != b'OK':
            self.log.error(f'Invalid response received.\nCommand: {cmd}\nResponse: {response}')
            raise InvalidSaberResponseException
    
# ---------------------------------------------------------------------- #
# Command Line Operations                                                #
# ---------------------------------------------------------------------- #

def main_func():
    log = logging.getLogger()
    if not log.hasHandlers():
        log.setLevel(logging.ERROR)
        stream = logging.StreamHandler(sys.stdout)
        stream.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))
        log.addHandler(stream)

    exit_code = 0

    # configure command line parser and options
    parser = argparse.ArgumentParser(prog='py2saber',
                                     description='A utility for working with OpenCore-based sabers, based on "sendtosaber" by Nuntis')
    parser.add_argument('-v', '--version',
                        action="version",
                        help='display version and author information, then exit',
                        version='%(prog)s v{ver} - Author(s): {auth} - {page}'.format(ver=script_version, auth=script_authors, page=script_repo))
    parser.add_argument('-i', '--info', 
                        action="store_true",
                        help='read and display saber firmware version and serial number')
    parser.add_argument('-l', '--list',
                        action="store_true",
                        help='list all files on saber')
    parser.add_argument('files', nargs="*",
                        help='one or more files to upload to saber (separated by spaces)')
    exit_behavior = parser.add_mutually_exclusive_group()
    exit_behavior.add_argument('-s', '--silent', 
                               action="store_true",
                               help='exit without waiting for keypress (default)')
    exit_behavior.add_argument('-w', '--wait', 
                               action="store_true",
                               help='wait for keypress before exiting')
    parser.add_argument('-c', '--continue-on-file-not-found',
                        action="store_true",
                        help='if one or more specified files do not exist, continue processing the remaining files (otherwise program will exit)')
    parser.add_argument('-D', '--debug',
                        action="store_true",
                        help='Show debugging information')
    parser.add_argument('--erase-all', 
                        action="store_true",
                        help='erase all files on saber')
    parser.add_argument('--config',
                        action="store_true",
                        help='Display config.ini from saber')
    
    args = parser.parse_args(args=None if sys.argv[1:] else ['--help'])

    if args.debug:
        log.setLevel(logging.DEBUG)

    # Find port that Anima is connected to
    try:
        print('Searching for OpenCore saber.')
        sc = Saber_Controller(loglevel=logging.DEBUG if args.debug else logging.ERROR)
        print(f'OpenCore saber found on port {sc.port}')
        
        # Execute functions based on arguments provided
        if args.info: # display saber information
            print('\nRetrieving saber information')
            inf = sc.get_saber_info()
            print(f'Firmware version:\tv{inf["version"]}\nSerial Number:\t\t{inf["serial"]}')
        
        if args.list: # list files on saber
            print('\nRetrieving list of files on saber.')
            file_list = sc.list_files_on_saber_as_bytes()
            print(file_list.decode().strip())
        
        if args.config:
            print('\nRetrieving config.ini from saber')
            config = sc.read_config_ini()
            print('Config.ini:\n')
            print(config)

        if args.erase_all: # erase all files on saber
            print('\n*** This will erase ALL files on the saber! ***')
            yorn = input('Do you want to continue? (Y/N): ')
            if yorn.lower() == 'y' or yorn.lower() == 'yes':
                # do the thing
                print('Erasing all files on saber. This may take several minutes.')
                sc.erase_all_files()
                print("\nAll sound files on saber have been erased. Please re-load your sound files.")
            else:
                print('Aborting saber erase command.')
                sys.exit(1)

        if args.files: # write file(s) to saber
            # for Windows, we need to manually expand any wildcards in the input list
            if platform.system() == 'Windows':
                log.info('Windows system detected. Expanding any wildcards in file names.')
                expanded_files = []
                for file in args.files:
                    expanded_files.extend(glob.glob(file))
                args.files = expanded_files
                log.debug(f'Expanded files: {args.files}')

            print(f'\nPreparing to upload file{"s" if len(args.files)>1 else ""} to saber.')
            
            # verify that files exist
            verified_files = []
            for file in args.files:
                try:
                    if not os.path.isfile(file):
                        raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), file)
                    else:
                        verified_files.append(file)
                except FileNotFoundError as e:
                    log.error(f'File not found: {e.filename}')
                    if args.continue_on_file_not_found:
                        continue
                    else:
                        print('Aborting operation.')
                        exit_code = 1
                        sys.exit(1)
            
            # send files to saber
            try:
                sc.write_files_to_saber(verified_files)
                print(f'\nSuccessfully wrote file{"s" if len(verified_files)>1 else ""} to saber: {verified_files}')
            except AnimaFileWriteException:
                log.error('Error writing to saber. You should erase your saber and re-upload all files to avoid corrupted files.')
                exit_code = 1

    except NoAnimaSaberException as e:
        log.error('No OpenCore saber found. If the saber is connected, try restarting it with the on/off switch. If the problem persists, try a different USB cable.')
        exit_code = 1

    except AnimaNotReadyException:
        log.error('Saber is not ready to receive commands. Try restarting the saber using the on/off switch.')

    except Exception as e:
        log.error(e)
        exit_code = 1
    
    finally:
        if args.wait:
            pause_exit(exit_code, '\nPress any key to exit.')
        else:
            sys.exit(exit_code)

if __name__ == '__main__':
    main_func()