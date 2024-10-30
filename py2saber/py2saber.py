# py2saber Copyright © 2023-2024 Jason Ramboz
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
import serial.rs485
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
from deprecated import deprecated

basedir = os.path.dirname(os.path.realpath(__file__))

script_version = '0.18.1'
script_authors = 'Jason Ramboz'
script_repo = 'https://github.com/jramboz/py2saber'

# adapted from https://stackoverflow.com/a/66491013
class DocDefaultException(Exception):
    """Subclass exceptions use docstring as default message"""
    def __init__(self, msg=None, *args, **kwargs):
        if msg:
            msg = self.__doc__ + '\n' + msg
        super().__init__(msg or self.__doc__, *args, **kwargs)

# Custom Exceptions
class NoAnimaSaberException(DocDefaultException):
    '''No compatible Anima found. If an Anima is connected, try restarting it with the on/off switch. If the problem persists, try a different USB cable.'''

class AnimaNotReadyException(DocDefaultException):
    '''Anima is not ready to receive commands. Try restarting the Anima using the on/off switch.'''

class NotEnoughFreeSpaceException(DocDefaultException):
    '''Not enough free space on Anima to write the requested file(s).'''

class AnimaFileWriteException(DocDefaultException):
    '''Error writing file(s) to Anima. Files are possibly corrupt. You should erase all files and re-upload.'''

class InvalidSaberResponseException(DocDefaultException):
    '''Received invalid response from saber.'''

class InvalidSoundEffectSpecifiedException(DocDefaultException):
    '''Invalid sound effect specified.'''

def getHumanReadableSize(size,precision=2):
    '''Takes a size in bytes and outputs human-readable string.'''
    # taken from https://stackoverflow.com/a/32009595
    suffixes=['B','KB','MB','GB','TB']
    suffixIndex = 0
    while size >= 1024 and suffixIndex < len(suffixes)-1:
        suffixIndex += 1 #increment the index of the suffix
        size = size/1024.0 #apply the division
    return "%.*f%s"%(precision,size,suffixes[suffixIndex])

class Saber_Controller:
    '''Controls communication with an OpenCore-based lightsaber.'''
    # Serial port communication settings
    _SERIAL_SETTINGS = {'baudrate': 1152000, 
                        'bytesize': 8, 
                        'parity': 'N', 
                        'stopbits': 1, 
                        'xonxoff': False, 
                        'dsrdtr': False, 
                        'rtscts': False, 
                        'timeout': 3, 
                        'write_timeout': None, 
                        'inter_byte_timeout': None}
    
    _CHUNK_SIZE = 128   # NXTs run into buffer problems if you try to send more than 128 bytes at a time

    _FILE_DELAY = 5 # Number of seconds to pause between file uploads

    def __init__(self, port: str=None, gui: bool = False, loglevel: int = logging.ERROR) -> None:
        self.log = logging.getLogger('Saber_Controller')
        self.log.setLevel(loglevel)
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
            ports = Saber_Controller.get_anima_ports()
            if ports:
                self.port = ports[0]
            else:  # No Anima was found
                raise NoAnimaSaberException
        
        # Initialize the serial connection
        self._ser = serial.Serial(self.port)
        self._ser.apply_settings(self._SERIAL_SETTINGS)
        #self._ser.rs485_mode = serial.rs485.RS485Settings()

    def __del__(self):
        try: # Exception handling is necessary for the case that no saber was found during initialization. In this case, self._ser never gets created
            self._ser.close()
        except Exception:
            pass

    @staticmethod
    def get_ports() -> list[str]:
        '''DEPRECATED: Use get_anima_ports() instead.
        Returns available serial ports as list of strings.'''
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
    def get_anima_ports() -> list[str]:
        '''Returns a list of found ports with an Anima connected.
        If no Anima is found, it will return an empty list.'''
        anima_ports = []
        # Search by VID and PID.
        # EVO: VID=16C0 PID=0483
        # NXT: VID=0483 PID=5740
        ports = lp.grep(r"VID:PID=(16C0|0483):(0483|5740)")
        for port in ports:
            anima_ports.append(port.device)
        return anima_ports

    @staticmethod
    def port_is_anima(port: str) -> bool:
        '''Checks to see if the device attached to port is a Polaris Anima EVO.
        Returns True if an Anima is found, False otherwise.
        
        NB: This method will not throw exceptions. Any exceptions will cause result of False.'''

        try:
            log = logging.getLogger('Saber_Controller')

            log.debug(f'Checking if decvice on port {port} is a Polaris Anima EVO.')
            # Old Checking Logic
            # ------------------
            # ser = serial.Serial(port)
            # ser.apply_settings(Saber_Controller._SERIAL_SETTINGS)
            
            # # Checking logic (based on Nuntis' script):
            # # Send 'V?'. Return false if no response or respond with a 1.x version. Otherwise continue.
            # # Send 'S?'. Return false if no response or response doesn't start with 'S='. Otherwise continue.
            # # Send 'WR?'. Return false if empty or response doesn't sart with 'OK, Write'.
            # # Otherwise return true.
            # log.debug('Sending command: V?')
            # ser.write(b'V?\n')
            # response = ser.readline()
            # log.debug(f'Received response: {response}')
            # if not response or response.startswith(b'V=1.'):
            #     ser.close()
            #     log.info(f'No Polaris Anima EVO found on port {port}')
            #     return False
            
            # log.debug('Sending command: S?')
            # ser.write(b'S?\n')
            # response = ser.readline()
            # log.debug(f'Received response: {response}')
            # if not response or not response.startswith(b'S='):
            #     ser.close()
            #     log.debug(f'No Polaris Anima EVO found on port {port}')
            #     return False
            
            # log.debug('Sending command: WR?')
            # ser.write(b'WR?\n')
            # response = ser.readline()
            # log.debug(f'Received response: {response}')
            # if not response or not response.startswith(b'OK, Write'):
            #     ser.close()
            #     log.debug(f'No Polaris Anima EVO found on port {port}')
            #     return False

            # ser.close()
            # log.info(f'Found Polaris Anima EVO on port {port}')
            # return True

            # New Checking Logic - check VID and PID of device
            # ------------------------------------------------
            ports = lp.grep(port)
            p = next(ports)  # Will raise a StopIteration exception if no port found
            if (p.vid == 0x16C0 and p.pid == 0x0483) or (p.vid == 0x483 and p.pid == 0x5740):
                return True
            return False
        except:
            log.debug(f'No Polaris Anima EVO found on port {port}')
            return False

    def send_command(self, cmd: bytes) -> None:
        '''Send command string to attached saber. It will automatically add b'\n' terminator if not already present.
        
        Note: this method does not check that the saber is write-ready.'''
        if not cmd.endswith(b'\n'):
            cmd += b'\n'
        if len(cmd) <= self._CHUNK_SIZE: # Send the whole thing at once
            self.log.debug(f'Sending command to saber: {cmd}')
            self._ser.write(cmd)
        else: # send in chunks
            self.log.debug(f'Sending command in chunks of {self._CHUNK_SIZE} bytes')
            pos = 0
            while pos < len(cmd):
                self.log.debug(f'Sending command to saber: {cmd[pos:pos+self._CHUNK_SIZE]}')
                self._ser.write(cmd[pos:pos+self._CHUNK_SIZE])
                pos += self._CHUNK_SIZE
                time.sleep(0.5)
    
    def read_line(self) -> bytes:
        '''Reads the next line (terminated by b'\n') from the serial buffer.
        
        Note: this removes the line from the buffer.'''
        response = self._ser.readline()
        self.log.debug(f'Received response: {response}')
        # Sometimes it seems the read request comes too fast, before the anima has had a chance to respond.
        # So now, if I get a blank response, I wait and try again until a response is received or timeout.
        if not response:
            tries = 1
            max_tries = 2
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
                raise InvalidSaberResponseException(f'Command: {cmd}\nResponse: {response}')
            else:
                info['version'] = response.decode().strip()[2:]

            # get serial number
            cmd = b'S?'
            self.send_command(cmd)
            response = self.read_line()
            if not response or not response.startswith(b'S='):
                raise InvalidSaberResponseException(f'Command: {cmd}\nResponse: {response}')
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
            if self.gui: progress_callback.emit(100)
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
    
    def get_used_space(self) -> int:
        '''Returns the amount of used space in Anima storage in bytes.'''
        self.log.info('Getting used space on Anima.')
        cmd = b'USED?'
        self.send_command(cmd)
        response = self.read_line()
        r = response.decode().strip()
        used_space = int(r[5:])
        self.log.info(f'Used space: {used_space} bytes')
        return used_space
    
    def get_total_space(self) -> int:
        '''Returns the amount of total storage space in Anima storage in bytes.'''
        self.log.info('Getting total storage space on Anima.')
        cmd = b'SIZE?'
        self.send_command(cmd)
        response = self.read_line()
        r = response.decode().strip()
        total_space = int(r[5:])
        self.log.info(f'Total space: {total_space} bytes')
        return total_space

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

    def anima_is_NXT(self) -> bool:
        '''Returns True if the attached saber is an NXT, False if not.'''
        info = self.get_saber_info()
        if info['version'][:4] == 'NXT_':
            return True
        return False

    def write_files_to_saber(self, files: list[str], progress_callback: callable = None, add_beep: bool = True) -> None:
        '''Write file(s) to saber. Expects a list of file names.
        
        If add_beep is True (default), this method will automatically add the defauly BEEP.RAW for NXT sabers if no other BEEP.RAW is supplied or already on saber.

        NB: This method does no checking that files exist either on disk or saber. Please verify files before calling this method.'''
        files.sort()

        if self.anima_is_NXT():
            beep_files = [file for file in files if "BEEP.RAW" in file]
            # if a BEEP.RAW is specified, move it to the end of the list.
            # NXTs seem to do better if BEEP.RAW is the last file uploaded
            if beep_files:
                for file in beep_files:
                    self.log.debug(f'Moving file {file} to end of upload list.')
                    files.remove(file)
                    files.append(file)
            elif add_beep:
                current_files = self.list_files_on_saber()
                if 'BEEP.RAW' not in current_files.keys():
                    self.log.info('NXT saber detected and no BEEP.RAW provided. Adding default BEEP.RAW.')
                    files.append(os.path.join(basedir, 'OpenCore_OEM', 'BEEP.RAW'))

        self.log.info(f'Preparing to write file(s) to saber: {files}')
        for file in files:
            self.log.info(f'Writing file to saber: {file}')
            
            # Check for enough free space
            file_size = os.path.getsize(file)
            self.log.debug(f'File size: {file_size}')
            free_space = self.get_free_space()
            if free_space < file_size:
                raise NotEnoughFreeSpaceException(f'File: {file}')

            # Write the file
            with open(file, mode='rb') as binary_file:
                if self.saber_is_ready():
                    bytes_sent = 0
                    fname = os.path.basename(file)
                    report_every_n_bytes = self._CHUNK_SIZE*3 # How often to update the bytes sent display

                    cmd = b'WR=' + fname.encode('utf-8') + b', ' + str(file_size).encode('utf-8') + b'\n'
                    self.send_command(cmd)
                    response = self.read_line()
                    self.log.debug(f'Beginning byte stream for file {fname}')
                    bytes = binary_file.read(1)
                    print(f'{fname} - Data sent: {getHumanReadableSize(bytes_sent)} - Data remaining: {getHumanReadableSize(file_size - bytes_sent)} - Speed: 0.00B/s', end='', flush=True)
                    time.sleep(0.000087)
                    start_time = time.time()
                    while bytes:
                        self._ser.write(bytes)
                        self._ser.flush()
                        if platform.system() != 'Windows': time.sleep(0.000087) # serial drivers write too fast on mac and linux, have to manually force wait
                        bytes_sent += 1#len(bytes)
                        bytes = binary_file.read(1)
                        if bytes_sent % report_every_n_bytes == 0:
                            print(f'\r{fname} - Data sent: {getHumanReadableSize(bytes_sent)} - Data remaining: {getHumanReadableSize(file_size - bytes_sent)} - Speed: {getHumanReadableSize(bytes_sent/(time.time()-start_time))}/s   ', end='', flush=True)
                            if self.gui:
                                progress_callback.emit(bytes_sent)
                    print(f'\r{fname} - Data sent: {getHumanReadableSize(bytes_sent)} - Data remaining: {getHumanReadableSize(file_size - bytes_sent)} - Speed: {getHumanReadableSize(bytes_sent/(time.time()-start_time))}/s           ')
                    if self.gui: progress_callback.emit(bytes_sent)
                else:
                    raise AnimaNotReadyException
            
            response = self.read_line()
            if not response == b'OK, Write Complete\n':
                raise AnimaFileWriteException(f'Error message: {response.strip().decode()}')
        
            self.log.info(f'Successfully wrote file to saber: {file}')
            self.log.info(f'Pausing {self._FILE_DELAY} seconds between file uploads.')
            print(f'Pausing {self._FILE_DELAY} seconds between file uploads', end='', flush=True)
            i = self._FILE_DELAY
            while i > 0:
                if i < 1:
                    time.sleep(i)
                    i = 0
                else:
                    time.sleep(1)
                    print(' .', end='', flush=True)
                    i -= 1
            print('\r', end='', flush=True)

    @staticmethod
    def rgbw_to_byte_str(r: int, g: int, b:int, w: int):
        return bytes(str(r), 'ascii') + b',' + bytes(str(g), 'ascii') + b',' + bytes(str(b), 'ascii') + b','+ bytes(str(w), 'ascii')

    def preview_color(self, r: int, g: int, b: int, w: int):
        cmd = b'P=' + self.rgbw_to_byte_str(r, g, b, w) + b'\n'
        self.send_command(cmd)
        response = self.read_line()
        if response[:2] != b'OK':
            raise InvalidSaberResponseException(f'Command: {cmd}\nResponse: {response}')
    
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
            raise InvalidSaberResponseException(f'Command: {cmd}\nResponse: {response}')
        self.save_config()
    
    def set_active_bank(self, bank:int):
        '''Sets the active bank (0-7)'''
        cmd = b'B=' + bytes(str(bank), 'ascii') + b'\n'
        self.send_command(cmd)
        response = self.read_line()
        if response[:2] != b'OK':
            raise InvalidSaberResponseException(f'Command: {cmd}\nResponse: {response}')
        self.save_config()
    
    @staticmethod
    def _get_cmd_for_sound_effect(effect: str) -> bytes:
        '''returns the command header for the given sound effect. Calling function needs to either add '?' or '=' to the end.'''
        match effect:
            case 'on':
                return b'sON'
            case 'off':
                return b'sOFF'
            case 'hum':
                return b'sHUM'
            case 'swing':
                return b'sSW'
            case 'clash':
                return b'sCL'
            case 'smoothSwingA':
                return b'sSMA'
            case 'smoothSwingB':
                return b'sSMB'
            case _:
                raise InvalidSoundEffectSpecifiedException(f'Specified effect: {effect}')

    def set_sounds_for_effect(self, effect: str, files: list[str]):
        '''Sets the sound list for a given effect.'''
        if not self.gui: print(f'Setting sound files for effect "{effect}".')
        self.log.info(f'Setting sound files for effect "{effect}".')
        self.log.debug(f'{effect}: {str(files)}.')
        cmd = self._get_cmd_for_sound_effect(effect) + b'=' + ','.join(files).encode('utf-8')
        self.send_command(cmd)
        response = self.read_line()
        if response != b'OK ' + cmd + b'\n':
            raise InvalidSaberResponseException(f'Command: {cmd}\nResponse: {response}')
        self.save_config()

    def get_sounds_for_effect(self, effect: str) -> list[str]:
        '''Returns a list of filenames Anima is using for the specified effect.'''
        self.log.debug(f'Retrieving list of sound files for effect "{effect}".')
        cmd = self._get_cmd_for_sound_effect(effect) + b'?'
        self.send_command(cmd)
        response = self.read_line()
        r = response.split(b'=')
        if r[0] != self._get_cmd_for_sound_effect(effect):
            raise InvalidSaberResponseException(f'Command: {cmd}\nResponse: {response}')
        return r[1].decode().strip().split(',')

    def save_config(self):
        '''Save the current configuration to the saber.'''
        self.log.debug(f'Saving configuration on saber.')
        cmd = b'SAVE'
        self.send_command(cmd)
        response = self.read_line()
        if response != b'OK SAVE\n':
            raise InvalidSaberResponseException(f'Command: {cmd}\nResponse: {response}')
    
    def auto_assign_sound_effects(self):
        '''Attempt to automatically set sound effects based on the files currently on the saber.'''
        effects = { # dictionary of effects and matching filename patterns
            'on': 'POWERON',
            'off': 'POWEROFF',
            'hum': 'HUM',
            'swing': 'SWING',
            'clash': 'CLASH',
            'smoothSwingA': 'SMOOTHSWINGH',
            'smoothSwingB': 'SMOOTHSWINGL'
        }
        files = self.list_files_on_saber().keys() # Just the keys because we only need the filenames

        if not self.gui: print('Automatically assigning effects based on the default naming scheme.')
        self.log.info('Automatically assigning effects based on the default naming scheme.')

        # NXTs have problems if you try to set both regular Swing and SmoothSwing effects at the same time.
        # EVOs don't seem to have this problem, but it's probably good practice not to set it anyway.
        # So we search the file list, and if there are any SmoothSwing files, we remove the regular Swing.
        if any("SMOOTHSWING" in file for file in files):
            self.log.info('Detected SmoothSwing files. Ignoring any standard Swing files.')
            files = [f for f in files if not f.startswith('SWING')]

        for effect in effects.keys():
            list = [f for f in files if f.startswith(effects[effect])]
            self.set_sounds_for_effect(effect, list)
            time.sleep(1)

# ---------------------------------------------------------------------- #
# Command Line Operations                                                #
# ---------------------------------------------------------------------- #

def error_handler(e: DocDefaultException):
    '''Print an exception to the log.'''
    log = logging.getLogger()
    log.error(e)
    log.debug(e, exc_info=True)

def main_func():
    log = logging.getLogger()
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))
    log.addHandler(handler)

    exit_code = 0

    # configure command line parser and options
    parser = argparse.ArgumentParser(prog='py2saber',
                                     description='A utility for working with OpenCore-based sabers, based on "sendtosaber" by Nuntis')
    parser.add_argument('-v', '--version',
                        action="version",
                        help='Display version and author information, then exit',
                        version='%(prog)s v{ver} - Author(s): {auth} - {page}'.format(ver=script_version, auth=script_authors, page=script_repo))
    parser.add_argument('-i', '--info', 
                        action="store_true",
                        help='Read and display saber firmware version and serial number')
    parser.add_argument('-l', '--list',
                        action="store_true",
                        help='List all files on saber')
    parser.add_argument('files', nargs="*",
                        help='One or more files to upload to saber (separated by spaces)')
    exit_behavior = parser.add_mutually_exclusive_group()
    exit_behavior.add_argument('-s', '--silent', 
                               action="store_true",
                               help='Exit without waiting for keypress (default)')
    exit_behavior.add_argument('-w', '--wait', 
                               action="store_true",
                               help='Wait for keypress before exiting')
    parser.add_argument('-c', '--continue-on-file-not-found',
                        action="store_true",
                        help='If one or more specified files do not exist, continue processing the remaining files (otherwise program will exit)')
    parser.add_argument('-D', '--debug',
                        action="store_true",
                        help='Show debugging information')
    parser.add_argument('--erase-all', 
                        action="store_true",
                        help='Erase all files on saber')
    parser.add_argument('-y', '--yes',
                        action="store_true",
                        help='Automatically answer "yes" when prompted for Y/N')
    parser.add_argument('--config',
                        action="store_true",
                        help='Display config.ini from saber')
    parser.add_argument('-N', '--no-beep',
                        action="store_true",
                        help="Do not automatically add BEEP.RAW to NXT sabers")
    auto_set_effects = parser.add_mutually_exclusive_group()
    auto_set_effects.add_argument('-e', '--set-effects',
                                  action='store_true',
                                  help='Automatically assign sound files to effects based on the default file naming scheme. (default) (Can be used alone to set effects for the files currently on the saber)')
    auto_set_effects.add_argument('-n', '--no-set-effects',
                                  action='store_true',
                                  help='Do not attempt to automatically assign sound files to effects after uploading')
    parser.add_argument('-t', '--command',
                        action='store', dest='cmd',
                        help='Send literal command CMD to saber (NOTE: only use if you know what you are doing!)')
    
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

        if args.cmd:
            print(f'\nSending command to saber: {args.cmd}')
            sc.send_command(args.cmd.encode('utf-8'))
            response = sc.read_line().strip().decode('utf-8')
            print("Received response: ")
            while response:
                print(response)
                response = sc.read_line().strip().decode('utf-8')
            return
        
        if args.set_effects and not args.files:
            sc.auto_assign_sound_effects()

        if args.erase_all: # erase all files on saber
            print('\n*** This will erase ALL files on the saber! ***')
            if not args.yes:
                yorn = input('Do you want to continue? (Y/N): ')
            else:
                yorn = 'y'
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
                log.debug('Windows system detected. Expanding any wildcards in file names.')
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
                sc.write_files_to_saber(verified_files, add_beep=not args.no_beep)
                print(f'\nSuccessfully wrote file{"s" if len(verified_files)>1 else ""} to saber: {verified_files}')
                if not args.no_set_effects:
                    sc.auto_assign_sound_effects()
            except AnimaFileWriteException as e:
                error_handler(e)
                exit_code = 1

    except Exception as e:
        error_handler(e)
        exit_code = 1
    
    finally:
        if args.wait:
            pause_exit(exit_code, '\nPress any key to exit.')
        else:
            sys.exit(exit_code)

if __name__ == '__main__':
    main_func()