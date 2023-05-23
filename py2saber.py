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
from win_wildcard import expand_windows_wildcard

script_version = '0.9b'
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

# Serial port communication settings
serial_settings = {'baudrate': 115200, 
                   'bytesize': 8, 
                   'parity': 'N', 
                   'stopbits': 1, 
                   'xonxoff': False, 
                   'dsrdtr': False, 
                   'rtscts': False, 
                   'timeout': 3, 
                   'write_timeout': None, 
                   'inter_byte_timeout': None}

log = logging.Logger('py2saber')
#log.setLevel(logging.DEBUG)
#stream = logging.StreamHandler(sys.stdout)
#stream.setLevel(logging.DEBUG)
#stream.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s: %(message)s'))
#log.addHandler(stream)

def get_ports() -> list[str]:
    '''Returns available serial ports as list of strings.'''
    serial_ports = []
    match_string = r''

    log.info('Searching for available serial ports.')
    port_list = lp.comports()
    log.debug(f'Found {len(port_list)} ports before filtering: {port_list}')

    # Filter down list of ports depending on OS
    system = platform.system()
    logging.info(f'Detected OS: {system}')
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
    
    log.info(f'Found {len(serial_ports)} port(s).')
    log.debug(f'Found ports: {serial_ports}')
    return serial_ports

def port_is_anima(port: str) -> bool:
    '''Checks to see if the device attached to port is a Polaris Anima EVO.
    Returns True if an Anima is found, False otherwise.
    
    NB: This method will not throw serial exceptions. Any exceptions will cause result of False.'''

    try:
        log.info(f'Checking if decvice on port {port} is a Polaris Anima EVO.')
        ser = serial.Serial(port)
        ser.apply_settings(serial_settings)
        
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

def get_saber_info(port: str) -> dict:
    '''Retrieve firmware version and serial number from saber. Returns a dict with keys 'version' and 'serial'.'''
    if port_is_anima(port):
        log.info('Retrieving firmware version and serial number from saber.')
        info = {}
        ser = serial.Serial(port)
        ser.apply_settings(serial_settings)

        # get firmware version
        log.debug('Sending command: V?')
        ser.write(b'V?\n')
        response = ser.readline()
        log.debug(f'Received response: {response}')
        if not response or not response.startswith(b'V='):
            ser.close()
            log.error('Anima not ready.')
            raise AnimaNotReadyException
        else:
            info['version'] = response.decode().strip()[2:]

        # get serial number
        log.debug('Sending command: S?')
        ser.write(b'S?\n')
        response = ser.readline()
        log.debug(f'Received response: {response}')
        if not response or not response.startswith(b'S='):
            ser.close()
            log.error('Anima not ready.')
            raise AnimaNotReadyException
        else:
            info['serial'] = response.decode().strip()[2:]
        
        log.info(f'Found saber info: {info}')
        return info
    else:
        log.error('No Anima saber found.')
        raise NoAnimaSaberException

def list_files_on_saber_as_bytes(port: str) -> bytes:
    '''Returns the raw byte string reported by the saber LIST? command.'''
    file_list = b''
    log.info(f'Retrieving file list from saber on port {port}')

    if port_is_anima(port):
        ser = serial.Serial(port)
        ser.apply_settings(serial_settings)
        log.debug('Sending command: LIST?')
        ser.write(b'LIST?\n')
        response = ser.readline()
        log.debug(f'Received response: {response}')
        while response:
            file_list += response
            response = ser.readline()
            log.debug(f'Received response: {response}')
        ser.close()
    
        log.debug(f'Final byte string: {file_list}')
        return file_list
    else:
        log.error('No Anima saber found.')
        raise NoAnimaSaberException

def erase_all_files(port: str) -> None:
    '''Erases all files on the anima. USE CAREFULLY.'''
    if port_is_anima(port):
        ser = serial.Serial(port)
        ser.apply_settings(serial_settings)
        log.info('Checking that saber is ready.')
        log.debug('Sending command: WR?')
        ser.write(b'WR?\n') # check that saber is ready
        response = ser.readline()
        log.debug(f'Received response: {response}')
        if response == b'OK, Write Ready\n':
            log.info('Erasing all files on saber. This may take several minutes.')
            log.debug('Sending command: ERASE=ALL')
            ser.write(b'ERASE=ALL\n')
            
            # Listen to output stream to make sure process is ongoing/completed
            response = ser.readline() # "Erasing Serial Flash, this may take 20s to 2 minutes\n"
            log.debug(f'Received response: {response}')
            #print(response.decode().strip())
            ser.timeout = 5 # increase timeout while we wait for #s to display
            c = ser.read(1)
            while c < b'\x41': # any ascii character before 'A'
                print(c.decode(), end='', flush=True)
                c = ser.read(1)
            response = c + ser.readline()
            log.debug(f'Received response: {response}')
            print(response.decode().strip())
            response = ser.readline()
            log.debug(f'Received response: {response}')
            print(response.decode().strip())
            log.info(response.decode().strip())

            ser.close()
        else:
            ser.close()
            log.error('Anima not ready.')
            raise AnimaNotReadyException
    else:
        log.error('No Anima saber found.')
        raise NoAnimaSaberException

def get_free_space(port: str) -> int:
    '''Returns the amount of free space in Anima storage in bytes.'''
    ser = serial.Serial(port)
    ser.apply_settings(serial_settings)
    log.info('Getting free space on Anima.')
    log.debug('Sending command: FREE?')
    ser.write(b'FREE?\n')
    response = ser.readline()
    log.debug(f'Received response: {response}')
    r = response.decode().strip()
    free_space = int(r[5:])
    log.info(f'Free space: {free_space} bytes')
    ser.close()
    return free_space

def write_files_to_saber(port: str, files: list[str]) -> None:
    '''Write file(s) to saber. First argument is the port to write to; second argument is a list of filenames.

    NB: This method does no checking that files exist either on disk or saber. Please verify files before calling this method.'''
    if port_is_anima(port):
        ser = serial.Serial(port)
        ser.apply_settings(serial_settings)
        log.info('Checking that saber is ready.')
        log.debug('Sending command: WR?')
        ser.write(b'WR?\n') # check that saber is ready
        response = ser.readline()
        log.debug(f'Received response: {response}')
        if response == b'OK, Write Ready\n':
            log.info(f'List of files to write to saber: {files}')
            for file in files:
                log.info(f'Writing file to saber: {file}')
                
                # Check for enough free space
                file_size = os.path.getsize(file)
                log.debug(f'File size: {file_size}')
                ser.close() # have to close the connection so other method can access it
                free_space = get_free_space(port)
                if free_space < file_size:
                    log.error(f'Not enough free space on saber for file {file}')
                    ser.close()
                    raise NotEnoughFreeSpaceException
                ser.open()

                # Write the file
                with open(file, mode='rb') as binary_file:
                    bytes_sent = 0
                    fname = os.path.basename(file)
                    report_every_n_bytes = 512 # How often to update the bytes sent display

                    cmd = b'WR=' + fname.encode('utf-8') + b',' + str(file_size).encode('utf-8') + b'\n'
                    log.debug(f'Sending command: {cmd}')
                    ser.write(cmd)
                    log.debug(f'Beginning byte stream for file {fname}')
                    byte = binary_file.read(1)
                    print(f'{fname} - Bytes sent: {bytes_sent} - Bytes remaining: {file_size - bytes_sent}', end='', flush=True)
                    while byte:
                        ser.write(byte)
                        bytes_sent += 1
                        byte = binary_file.read(1)
                        if bytes_sent % report_every_n_bytes == 0:
                            print(f'\r{fname} - Bytes sent: {bytes_sent} - Bytes remaining: {file_size - bytes_sent}', end='', flush=True)
                    print(f'\r{fname} - Bytes sent: {bytes_sent} - Bytes remaining: {file_size - bytes_sent}     ')
                
                response = ser.readline()
                log.debug(f'Received response: {response}')
                if not response == b'OK, Write Complete\n':
                    ser.close()
                    log.error(f'Error writing file to Anima. Error message: {response}')
                    raise AnimaFileWriteException
            
                log.info(f'Successfully wrote file to Anima: {file}')
            ser.close()
        else:
            ser.close()
            log.error('Anima not ready.')
            raise AnimaNotReadyException
    else:
        log.error('No Anima saber found.')
        raise NoAnimaSaberException
    
# ---------------------------------------------------------------------- #
# Command Line Operations                                                #
# ---------------------------------------------------------------------- #

def main_func():
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
    
    args = parser.parse_args(args=None if sys.argv[1:] else ['--help'])

    if args.debug:
        log.setLevel(logging.DEBUG)

    # Find port that Anima is connected to
    try:
        print('Searching for OpenCore saber.')
        port_list = get_ports()
        port = ''
        # look for first port with an Anima connected
        for p in port_list:
            if port_is_anima(p):
                port = p
                print('OpenCore saber found.')
                break
        if not port:
            raise NoAnimaSaberException
        
        # Execute functions based on arguments provided
        if args.info: # display saber information
            print('\nRetrieving saber information')
            inf = get_saber_info(port)
            print(f'Firmware version:\tv{inf["version"]}\nSerial Number:\t\t{inf["serial"]}')
        
        if args.list: # list files on saber
            print('\nRetrieving list of files on saber.')
            file_list = list_files_on_saber_as_bytes(port)
            print(file_list.decode().strip())
        
        if args.erase_all: # erase all files on saber
            print('\n*** This will erase ALL files on the saber! ***')
            yorn = input('Do you want to continue? (Y/N): ')
            if yorn.lower() == 'y' or yorn.lower() == 'yes':
                # do the thing
                print('Erasing all files on saber. This may take several minutes.')
                erase_all_files(port)
            else:
                print('Aborting saber erase command.')
                sys.exit(1)

        if args.files: # write file(s) to saber
            # for Windows, we need to manually expand any wildcards in the input list
            if platform.system() == 'Windows':
                log.info('Windows system detected. Expanding any wildcards in file names.')
                expanded_files = []
                for file in args.files:
                    expanded_files.extend(expand_windows_wildcard(file, only_files=True))
                args.files = expanded_files              

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
                    log.error(e)
                    if args.continue_on_file_not_found:
                        continue
                    else:
                        exit_code = 1
                        sys.exit(1)
            
            # send files to saber
            write_files_to_saber(port, verified_files)

    except NoAnimaSaberException as e:
        log.error('No OpenCore saber found.')
        exit_code = 1

    except Exception as e:
        log.error(f'An error has occurred: {e}')
        exit_code = 1
    
    finally:
        if args.wait:
            pause_exit(exit_code, '\nPress any key to exit.')
        else:
            sys.exit(exit_code)

if __name__ == '__main__':
    main_func()