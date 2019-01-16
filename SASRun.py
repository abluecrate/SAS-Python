import subprocess
import argparse
from pathlib import Path

# CONSOLE FORMATTING
from colorama import init; init()
from colorama import Fore, Back, Style
logError = '\n[LOG]' + Fore.RED + Style.BRIGHT + ' [ERROR] ' + Style.RESET_ALL

#--------------------------------------------------------------------------------------------------------
ap = argparse.ArgumentParser()  # Intilialize Argument Parser
ap.add_argument('-p', '--project', help = 'Project Path')
args = vars(ap.parse_args())    # Gather Arguments
#--------------------------------------------------------------------------------------------------------

def runSAS(projectPath):
    if Path(projectPath).is_file():
        print('\n--------------------------------------------------------------------------')
        print('Launching SAS Project From:\n' + projectPath)
        print('--------------------------------------------------------------------------')
        print(Back.BLUE)            
        subprocess.check_call('cscript run.vbs "' + projectPath + '"')    # Call SAS and Run Program
        print(Style.RESET_ALL)
        print('\n[LOG] SAS Project Complete')
    else:
        print(logError + 'PROJECT FILE NOT FOUND')
        
projectPath = args['project']
runSAS(projectPath)