import os, shutil
import win32com.client as win32
from clear_screen import clear

def create_shortcut(file_path, shortcut_path):
    shell = win32.Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = file_path
    shortcut.save()

def hydrogen():
    try:
        clear()
        with open('start.bat', 'w') as batch:
            batch.write('cd Ripcords\n')

        with open('kill.bat', 'w') as kill_batch:
            kill_batch.write('taskkill /im Ripcord.exe')

        amount = int(input('Amount of Ripcords: '))
        done = 0

        for _ in range(amount):
            done +=1
            shutil.rmtree(f'Ripcords\\Ripcord_{done}', ignore_errors=True)
            shutil.copytree('Ripcord', f'Ripcords\\Ripcord_{done}')

            file_path = os.path.join(os.path.abspath(f'Ripcords\\Ripcord_{done}'),'Ripcord.exe')
            shortcut_path = os.path.join(os.path.abspath('Ripcords'), f'Ripcord_{done}.lnk')
            create_shortcut(file_path, shortcut_path)

            
            print(f'Made 1 Ripcord | Copy #{done}')

            with open('start.bat', 'a') as batch:
                batch.write(f'start Ripcord_{done}.lnk\n')
        
        input('\nFinished, press enter to return to main menu.')
        hydrogen()


    except KeyboardInterrupt:
        hydrogen()

if __name__ == '__main__':
    hydrogen()