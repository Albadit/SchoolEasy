import winreg
import urllib.request
from pathlib import Path
import win32com.client
import ctypes
import os
import sys

def is_admin():
  """Check if the script is running with administrative privileges."""
  try:
    return ctypes.windll.shell32.IsUserAnAdmin()
  except:
    return False

def create_config(config_path: Path, file_name: str) -> None:
  settings = {
    'Api': {
      'OpenAiKey': '',
    },
    'Settings': {
      'PopTimer': '3000',
      'PopPosition': '+300+200',
      'PopOpacity': '0.5',
      'Model': 'gpt-4',
      'PromptSystem': 'You are a helpful assistant specialized in answering multiple-choice questions.',
      'PromptUser': 'Give me only the correct answer and nothing else:',
    },
    'Keys': {
      'KeyPop': '~',
      'KeyRepop': '`',
    }
  }

  contents = '\n'.join(f'[{section}]\n' + '\n'.join(f'{k} = {v}' for k, v in options.items()) + '\n' for section, options in settings.items())
  Path(config_path / file_name).write_text(contents)
  print(f'{file_name} has been Generate.')

def task_scheduler(setup_path: Path) -> None:
  # Connect to the Task Scheduler
  scheduler = win32com.client.Dispatch('Schedule.Service')
  scheduler.Connect()
  root_folder = scheduler.GetFolder('\\')

  # Create a new task definition
  task_def = scheduler.NewTask(0)

  # Set registration info
  task_def.RegistrationInfo.Description = 'SchoolEasy Task'
  task_def.RegistrationInfo.Author = 'SchoolEasy'

  # Create Triggers
  # Boot Trigger
  task_def.Triggers.Create(8)  # 8 = Boot Trigger

  # Logon Trigger
  task_def.Triggers.Create(9)  # 9 = Logon Trigger

  # Session State Change Trigger for Session Unlock
  session_unlock_trigger = task_def.Triggers.Create(11)  # 11 = Session State Change Trigger
  session_unlock_trigger.StateChange = 8  # 8 = Session unlock

  # Idle Trigger
  task_def.Triggers.Create(6)  # 6 = Idle Trigger

  # Session State Change Trigger for Console Connect
  console_connect_trigger = task_def.Triggers.Create(11)  # Reusing 11 for a different session state change
  console_connect_trigger.StateChange = 4  # 4 = Console connect

  # Set the principal for the task
  task_def.Principal.LogonType = 3  # 3 = Interactive Token
  task_def.Principal.RunLevel = 1   # 1 = Highest Available

  # Set task settings
  settings = task_def.Settings
  settings.MultipleInstances = 2  # 2 = Ignore New
  settings.DisallowStartIfOnBatteries = False
  settings.StopIfGoingOnBatteries = False
  settings.AllowHardTerminate = False
  settings.StartWhenAvailable = True
  settings.RunOnlyIfNetworkAvailable = False
  settings.Enabled = True
  settings.Hidden = False
  settings.RunOnlyIfIdle = False
  settings.WakeToRun = True
  settings.ExecutionTimeLimit = "P1D"  # Period of one day
  settings.Priority = 7

  # Idle settings
  settings.IdleSettings.StopOnIdleEnd = True
  settings.IdleSettings.RestartOnIdle = False

  # Set an action
  action = task_def.Actions.Create(0)  # 0 = Execute
  action.Path = f'{setup_path}\\SchoolEasy.exe'
  action.Arguments = f'{setup_path}\\config.cfg'
  action.WorkingDirectory = f'{setup_path}' 

  # Register the task
  task_name = '\\SchoolEasy'
  root_folder.RegisterTaskDefinition(task_name, task_def, 6, '', '', 3)  # 6 = CreateOrUpdate, 3 = TaskLogonType.InteractiveToken
  print('Task Scheduler has been created.')

def download_exe(setup_path: Path) -> None:
  try:
    url = 'https://raw.githubusercontent.com/Albadit/SchoolEasy/main/app/SchoolEasy.exe'
    cfg_path = Path(setup_path / Path(url).name)
    urllib.request.urlretrieve(url, cfg_path)
    print(f'{Path(url).name} has been downloaded.')
  except Exception as e: 
    print(e)

def create_shortcut(setup_path: Path, cfg_path: Path, shortcut_link: Path) -> None:
  shortcut = win32com.client.Dispatch("WScript.Shell").CreateShortcut(str(shortcut_link))
  shortcut.Targetpath = str(cfg_path)
  shortcut.WorkingDirectory = str(setup_path)
  shortcut.IconLocation = str(cfg_path)
  shortcut.save()
  print('Shorcut has been created.')

def parse_selection(install_list: list) -> list:
  selected_indices = []

  while True:
    os.system('cls' if os.name == 'nt' else 'clear')
    print("Select the components to install/reinstall:")
    for i, item in enumerate(install_list):
      print(f"{i + 1}. {item}")
    print(f"{len(install_list) + 1}. All")
    print("Selected:", ", ".join(selected_indices))

    selection = input("Enter your choice (e.g., 1, 3-4, 5 for all, Press enter if you're done): ")

    # Check for 'Enter' to finish selection
    if selection == '':
      break
    elif selection == f'{len(install_list) + 1}':
      return [int(i + 1) for i in range(len(install_list))]

    # Convert selection to integer and validate
    try:
      selection_int = int(selection)
      if selection_int < 1 or selection_int > len(install_list) + 2:
        print("Invalid selection. Please enter a valid number.")
        input(exit_text)
        continue
    except ValueError:
      print("Invalid input. Please enter a number.")
      input(exit_text)
      continue

    # Handle addition or removal of selection
    if selection in selected_indices:
      selected_indices.remove(selection)
    else:
      selected_indices.append(selection)

  # Converting selected_indices to integers and sorting
  return sorted([int(idx) for idx in selected_indices])

def setup() -> None:
  key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
  try:
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
      paths = {name: winreg.QueryValueEx(key, name)[0] for name in ("Personal", "Desktop")}
  except (FileNotFoundError, OSError) as e:
    print(f"Error accessing registry: {e}")
    return
  
  setup_path = Path(paths["Personal"], 'SchoolEasy')

  install_list = ["Download exe", "Generate config", "Scheduler Tasks", "Create shortcut"]

  if setup_path.exists():
    context = "The setup path already exists. Do you want to install/reinstall components (y/n): "
    user_input = input(context).strip().lower()
    while user_input not in ['no', 'n' , 'yes', 'y']:
      os.system('cls')
      user_input = input(context).strip().lower()
      if user_input == ['no', 'n']:
        return
    selected_index = parse_selection(install_list)
  else:
    selected_index = [int(i + 1) for i in range(len(install_list))]

  setup_path.mkdir(parents=True, exist_ok=True)

  if 1 in selected_index:
    download_exe(setup_path)
  if 2 in selected_index:
    create_config(setup_path, "config.cfg")
  if 3 in selected_index:
    task_scheduler(setup_path)
  if 4 in selected_index:
    cfg_path = Path(setup_path / "SchoolEasy.exe")
    if cfg_path.exists():
      create_shortcut(setup_path, cfg_path, Path(paths["Desktop"], "School Easy.lnk"))

  # generate_config(setup_path / "config.cfg")
  # task_scheduler(setup_path)
  # cfg_path = download_exe(setup_path)
  # create_shortcut(setup_path, cfg_path, Path(paths["Desktop"], "School Easy.lnk"))

if __name__ == "__main__":
  if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit(0)

  exit_text = "Press Enter to continue..."
  setup()
  print("Your setup is done. You're ready to go.")
  input(exit_text)
