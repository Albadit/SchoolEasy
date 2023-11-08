import winreg
import urllib.request
from pathlib import Path
import win32com.client
from xml.etree.ElementTree import Element, SubElement, tostring
from xml.dom import minidom
import os

def generate_config(config_path: Path, file_name: str) -> None:
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
      'KeyWrite': '`',
      'KeyPop': '~',
    }
  }

  contents = '\n'.join(f'[{section}]\n' + '\n'.join(f'{k} = {v}' for k, v in options.items()) + '\n' for section, options in settings.items())
  Path(config_path / file_name).write_text(contents)
  print(f'{file_name} has been Generate.')

def generate_xml(setup_path: Path, file_name: str) -> None:
  task = Element('Task', {'version': '1.2', 'xmlns': 'http://schemas.microsoft.com/windows/2004/02/mit/task'})

  # RegistrationInfo
  registration_info = SubElement(task, 'RegistrationInfo')
  SubElement(registration_info, 'Author').text = 'SchoolEasy'
  SubElement(registration_info, 'URI').text = '\\SchoolEasy'

  # Triggers
  triggers = SubElement(task, 'Triggers')
  for i, trigger_type in enumerate(['BootTrigger', 'LogonTrigger', 'SessionStateChangeTrigger', 'IdleTrigger', 'SessionStateChangeTrigger']):
    trigger = SubElement(triggers, trigger_type)
    SubElement(trigger, 'Enabled').text = 'true'
    if trigger_type == 'SessionStateChangeTrigger':
      state_change_text = 'SessionUnlock' if i == 2 else 'ConsoleConnect'
      SubElement(trigger, 'StateChange').text = state_change_text

  # Principals
  principals = SubElement(task, 'Principals')
  principal = SubElement(principals, 'Principal', {'id': 'Author'})
  SubElement(principal, 'LogonType').text = 'InteractiveToken'
  SubElement(principal, 'RunLevel').text = 'HighestAvailable'

  # Settings
  settings = SubElement(task, 'Settings')
  settings_dict = {
    'MultipleInstancesPolicy': 'IgnoreNew', 'DisallowStartIfOnBatteries': 'false',
    'StopIfGoingOnBatteries': 'false', 'AllowHardTerminate': 'false',
    'StartWhenAvailable': 'true', 'RunOnlyIfNetworkAvailable': 'false',
    'AllowStartOnDemand': 'true', 'Enabled': 'true', 'Hidden': 'false',
    'RunOnlyIfIdle': 'false', 'WakeToRun': 'true', 'ExecutionTimeLimit': 'P1D',
    'Priority': '7'
  }
  for key, value in settings_dict.items():
    SubElement(settings, key).text = value

  idle_settings = SubElement(settings, 'IdleSettings')
  SubElement(idle_settings, 'StopOnIdleEnd').text = 'true'
  SubElement(idle_settings, 'RestartOnIdle').text = 'false'

  restart_on_failure = SubElement(settings, 'RestartOnFailure')
  SubElement(restart_on_failure, 'Interval').text = 'PT1M'
  SubElement(restart_on_failure, 'Count').text = '10'

  # Actions
  actions = SubElement(task, 'Actions', {'Context': 'Author'})
  exec_element = SubElement(actions, 'Exec')
  SubElement(exec_element, 'Command').text = f'{setup_path}\\SchoolEasy.exe'
  SubElement(exec_element, 'Arguments').text = f'{setup_path}\\config.cfg'
  SubElement(exec_element, 'WorkingDirectory').text = f'{setup_path}'

  # Pretty print the XML
  xml_str = minidom.parseString(tostring(task)).toprettyxml(indent="  ")

  xml_str_finish = xml_str.replace('<?xml version="1.0" ?>', '<?xml version="1.0" encoding="UTF-16"?>')
  
  Path(setup_path / file_name).write_text(xml_str_finish, encoding='utf-16')
  print(f'{file_name} has been Generate.')

def download_exe(setup_path: Path) -> None:
  try:
    url = 'https://raw.githubusercontent.com/Albadit/SchoolEasy/main/app/SchoolEasy.exe'
    cfg_path = Path(setup_path / Path(url).name)
    urllib.request.urlretrieve(url, cfg_path)
    print(f'{Path(url).name} has been downloaded.')
  except: 
    print(f"Error: You can't install or reinstall the .exe file while it is running. Stop the program before installing.")

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
        input("Press Enter to continue...")
        continue
    except ValueError:
      print("Invalid input. Please enter a number.")
      input("Press Enter to continue...")
      continue

    # Handle addition or removal of selection
    if selection in selected_indices:
      selected_indices.remove(selection)
    else:
      selected_indices.append(selection)

  # Converting selected_indices to integers and sorting
  return sorted([int(idx) for idx in selected_indices])

def setup_config_folder() -> None:
  key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
  try:
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
      paths = {name: winreg.QueryValueEx(key, name)[0] for name in ("Personal", "Desktop")}
  except (FileNotFoundError, OSError) as e:
    print(f"Error accessing registry: {e}")
    return
  
  setup_path = Path(paths["Personal"], 'SchoolEasy')

  install_list = ["Generate config", "Generate xml", "Download exe", "Create shortcut"]

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
    generate_config(setup_path, "config.cfg")
  if 2 in selected_index:
    generate_xml(setup_path, "schooleasy.xml")
  if 3 in selected_index:
    download_exe(setup_path)
  if 4 in selected_index:
    cfg_path = Path(setup_path / "SchoolEasy.exe")
    if cfg_path.exists():
      create_shortcut(setup_path, cfg_path, Path(paths["Desktop"], "School Easy.lnk"))

  # generate_config(setup_path / "config.cfg")
  # generate_xml(setup_path, "schooleasy.xml")
  # cfg_path = download_exe(setup_path)
  # create_shortcut(setup_path, cfg_path, Path(paths["Desktop"], "School Easy.lnk"))

if __name__ == "__main__":
  setup_config_folder()
  print("Your setup is done. You're ready to go.")
  input("Press Enter to continue...")
