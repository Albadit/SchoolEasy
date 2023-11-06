import winreg
import urllib.request
from pathlib import Path
import win32com.client
from xml.etree.ElementTree import Element, SubElement, tostring
from xml.dom import minidom

def create_config(config_file_path: Path, settings: dict) -> None:
  contents = '\n'.join(f'[{section}]\n' + '\n'.join(f'{k} = {v}' for k, v in options.items()) + '\n' for section, options in settings.items())
  config_file_path.write_text(contents)

def generate_xml_content(setup_path: Path, config_file_path: Path, file_name: str) -> None:
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
    'MultipleInstancesPolicy': 'Queue', 'DisallowStartIfOnBatteries': 'false',
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
  
  config_file_path.write_text(xml_str_finish, encoding='utf-16')

def setup_config_folder() -> None:
  key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
  try:
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
      paths = {name: winreg.QueryValueEx(key, name)[0] for name in ("Personal", "Desktop")}
  except (FileNotFoundError, OSError) as e:
    print(f"Error accessing registry: {e}")
    return
  
  setup_path = Path(paths["Personal"], 'SchoolEasy')

  settings = {
    'Api': {
      'OpenAiKey': ''
    },
    'Settings': {
      'PopTimer': '3000',
      'PopPosition': '+300+200',
      'PopOpacity': '0.5',
      'Model': 'gpt-4',
      'PromptSystem': 'You are a helpful assistant specialized in answering multiple-choice questions.',
      'PromptUser': 'Give me only the correct answer and nothing else:',
      'KeyWrite': '`',
      'KeyPop': '~',
    }
  }

  setup_path.mkdir(parents=True, exist_ok=True)
  # (setup_path / 'main').mkdir(exist_ok=True)

  create_config(setup_path / "config.cfg", settings)
  generate_xml_content(setup_path, setup_path / "schooleasy.xml", "schooleasy.xml")
  
  # Download practice config
  # releases/latest
  url = 'https://raw.githubusercontent.com/Albadit/SchoolEasy/main/app/SchoolEasy.exe'
  practice_cfg_path = setup_path / Path(url).name
  urllib.request.urlretrieve(url, practice_cfg_path)

  # Create desktop shortcut
  shortcut = win32com.client.Dispatch("WScript.Shell").CreateShortcut(str(Path(paths["Desktop"], "School Easy.lnk")))
  shortcut.Targetpath = str(practice_cfg_path)
  shortcut.WorkingDirectory = str(setup_path)
  shortcut.IconLocation = str(practice_cfg_path)
  shortcut.save()

if __name__ == "__main__":
  setup_config_folder()
