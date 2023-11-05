import openai
import keyboard
import pyperclip
import configparser
import tkinter as tk

def load_config(config_file: str) -> configparser.ConfigParser:
  """Load the configuration file."""
  config = configparser.ConfigParser()
  config.read(config_file)
  return config

def display_window(answer: str) -> None:
  """Display a popup window with the given answer."""
  config = load_config('config.cfg')

  root = tk.Tk()
  root.withdraw()
  
  popup = tk.Toplevel(root)
  popup.title("Update")
  popup.overrideredirect(True)

  tk.Label(popup, text=answer, padx=10, pady=10).pack()

  try:
    popup.attributes("-alpha", float(config["Settings"]["PopOpacity"]))
    popup.geometry(config["Settings"]["PopPosition"])
    popup.attributes("-topmost", True)

    popup.after(int(config["Settings"]["PopTimer"]), lambda: (popup.destroy(), root.quit()))
  except ValueError as ve:
    print(f"Error with configuration values: {ve}")

  root.mainloop()

def generate_response(message: str) -> str:
  """Generate a response for the given message using the OpenAI API."""
  config = load_config('config.cfg')

  messages = [
    {"role": "system", "content": config["Settings"]["PromptSystem"]},
    {"role": "user", "content": f'{config["Settings"]["PromptUser"]} {message}'}
  ]
  
  try:
    chat = openai.ChatCompletion.create(model=config["Settings"]["Model"], messages=messages)
    return chat.choices[0].message.content
  except Exception as e:
    return f"Error generating response: {e}"

def on_key_event(e: keyboard.KeyboardEvent) -> None:
  """Handle the key event."""
  config = load_config('config.cfg')
  
  if e.event_type == keyboard.KEY_DOWN:
    copied_text = pyperclip.paste()
    response = generate_response(copied_text)
    if e.name == config["Settings"]["KeyWrite"]:
      keyboard.write(response)
    elif e.name == config["Settings"]["KeyPop"]:
      display_window(response)

if __name__ == "__main__":
  config = load_config('config.cfg')
  
  openai.api_key = config["Api"]["OpenAiKey"]
  keyboard.hook(on_key_event)
  keyboard.wait()
