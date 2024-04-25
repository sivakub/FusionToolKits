
from .backToBody import entry as backToBody

commands = [
    backToBody
]

def start():
    for command in commands:
        command.start()

def stop():
    for command in commands:
        command.stop()