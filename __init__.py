try:
    from wmi_wrapper import WmiHandler
except ModuleNotFoundError:
    from .wmi_wrapper import WmiHandler