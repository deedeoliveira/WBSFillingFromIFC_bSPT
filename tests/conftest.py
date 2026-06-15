import sys
from unittest import mock

sys.modules.setdefault("tkinter", mock.MagicMock())
sys.modules.setdefault("tkinter.ttk", mock.MagicMock())
sys.modules.setdefault("tkinter.filedialog", mock.MagicMock())
sys.modules.setdefault("tkinter.messagebox", mock.MagicMock())
sys.modules.setdefault("tkinter.scrolledtext", mock.MagicMock())
