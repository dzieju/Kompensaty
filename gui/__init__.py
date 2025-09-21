"""
Moduł GUI aplikacji.
"""
from .main_window import MainWindow
from .config_tab import ConfigTab
from .search_criteria_tab import SearchCriteriaTab
from .templates_tab import TemplatesTab
from .logs_tab import LogsTab
from .about_tab import AboutTab

__all__ = [
    'MainWindow',
    'ConfigTab', 
    'SearchCriteriaTab',
    'TemplatesTab',
    'LogsTab',
    'AboutTab'
]