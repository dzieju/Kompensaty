"""
Zakładka logów aplikacji.
"""
import tkinter as tk
from tkinter import ttk, scrolledtext
import logging
from gui.base_tab import BaseTab


class LogsTab(BaseTab):
    """Zakładka z logami aplikacji."""
    
    def __init__(self, parent_frame: ttk.Frame):
        """
        Inicjalizuje zakładkę logów.
        
        Args:
            parent_frame: Ramka rodzica
        """
        self.log_text_area = None
        self.log_handler = None
        super().__init__(parent_frame)
        self._setup_log_handler()

    def setup_ui(self) -> None:
        """Konfiguruje interfejs użytkownika."""
        # Ramka główna dla logów
        main_frame = ttk.LabelFrame(self.parent_frame, text="Logi Aplikacji", padding=10)
        main_frame.pack(expand=True, fill='both', padx=10, pady=5)
        
        # Panel sterowania
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill='x', pady=(0, 5))
        
        ttk.Button(
            control_frame, 
            text="Wyczyść logi", 
            command=self.clear_logs
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            control_frame, 
            text="Zapisz logi", 
            command=self.save_logs
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        # Poziom logowania
        ttk.Label(control_frame, text="Poziom:").pack(side=tk.LEFT, padx=(10, 5))
        
        self.log_level_var = tk.StringVar(value="INFO")
        log_level_combo = ttk.Combobox(
            control_frame,
            textvariable=self.log_level_var,
            values=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
            state="readonly",
            width=10
        )
        log_level_combo.pack(side=tk.LEFT)
        log_level_combo.bind("<<ComboboxSelected>>", self._on_log_level_changed)
        
        # Checkbox automatycznego przewijania
        self.auto_scroll_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            control_frame,
            text="Auto-przewijanie",
            variable=self.auto_scroll_var
        ).pack(side=tk.RIGHT)
        
        # Obszar tekstowy dla logów - BEZ PROBLEMATYCZNYCH OPCJI KOLORÓW
        self.log_text_area = scrolledtext.ScrolledText(
            main_frame,
            wrap=tk.WORD,
            state='disabled',
            height=25,
            font=("Consolas", 9)
        )
        self.log_text_area.pack(expand=True, fill='both')
        
        # Konfiguracja kolorów dla różnych poziomów logów
        self._setup_log_colors()

    def _setup_log_colors(self) -> None:
        """Konfiguruje kolory dla różnych poziomów logów."""
        try:
            self.log_text_area.tag_config("DEBUG", foreground="blue")
            self.log_text_area.tag_config("INFO", foreground="black")
            self.log_text_area.tag_config("WARNING", foreground="orange")
            self.log_text_area.tag_config("ERROR", foreground="red")
            self.log_text_area.tag_config("CRITICAL", foreground="red", background="yellow")
        except Exception as e:
            # Jeśli kolory nie działają, po prostu ignoruj
            print(f"Uwaga: Nie można ustawić kolorów logów: {e}")

    def _setup_log_handler(self) -> None:
        """Konfiguruje handler logowania dla GUI."""
        if self.log_text_area:
            self.log_handler = TextHandler(self.log_text_area, self)
            self.log_handler.setLevel(logging.INFO)
            
            formatter = logging.Formatter(
                '%(asctime)s - %(levelname)-8s - %(name)s - %(message)s',
                datefmt='%H:%M:%S'
            )
            self.log_handler.setFormatter(formatter)
            
            # Dodaj handler do root loggera
            logging.getLogger('').addHandler(self.log_handler)

    def _on_log_level_changed(self, event=None) -> None:
        """Obsługuje zmianę poziomu logowania."""
        if self.log_handler:
            level_name = self.log_level_var.get()
            level = getattr(logging, level_name, logging.INFO)
            self.log_handler.setLevel(level)
            
            # Dodaj informację do logów
            logging.info(f"Zmieniono poziom logowania na: {level_name}")

    def clear_logs(self) -> None:
        """Czyści obszar logów."""
        if self.log_text_area:
            self.log_text_area.configure(state='normal')
            self.log_text_area.delete('1.0', tk.END)
            self.log_text_area.configure(state='disabled')
            logging.info("Wyczyszczono obszar logów GUI")

    def save_logs(self) -> None:
        """Zapisuje logi do pliku."""
        from tkinter import filedialog, messagebox
        
        try:
            filename = filedialog.asksaveasfilename(
                title="Zapisz logi",
                defaultextension=".log",
                filetypes=[("Pliki logów", "*.log"), ("Wszystkie pliki", "*.*")]
            )
            
            if filename and self.log_text_area:
                content = self.log_text_area.get('1.0', tk.END)
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                messagebox.showinfo("Sukces", f"Logi zapisane do: {filename}")
                logging.info(f"Logi GUI zapisane do pliku: {filename}")
                
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się zapisać logów: {e}")
            logging.error(f"Błąd zapisywania logów: {e}")

    def add_log_message(self, message: str, level: str = "INFO") -> None:
        """
        Dodaje wiadomość do obszaru logów.
        
        Args:
            message: Treść wiadomości
            level: Poziom logu
        """
        if not self.log_text_area:
            return
            
        try:
            self.log_text_area.configure(state='normal')
            
            # Dodaj wiadomość z odpowiednim kolorem
            start_pos = self.log_text_area.index(tk.END)
            self.log_text_area.insert(tk.END, message + '\n')
            end_pos = self.log_text_area.index(tk.END)
            
            # Zastosuj formatowanie (jeśli działa)
            try:
                line_start = f"{start_pos.split('.')[0]}.0"
                line_end = f"{start_pos.split('.')[0]}.end"
                self.log_text_area.tag_add(level, line_start, line_end)
            except:
                pass  # Ignoruj błędy formatowania
            
            # Auto-przewijanie
            if self.auto_scroll_var.get():
                self.log_text_area.see(tk.END)
                
            self.log_text_area.configure(state='disabled')
            self.log_text_area.update_idletasks()
            
        except tk.TclError:
            # Widget został zniszczony
            pass


class TextHandler(logging.Handler):
    """Handler logowania dla widgetu tekstowego."""
    
    def __init__(self, text_widget, logs_tab):
        """
        Inicjalizuje handler.
        
        Args:
            text_widget: Widget tekstowy
            logs_tab: Referencja do zakładki logów
        """
        super().__init__()
        self.text_widget = text_widget
        self.logs_tab = logs_tab

    def emit(self, record: logging.LogRecord) -> None:
        """
        Emituje rekord logu.
        
        Args:
            record: Rekord logu
        """
        try:
            msg = self.format(record)
            self.logs_tab.add_log_message(msg, record.levelname)
        except Exception:
            # Nie loguj błędów handlera żeby uniknąć rekurencji
            pass