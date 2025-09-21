"""
Test przycisków - uproszczona wersja.
"""
import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys

# Dodaj ścieżkę
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

def test_button_1():
    messagebox.showinfo("Test", "Przycisk 1 działa!")

def test_button_2():
    messagebox.showinfo("Test", "Przycisk 2 działa!")

def main():
    root = tk.Tk()
    root.title("Test Przycisków")
    root.geometry("800x600")
    
    # Główny obszar
    main_frame = ttk.Frame(root)
    main_frame.pack(expand=True, fill='both', padx=10, pady=10)
    
    # Notebook
    notebook = ttk.Notebook(main_frame)
    notebook.pack(expand=True, fill='both', pady=(0, 10))
    
    # Zakładka testowa
    tab1 = ttk.Frame(notebook)
    notebook.add(tab1, text="Test Tab")
    
    ttk.Label(tab1, text="To jest test - czy widzisz przyciski na dole?").pack(pady=20)
    
    # PANEL PRZYCISKÓW - TO JEST KLUCZOWE
    print("Tworzenie panelu przycisków...")
    button_frame = ttk.Frame(root)  # UWAGA: parent to root, nie main_frame
    button_frame.pack(fill='x', padx=10, pady=5)
    
    # Przyciski
    print("Dodawanie przycisków...")
    ttk.Button(button_frame, text="Test 1", command=test_button_1).pack(side=tk.LEFT, padx=5)
    ttk.Button(button_frame, text="Test 2", command=test_button_2).pack(side=tk.LEFT, padx=5)
    ttk.Button(button_frame, text="Test 3", command=lambda: print("Test 3")).pack(side=tk.LEFT, padx=5)
    
    # Status bar
    status_frame = ttk.Frame(root)
    status_frame.pack(fill='x', padx=10, pady=5)
    
    status_label = ttk.Label(status_frame, text="Status: Gotowy", relief='sunken', anchor='w')
    status_label.pack(fill='x')
    
    print("Uruchamianie aplikacji...")
    root.mainloop()

if __name__ == "__main__":
    main()