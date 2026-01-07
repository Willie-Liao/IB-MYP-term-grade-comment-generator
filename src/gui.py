#!/usr/bin/env python3
"""GUI entry point for ManageBac Gradebook Scraper.

This script provides a graphical interface for entering credentials
and scraping gradebook data from ManageBac.

Usage:
    python3 src/gui.py
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import logging
import sys

from .exceptions import AuthenticationError, ScraperError
from .scraper import GradebookScraper


class LoginWindow:
    """GUI window for ManageBac login and scraping."""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ManageBac Gradebook Scraper")
        self.root.geometry("500x400")
        self.root.resizable(False, False)
        
        # Center the window
        self.root.eval('tk::PlaceWindow . center')
        
        self._create_widgets()
        self._setup_logging()
    
    def _setup_logging(self):
        """Configure logging to show in status area."""
        logging.basicConfig(
            level=logging.INFO,
            format="%(message)s",
            handlers=[logging.StreamHandler(sys.stderr)]
        )
    
    def _create_widgets(self):
        """Create all GUI widgets."""
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="ManageBac Gradebook Scraper",
            font=("Helvetica", 16, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Form frame
        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill=tk.X)
        
        # School code
        ttk.Label(form_frame, text="School Code:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.school_code_var = tk.StringVar()
        self.school_code_entry = ttk.Entry(form_frame, textvariable=self.school_code_var, width=40)
        self.school_code_entry.grid(row=0, column=1, pady=5, padx=(10, 0))
        ttk.Label(form_frame, text="(e.g., 'myschool' for myschool.managebac.cn)", 
                  font=("Helvetica", 9)).grid(row=1, column=1, sticky=tk.W, padx=(10, 0))
        
        # Email
        ttk.Label(form_frame, text="Email:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.email_var = tk.StringVar()
        self.email_entry = ttk.Entry(form_frame, textvariable=self.email_var, width=40)
        self.email_entry.grid(row=2, column=1, pady=5, padx=(10, 0))
        
        # Password
        ttk.Label(form_frame, text="Password:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(form_frame, textvariable=self.password_var, width=40, show="*")
        self.password_entry.grid(row=3, column=1, pady=5, padx=(10, 0))

        # Gradebook URL
        ttk.Label(form_frame, text="Gradebook URL:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.url_var = tk.StringVar()
        self.url_entry = ttk.Entry(form_frame, textvariable=self.url_var, width=40)
        self.url_entry.grid(row=4, column=1, pady=5, padx=(10, 0))
        ttk.Label(form_frame, text="(Full URL of the gradebook page)", 
                  font=("Helvetica", 9)).grid(row=5, column=1, sticky=tk.W, padx=(10, 0))
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        self.scrape_button = ttk.Button(
            button_frame, 
            text="Scrape Gradebook", 
            command=self._start_scraping
        )
        self.scrape_button.pack(side=tk.LEFT, padx=5)
        
        self.quit_button = ttk.Button(
            button_frame, 
            text="Quit", 
            command=self.root.quit
        )
        self.quit_button.pack(side=tk.LEFT, padx=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate', length=400)
        self.progress.pack(pady=10)
        
        # Status label
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var, wraplength=450)
        self.status_label.pack(pady=10)
    
    def _validate_inputs(self) -> bool:
        """Validate all input fields."""
        if not self.school_code_var.get().strip():
            messagebox.showerror("Error", "Please enter the school code")
            return False
        if not self.email_var.get().strip():
            messagebox.showerror("Error", "Please enter your email")
            return False
        if not self.password_var.get():
            messagebox.showerror("Error", "Please enter your password")
            return False
        if not self.url_var.get().strip():
            messagebox.showerror("Error", "Please enter the gradebook URL")
            return False
        return True
    
    def _start_scraping(self):
        """Start the scraping process in a background thread."""
        if not self._validate_inputs():
            return
        
        # Disable button and start progress
        self.scrape_button.config(state=tk.DISABLED)
        self.progress.start()
        self.status_var.set("Authenticating...")
        
        # Run scraping in background thread
        thread = threading.Thread(target=self._scrape_thread, daemon=True)
        thread.start()
    
    def _scrape_thread(self):
        """Background thread for scraping."""
        try:
            scraper = GradebookScraper(
                school_code=self.school_code_var.get().strip(),
                email=self.email_var.get().strip(),
                password=self.password_var.get()
            )
            
            self._update_status("Logging in to ManageBac...")
            scraper.authenticate()
            
            self._update_status("Scraping gradebook data...")
            output_file = scraper.scrape(
                gradebook_url=self.url_var.get().strip()
            )
            
            self._scrape_complete(output_file)
            
        except AuthenticationError as e:
            self._scrape_error(f"Authentication failed: {e}")
        except ScraperError as e:
            self._scrape_error(f"Scraping failed: {e}")
        except Exception as e:
            self._scrape_error(f"Unexpected error: {e}")
    
    def _update_status(self, message: str):
        """Update status label from any thread."""
        self.root.after(0, lambda: self.status_var.set(message))
    
    def _scrape_complete(self, output_file: str):
        """Handle successful scraping completion."""
        def complete():
            self.progress.stop()
            self.scrape_button.config(state=tk.NORMAL)
            self.status_var.set(f"Success! Exported to: {output_file}")
            messagebox.showinfo("Success", f"Gradebook exported to:\n{output_file}")
        self.root.after(0, complete)
    
    def _scrape_error(self, message: str):
        """Handle scraping error."""
        def error():
            self.progress.stop()
            self.scrape_button.config(state=tk.NORMAL)
            self.status_var.set(f"Error: {message}")
            messagebox.showerror("Error", message)
        self.root.after(0, error)
    
    def run(self):
        """Start the GUI main loop."""
        self.root.mainloop()


def main():
    """Main entry point for GUI."""
    app = LoginWindow()
    app.run()


if __name__ == "__main__":
    main()
