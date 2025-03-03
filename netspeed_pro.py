# Copyright (c) 2025 Md. Tanvir Hossain <shourve@gmail.com>
#
# This software is released under the MIT License.
# See the LICENSE file for the full text.

import psutil
import time
import tkinter as tk
from tkinter import ttk, colorchooser, messagebox
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from PIL import Image, ImageDraw, ImageTk
from pystray import MenuItem as item, Icon
import threading
import sys
import os
from collections import deque, defaultdict
from queue import Queue
import webbrowser
import datetime
import csv

if sys.platform == 'win32':
    import winreg
    import ctypes

class SmoothNetMonitor:
    def __init__(self):
        self.root = tk.Tk()
        self.setup_main_window()

        self.taskbar_mode = False
        self.lock_movement = tk.BooleanVar(self.root, value=False)
        self.settings_window = None
        self.data_usage_window = None

        self.setup_variables()
        self.load_daily_data()
        self.create_widgets()
        self.setup_bindings()
        self.setup_tray_icon()
        self.start_speed_thread()
        self.start_ui_update()
        self.root.mainloop()

    def setup_main_window(self):
        self.root.title("NetSpeed Pro")
        self.root.geometry("300x250")  # Initial size, will be repositioned below
        self.root.overrideredirect(True)
        self.root.attributes("-alpha", 0.9)
        self.root.attributes("-topmost", True)
        self.root.config(bg="#2a2a2a")
        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)

        try:
            self.icon_image = tk.PhotoImage(file="icon.png")
            self.root.iconphoto(True, self.icon_image)
        except Exception as e:
            print("Could not load icon:", e)

        # --- Code to Position Window on Startup ---
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 300  # Initial window width (from root.geometry)
        window_height = 250 # Initial window height (from root.geometry)

        # Calculate X position (more to the right side of screen)
        x_coordinate = screen_width - window_width - 20  # Increased right margin to 20 pixels

        # Calculate Y position (more to the bottom, further above taskbar - approximate taskbar height)
        approx_taskbar_height = 40 # Approximate taskbar height in pixels (adjust if needed)
        y_coordinate = screen_height - window_height - approx_taskbar_height - 20 # Increased bottom margin to 20 pixels

        self.root.geometry(f"+{x_coordinate}+{y_coordinate}") # Set the window position
        # --- End of Positioning Code ---

    def setup_variables(self):
        self.speed_queue = Queue()
        self.download_data = deque(maxlen=50)
        self.upload_data = deque(maxlen=50)
        self.update_interval = 1000
        self.ui_refresh_rate = 500
        self.colors = {
            'background': "#2a2a2a",
            'download': "#00ff00",
            'upload': "#ff0000",
            'text': "#ffffff",
            'graph_bg': "#2a2a2a",
            'scale_color': "#ffffff"
        }
        self.graph_type = "Line"
        self.selected_adapter = "All"
        self.running = True

        self.graph_title = "Network Speed Graph"
        self.graph_title_font_size = 12
        self.graph_text_color = "#ffffff"

        if sys.platform == 'win32':
            self.startup_var = tk.BooleanVar(value=True)
            if self.startup_var.get():
                self.set_startup(True)

        self.minimal_width = 230
        self.minimal_height = 30
        self.minimal_font_size = 10
        self.normal_speed_font = None

        self.daily_download_bytes = 0
        self.daily_upload_bytes = 0
        self.last_data_update_day = datetime.date.today()
        self.data_usage_file = "data_usage.csv"

        self.hourly_data = defaultdict(lambda: {'download': 0, 'upload': 0})

    def create_widgets(self):
        self.top_bar = tk.Frame(self.root, bg=self.colors['background'])
        self.top_bar.pack(fill=tk.X)

        self.menu_button = tk.Button(
            self.top_bar, text="≡", bg=self.colors['background'],
            fg=self.colors['text'], borderwidth=0, command=self.show_settings
        )
        self.menu_button.grid(row=0, column=0, padx=5)

        self.title_label = tk.Label(
            self.top_bar, text="NetSpeed Pro", bg=self.colors['background'],
            fg=self.colors['text']
        )
        self.title_label.grid(row=0, column=1, padx=5, sticky="w")

        self.speed_frame = tk.Frame(self.top_bar, bg=self.colors['background'])
        self.speed_frame.grid(row=0, column=2, padx=5, sticky="e")

        self.down_label = tk.Label(
            self.speed_frame, text="▼ 0 KB/s", fg=self.colors['download'], bg=self.colors['background']
        )
        self.down_label.pack(side=tk.LEFT, padx=5)
        self.up_label = tk.Label(
            self.speed_frame, text="▲ 0 KB/s", fg=self.colors['upload'], bg=self.colors['background']
        )
        self.up_label.pack(side=tk.LEFT, padx=5)

        self.normal_speed_font = self.down_label.cget("font")
        self.top_bar.grid_columnconfigure(1, weight=1)

        self.fig = plt.figure(figsize=(4, 3), facecolor=self.colors['graph_bg'])
        self.ax = self.fig.add_subplot(111, facecolor=self.colors['graph_bg'])
        self.down_line, = self.ax.plot([], [], color=self.colors['download'], label="Download")
        self.up_line, = self.ax.plot([], [], color=self.colors['upload'], label="Upload")
        self.ax.legend(loc="upper left", fontsize=8)
        self.canvas = FigureCanvasTkAgg(self.fig, self.root)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def setup_bindings(self):
        self.title_label.bind("<ButtonPress-1>", self.start_drag)
        self.title_label.bind("<B1-Motion>", self.on_drag)

    def start_drag(self, event):
        if self.lock_movement.get():
            return
        self.drag_offset_x = event.x
        self.drag_offset_y = event.y

    def on_drag(self, event):
        if self.lock_movement.get():
            return
        new_x = event.x_root - self.drag_offset_x
        new_y = event.y_root - self.drag_offset_y
        self.root.geometry(f"+{new_x}+{new_y}")

    def show_settings(self):
        if self.settings_window is not None and tk.Toplevel.winfo_exists(self.settings_window):
            self.settings_window.lift()
            return

        start_time = time.time()
        print("Starting show_settings at:", start_time)

        self.settings_window = tk.Toplevel(self.root)
        self.settings_window.title("Settings")
        self.settings_window.attributes("-topmost", True)
        self.settings_window.resizable(False, False)
        self.settings_window.geometry("450x400") # Increased size here
        time1 = time.time()
        print("Toplevel window created:", time1 - start_time)

        settings_pady = 8
        settings_padx = 10
        label_sticky = "w"

        # --- Notebook (Tabbed Interface) ---
        settings_notebook = ttk.Notebook(self.settings_window)
        settings_notebook.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # --- Window Appearance Tab ---
        appearance_tab = ttk.Frame(settings_notebook)
        settings_notebook.add(appearance_tab, text="Appearance")

        # Transparency Control
        ttk.Label(appearance_tab, text="Transparency:").grid(row=0, column=0, padx=settings_padx, pady=settings_pady, sticky=label_sticky)
        self.transparency_var = tk.StringVar(value=str(self.root.attributes("-alpha")))
        trans_scale = ttk.Scale(appearance_tab, from_=0.1, to=1.0, variable=tk.DoubleVar(value=self.root.attributes("-alpha")), command=self.update_transparency_value)
        trans_scale.grid(row=0, column=1, padx=settings_padx, pady=settings_pady, sticky="ew")
        self.transparency_label = ttk.Label(appearance_tab, textvariable=self.transparency_var, width=4)
        self.transparency_label.grid(row=0, column=2, padx=settings_padx, pady=settings_pady, sticky=label_sticky)

        # Always On Top
        ttk.Label(appearance_tab, text="Always On Top:").grid(row=1, column=0, padx=settings_padx, pady=settings_pady, sticky=label_sticky)
        always_on_top_var = tk.BooleanVar(value=self.root.attributes("-topmost"))
        ttk.Checkbutton(appearance_tab, variable=always_on_top_var, command=lambda: self.set_always_on_top(always_on_top_var.get())).grid(row=1, column=1, padx=settings_padx, pady=settings_pady, sticky=label_sticky)

        # Lock Movement
        ttk.Label(appearance_tab, text="Lock Movement:").grid(row=2, column=0, padx=settings_padx, pady=settings_pady, sticky=label_sticky)
        ttk.Checkbutton(appearance_tab, variable=self.lock_movement).grid(row=2, column=1, padx=settings_padx, pady=settings_pady, sticky=label_sticky)

        # Window Size
        ttk.Label(appearance_tab, text="Width:").grid(row=3, column=0, padx=settings_padx, pady=settings_pady, sticky=label_sticky)
        self.width_var = tk.StringVar(value=str(self.root.winfo_width()))
        width_entry = ttk.Entry(appearance_tab, textvariable=self.width_var, width=7)
        width_entry.grid(row=3, column=1, padx=settings_padx, pady=settings_pady, sticky="ew")
        ttk.Label(appearance_tab, text="Height:").grid(row=4, column=0, padx=settings_padx, pady=settings_pady, sticky=label_sticky)
        self.height_var = tk.StringVar(value=str(self.root.winfo_height()))
        height_entry = ttk.Entry(appearance_tab, textvariable=self.height_var, width=7)
        height_entry.grid(row=4, column=1, padx=settings_padx, pady=settings_pady, sticky="ew")
        apply_size_btn = ttk.Button(appearance_tab, text="Apply Size", command=self.apply_window_size_from_entry)
        apply_size_btn.grid(row=5, column=0, columnspan=3, padx=settings_padx, pady=settings_pady, sticky="ew")
        appearance_tab.columnconfigure(1, weight=1)


        # --- Graph Settings Tab ---
        graph_tab = ttk.Frame(settings_notebook)
        settings_notebook.add(graph_tab, text="Graph")

        # Update Interval
        ttk.Label(graph_tab, text="Update Interval (ms):").grid(row=0, column=0, padx=settings_padx, pady=settings_pady, sticky=label_sticky)
        self.interval_var = tk.StringVar(value=str(self.update_interval))
        interval_scale = ttk.Scale(graph_tab, from_=100, to=5000, variable=tk.DoubleVar(value=self.update_interval), command=self.update_interval_value)
        interval_scale.grid(row=0, column=1, padx=settings_padx, pady=settings_pady, sticky="ew")
        self.interval_entry = ttk.Entry(graph_tab, textvariable=self.interval_var, width=7)
        self.interval_entry.grid(row=0, column=2, padx=settings_padx, pady=settings_pady, sticky=label_sticky)
        self.interval_entry.bind("<FocusOut>", self.apply_interval_from_entry)
        self.interval_entry.bind("<Return>", self.apply_interval_from_entry)

        # Graph Type
        ttk.Label(graph_tab, text="Graph Type:").grid(row=1, column=0, padx=settings_padx, pady=settings_pady, sticky=label_sticky)
        graph_type_cb = ttk.Combobox(graph_tab, values=["Line", "Bar"], state="readonly")
        graph_type_cb.set(self.graph_type)
        graph_type_cb.grid(row=1, column=1, padx=settings_padx, pady=settings_pady, sticky="ew")
        graph_type_cb.bind("<<ComboboxSelected>>", lambda event: self.set_graph_type(graph_type_cb.get()))

        # Network Adapter
        ttk.Label(graph_tab, text="Network Adapter:").grid(row=2, column=0, padx=settings_padx, pady=settings_pady, sticky=label_sticky)
        adapter_cb = ttk.Combobox(graph_tab, values=[], state="readonly") # Empty initially
        adapter_cb.set(self.selected_adapter)
        adapter_cb.grid(row=2, column=1, padx=settings_padx, pady=settings_pady, sticky="ew")
        adapter_cb.bind("<<ComboboxSelected>>", lambda event: self.set_adapter(adapter_cb.get()))
        self.update_adapter_list_in_settings(adapter_cb)
        graph_tab.columnconfigure(1, weight=1)


        # --- Colors Tab ---
        colors_tab = ttk.Frame(settings_notebook)
        settings_notebook.add(colors_tab, text="Colors")

        # Background Color Button and Indicator
        btn_bg = ttk.Button(colors_tab, text="Background", command=lambda: self.choose_color('background'))
        btn_bg.grid(row=0, column=0, padx=settings_padx, pady=settings_pady, sticky="ew")
        self.bg_color_indicator = tk.Frame(colors_tab, width=20, height=20, bg=self.colors['background'], relief=tk.SOLID, borderwidth=1) # Color indicator frame
        self.bg_color_indicator.grid(row=0, column=1, padx=settings_padx, pady=settings_pady, sticky="w") # Placed to the right

        # Download Line/Text Color Button and Indicator
        btn_dl = ttk.Button(colors_tab, text="Download Line/Text", command=lambda: self.choose_color('download'))
        btn_dl.grid(row=0, column=2, padx=settings_padx, pady=settings_pady, sticky="ew")
        self.dl_color_indicator = tk.Frame(colors_tab, width=20, height=20, bg=self.colors['download'], relief=tk.SOLID, borderwidth=1) # Color indicator frame
        self.dl_color_indicator.grid(row=0, column=3, padx=settings_padx, pady=settings_pady, sticky="w") # Placed to the right

        # Upload Line/Text Color Button and Indicator
        btn_ul = ttk.Button(colors_tab, text="Upload Line/Text", command=lambda: self.choose_color('upload'))
        btn_ul.grid(row=1, column=0, padx=settings_padx, pady=settings_pady, sticky="ew")
        self.ul_color_indicator = tk.Frame(colors_tab, width=20, height=20, bg=self.colors['upload'], relief=tk.SOLID, borderwidth=1) # Color indicator frame
        self.ul_color_indicator.grid(row=1, column=1, padx=settings_padx, pady=settings_pady, sticky="w") # Placed to the right

        # Graph Background Color Button and Indicator
        btn_graph = ttk.Button(colors_tab, text="Graph Background", command=lambda: self.choose_color('graph_bg'))
        btn_graph.grid(row=1, column=2, padx=settings_padx, pady=settings_pady, sticky="ew")
        self.graph_bg_color_indicator = tk.Frame(colors_tab, width=20, height=20, bg=self.colors['graph_bg'], relief=tk.SOLID, borderwidth=1) # Color indicator frame
        self.graph_bg_color_indicator.grid(row=1, column=3, padx=settings_padx, pady=settings_pady, sticky="w") # Placed to the right

        # Scale Color Button and Indicator
        btn_scale = ttk.Button(colors_tab, text="Scale Color", command=lambda: self.choose_color('scale_color'))
        btn_scale.grid(row=2, column=0, columnspan=2, padx=settings_padx, pady=settings_pady, sticky="ew")
        self.scale_color_indicator = tk.Frame(colors_tab, width=20, height=20, bg=self.colors['scale_color'], relief=tk.SOLID, borderwidth=1) # Color indicator frame
        self.scale_color_indicator.grid(row=2, column=2, padx=settings_padx, pady=settings_pady, sticky="w") # Placed to the right
        colors_tab.columnconfigure(0, weight=1)
        colors_tab.columnconfigure(2, weight=1) # Add weight to make buttons expand


        # --- Minimal Mode Tab ---
        minimal_tab = ttk.Frame(settings_notebook)
        settings_notebook.add(minimal_tab, text="Minimal Mode")

        ttk.Label(minimal_tab, text="Toggle Display Size (WxH):").grid(row=0, column=0, padx=settings_padx, pady=settings_pady, sticky=label_sticky)
        frame_min_size = ttk.Frame(minimal_tab)
        frame_min_size.grid(row=0, column=1, padx=settings_padx, pady=settings_pady, sticky="ew")
        width_min_entry = ttk.Entry(frame_min_size, width=6)
        width_min_entry.insert(0, str(self.minimal_width))
        width_min_entry.pack(side=tk.LEFT, padx=(0,5))
        width_min_entry.bind("<FocusOut>", self.apply_minimal_width)
        height_min_entry = ttk.Entry(frame_min_size, width=6)
        height_min_entry.insert(0, str(self.minimal_height))
        height_min_entry.pack(side=tk.LEFT)
        height_min_entry.bind("<FocusOut>", self.apply_minimal_height)

        ttk.Label(minimal_tab, text="Toggle Font Size:").grid(row=1, column=0, padx=settings_padx, pady=settings_pady, sticky=label_sticky)
        toggle_font_entry = ttk.Entry(minimal_tab, width=6)
        toggle_font_entry.insert(0, str(self.minimal_font_size))
        toggle_font_entry.grid(row=1, column=1, padx=settings_padx, pady=settings_pady, sticky="ew")
        toggle_font_entry.bind("<FocusOut>", self.apply_minimal_font_size)
        minimal_tab.columnconfigure(1, weight=1)


        # --- Other Options Tab ---
        other_tab = ttk.Frame(settings_notebook)
        settings_notebook.add(other_tab, text="Other")

        if sys.platform == 'win32': # Windows Startup
            ttk.Label(other_tab, text="Start with Windows:").grid(row=0, column=0, padx=settings_padx, pady=settings_pady, sticky=label_sticky)
            startup_cb = ttk.Checkbutton(other_tab, variable=self.startup_var, command=self.update_startup)
            startup_cb.grid(row=0, column=1, padx=settings_padx, pady=settings_pady, sticky=label_sticky)

        btn_graph_text = ttk.Button(other_tab, text="Graph Text Customization", command=self.show_graph_text_settings)
        btn_graph_text.grid(row=1, column=0, columnspan=2, padx=settings_padx, pady=settings_pady, sticky="ew")
        btn_toggle = ttk.Button(other_tab, text="Toggle Display", command=self.toggle_taskbar_display)
        btn_toggle.grid(row=2, column=0, columnspan=2, padx=settings_padx, pady=settings_pady, sticky="ew")
        btn_about = ttk.Button(other_tab, text="About", command=self.show_about)
        btn_about.grid(row=3, column=0, columnspan=2, padx=settings_padx, pady=settings_pady, sticky="ew")
        btn_exit = ttk.Button(other_tab, text="Exit", command=self.clean_exit)
        btn_exit.grid(row=4, column=0, columnspan=2, padx=settings_padx, pady=settings_pady, sticky="ew")
        other_tab.columnconfigure(0, weight=1)


        self.settings_window.columnconfigure(0, weight=1)
        self.settings_window.protocol("WM_DELETE_WINDOW", self.close_settings)

        end_time = time.time()
        print("Settings window setup complete:", end_time - time1)
        print("Total show_settings execution time:", end_time - start_time)


    def update_transparency_value(self, value):
        self.root.attributes("-alpha", float(value))
        self.transparency_var.set(f"{float(value):.2f}") # Update label value

    def update_interval_value(self, value):
        self.update_interval = int(float(value))
        self.interval_var.set(str(self.update_interval)) # Update entry value

    def apply_interval_from_entry(self, event=None): # Apply interval from Entry
        value_str = self.interval_var.get()
        try:
            value = int(value_str)
            if 100 <= value <= 5000:
                self.update_interval = value
                interval_scale = self.settings_window.winfo_children()[2] # Get the scale - careful with index if structure changes
                interval_scale.set(value) # Update the scale to match entry
            else:
                messagebox.showerror("Invalid Interval", "Interval must be between 100 and 5000 ms.")
                self.interval_var.set(str(self.update_interval)) # Revert entry value
        except ValueError:
            messagebox.showerror("Invalid Interval", "Interval must be an integer.")
            self.interval_var.set(str(self.update_interval)) # Revert entry value

    def apply_window_size_from_entry(self): # Apply window size from Entries
        width_str = self.width_var.get()
        height_str = self.height_var.get()
        try:
            w = int(width_str)
            h = int(height_str)
            if w > 0 and h > 0: # Basic validation: width and height must be positive
                self.root.geometry(f"{w}x{h}")
            else:
                messagebox.showerror("Invalid Size", "Width and Height must be positive integers.")
                self.width_var.set(str(self.root.winfo_width()))    # Revert entry value
                self.height_var.set(str(self.root.winfo_height())) # Revert entry value

        except ValueError:
            messagebox.showerror("Invalid Size", "Width and Height must be integers.")
            self.width_var.set(str(self.root.winfo_width()))    # Revert entry value
            self.height_var.set(str(self.root.winfo_height())) # Revert entry value


    def update_adapter_list_in_settings(self, adapter_combobox):
        threading.Thread(target=self._get_adapters_threaded, args=(adapter_combobox,), daemon=True).start()

    def _get_adapters_threaded(self, adapter_combobox):
        try:
            adapters = list(psutil.net_io_counters(pernic=True).keys())
            adapters.sort()
            adapters.insert(0, "All")
        except Exception:
            adapters = ["All", "Error retrieving adapters"]  # Fallback in case of error

        self.root.after(0, lambda: self._update_adapter_combobox_callback(adapter_combobox, adapters))

    def _update_adapter_combobox_callback(self, adapter_combobox, adapters):
        adapter_combobox['values'] = adapters
        if self.selected_adapter in adapters:
            adapter_combobox.set(self.selected_adapter)
        else:
            adapter_combobox.set("All") # Default if selected adapter is not available


    def close_settings(self):
        if self.settings_window:
            self.settings_window.destroy()
            self.settings_window = None

    def apply_minimal_width(self, event=None):
        w_str = event.widget.get()
        try:
            self.minimal_width = int(w_str)
        except ValueError:
            messagebox.showerror("Invalid Width", "Width must be an integer.")
            event.widget.delete(0, tk.END)
            event.widget.insert(0, str(self.minimal_width))

    def apply_minimal_height(self, event=None):
        h_str = event.widget.get()
        try:
            self.minimal_height = int(h_str)
        except ValueError:
            messagebox.showerror("Invalid Height", "Height must be an integer.")
            event.widget.delete(0, tk.END)
            event.widget.insert(0, str(self.minimal_height))

    def apply_minimal_font_size(self, event=None):
        font_str = event.widget.get()
        try:
            self.minimal_font_size = int(font_str)
        except ValueError:
            messagebox.showerror("Invalid Font Size", "Font size must be an integer.")
            event.widget.delete(0, tk.END)
            event.widget.insert(0, str(self.minimal_font_size))

    def set_window_size(self, width, height):
        try:
            w = int(width)
            h = int(height)
            self.root.geometry(f"{w}x{h}")
        except ValueError:
            messagebox.showerror("Invalid Size", "Width and Height must be integers.")
    def set_always_on_top(self, always_top):
        self.root.attributes("-topmost", always_top)

    def set_graph_type(self, graph_type):
        self.graph_type = graph_type
        self.update_graph() # Update graph to reflect type change

    def set_adapter(self, adapter_name):
        self.selected_adapter = adapter_name

    def choose_color(self, setting_name):
        initial_color = self.colors.get(setting_name, "#ffffff") # Default white if not found
        color_code = colorchooser.askcolor(initialcolor=initial_color)
        if color_code and color_code[1]:
            color_hex = color_code[1]
            self.colors[setting_name] = color_hex
            if setting_name == 'background':
                self.root.config(bg=color_hex)
                self.top_bar.config(bg=color_hex)
                self.menu_button.config(bg=color_hex)
                self.title_label.config(bg=color_hex)
                self.speed_frame.config(bg=color_hex)
                self.bg_color_indicator.config(bg=color_hex) # Update indicator color
                if self.settings_window: # Check if settings window exists to avoid errors
                    settings_bg = self.settings_window.winfo_children()[0] # Notebook is the first child
                    if settings_bg:
                        settings_bg.config(bg=color_hex) # Apply to notebook
            elif setting_name == 'download':
                self.down_label.config(fg=color_hex)
                self.down_line.set_color(color_hex) # Update graph line color
                self.dl_color_indicator.config(bg=color_hex) # Update indicator color
                self.update_graph() # Redraw graph to apply color change
            elif setting_name == 'upload':
                self.up_label.config(fg=color_hex)
                self.up_line.set_color(color_hex) # Update graph line color
                self.ul_color_indicator.config(bg=color_hex) # Update indicator color
                self.update_graph() # Redraw graph to apply color change
            elif setting_name == 'graph_bg':
                self.ax.set_facecolor(color_hex)
                self.fig.patch.set_facecolor(color_hex)
                self.canvas.draw() # Redraw the canvas to update background color
                self.graph_bg_color_indicator.config(bg=color_hex) # Update indicator color
            elif setting_name == 'scale_color':
                self.set_scale_color(color_hex)
                self.scale_color_indicator.config(bg=color_hex) # Update indicator color
        self.save_settings() # Save settings after color change

    def set_scale_color(self, color_hex):
        self.colors['scale_color'] = color_hex
        self.ax.tick_params(axis='x', colors=color_hex)
        self.ax.tick_params(axis='y', colors=color_hex)
        self.update_data_usage_graph() # Update graph to apply scale color change


    def show_graph_text_settings(self):
        if hasattr(self, 'graph_text_window') and self.graph_text_window and tk.Toplevel.winfo_exists(self.graph_text_window):
            self.graph_text_window.lift()
            return

        self.graph_text_window = tk.Toplevel(self.settings_window) # Child of settings window
        self.graph_text_window.title("Graph Text Settings")
        self.graph_text_window.geometry("300x200")
        self.graph_text_window.resizable(False, False)
        self.graph_text_window.attributes("-topmost", True)

        settings_pady = 8
        settings_padx = 10
        label_sticky = "w"

        # Graph Title Entry
        ttk.Label(self.graph_text_window, text="Graph Title:").grid(row=0, column=0, padx=settings_padx, pady=settings_pady, sticky=label_sticky)
        self.graph_title_var = tk.StringVar(value=self.graph_title)
        title_entry = ttk.Entry(self.graph_text_window, textvariable=self.graph_title_var)
        title_entry.grid(row=0, column=1, padx=settings_padx, pady=settings_pady, sticky="ew")
        title_entry.bind("<FocusOut>", self.apply_graph_title)

        # Font Size Control
        ttk.Label(self.graph_text_window, text="Title Font Size:").grid(row=1, column=0, padx=settings_padx, pady=settings_pady, sticky=label_sticky)
        font_size_scale = ttk.Scale(self.graph_text_window, from_=8, to=20, orient=tk.HORIZONTAL, command=self.update_graph_font_size_value, variable=tk.IntVar(value=self.graph_title_font_size))
        font_size_scale.grid(row=1, column=1, padx=settings_padx, pady=settings_pady, sticky="ew")
        self.font_size_label = ttk.Label(self.graph_text_window, text=str(self.graph_title_font_size), width=3) # Label to display font size value
        self.font_size_label.grid(row=1, column=2, padx=settings_padx, pady=settings_pady, sticky=label_sticky)

        # Text Color Button
        btn_text_color = ttk.Button(self.graph_text_window, text="Text Color", command=self.choose_graph_text_color)
        btn_text_color.grid(row=2, column=0, columnspan=3, padx=settings_padx, pady=settings_pady, sticky="ew")

        self.graph_text_window.columnconfigure(1, weight=1)
        self.graph_text_window.protocol("WM_DELETE_WINDOW", self.close_graph_text_settings)

    def update_graph_font_size_value(self, value):
        self.graph_title_font_size = int(value)
        self.font_size_label.config(text=str(self.graph_title_font_size)) # Update font size label
        self.update_graph_text_options() # Apply changes to graph

    def apply_graph_title(self, event=None):
        self.graph_title = self.graph_title_var.get()
        self.update_graph_text_options() # Apply changes to graph

    def choose_graph_text_color(self):
        initial_color = self.graph_text_color
        color_code = colorchooser.askcolor(initialcolor=initial_color)
        if color_code and color_code[1]:
            self.graph_text_color = color_code[1]
            self.update_graph_text_options() # Apply changes to graph color

    def update_graph_text_options(self):
        self.ax.set_title(self.graph_title, fontsize=self.graph_title_font_size, color=self.graph_text_color)
        self.canvas.draw_idle() # Update graph title and redraw

    def close_graph_text_settings(self):
        if self.graph_text_window:
            self.graph_text_window.destroy()
            self.graph_text_window = None
    def set_always_on_top(self, value):
        self.root.attributes("-topmost", value)

    def choose_color(self, element):
        color = colorchooser.askcolor(title=f"Choose {element} color")
        if color[1]:
            self.colors[element] = color[1]
            if element == 'background':
                self.root.config(bg=color[1])
                self.top_bar.config(bg=color[1])
                self.menu_button.config(bg=color[1])
                self.title_label.config(bg=color[1])
                self.speed_frame.config(bg=color[1])
                self.down_label.config(bg=color[1])
                self.up_label.config(bg=color[1])
                self.bg_color_indicator.config(bg=color[1]) # Update indicator color
            elif element == 'download':
                self.down_label.config(fg=color[1])
                self.down_line.set_color(color[1])
                self.dl_color_indicator.config(bg=color[1]) # Update indicator color
            elif element == 'upload':
                self.up_label.config(fg=color[1])
                self.up_line.set_color(color[1])
                self.ul_color_indicator.config(bg=color[1]) # Update indicator color
            elif element == 'graph_bg':
                self.ax.set_facecolor(color[1])
                self.fig.patch.set_facecolor(color[1])
                self.graph_bg_color_indicator.config(bg=color[1]) # Update indicator color
            elif element == 'scale_color':
                self.set_scale_color(color[1])
                self.scale_color_indicator.config(bg=color[1]) # Update indicator color
            self.canvas.draw_idle()

    def set_scale_color(self, color_value):
        self.colors['scale_color'] = color_value
        self.update_graph()
        self.update_data_usage_graph()

    def set_graph_type(self, value):
        self.graph_type = value
        self.update_graph()

    def set_adapter(self, value):
        self.selected_adapter = value
        self.download_data.clear()
        self.upload_data.clear()

    def update_startup(self):
        if self.startup_var.get():
            self.set_startup(True)
        else:
            self.set_startup(False)

    def set_startup(self, enable):
        if sys.platform != 'win32':
            return
        try:
            reg_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_ALL_ACCESS)
            if enable:
                winreg.SetValueEx(key, "NetSpeedPro", 0, winreg.REG_SZ, os.path.abspath(sys.argv[0]))
            else:
                try:
                    winreg.DeleteValue(key, "NetSpeedPro")
                except FileNotFoundError:
                    pass
            winreg.CloseKey(key)
        except Exception as e:
            print("Error updating startup setting:", e)

    def show_about(self):
        about = tk.Toplevel(self.root)
        about.title("About")
        about.attributes("-topmost", True)
        about.geometry("400x300")
        about.resizable(False, False)

        try:
            photo_image = Image.open("my_photo.png")
            photo_image = photo_image.resize((100, 100), Image.Resampling.LANCZOS)
            self.tk_photo = ImageTk.PhotoImage(photo_image)
            photo_label = ttk.Label(about, image=self.tk_photo)
            photo_label.pack(pady=10)
        except FileNotFoundError:
            ttk.Label(about, text="Photo not found! Please place 'my_photo.png' in the same directory.", wraplength=200).pack(pady=5)
        except Exception as e:
            ttk.Label(about, text=f"Error loading photo: {e}", wraplength=200).pack(pady=5)

        ttk.Label(about, text="Network Monitor Pro", font=('Arial', 12)).pack(pady=2)
        ttk.Label(about, text="Developer: Md. Tanvir Hossain").pack(pady=2)
        ttk.Label(about, text="Email: tanvirofficial.242@gmail.com").pack(pady=2)

        linkedin_link = ttk.Button(
            about, text="LinkedIn", style="Link.TButton",
            command=lambda: webbrowser.open_new_tab("https://www.linkedin.com/in/tanvir016/")
        )
        linkedin_link.pack(pady=2)

        website_link = ttk.Button(
            about, text="Website", style="Link.TButton",
            command=lambda: webbrowser.open_new_tab("www.boostguru.xyz")
        )
        website_link.pack(pady=2)

        ttk.Button(about, text="Close", command=about.destroy).pack(pady=10)

        style = ttk.Style(about)
        style.configure("Link.TButton", foreground="blue", relief=tk.FLAT,  background=about.cget("background"), font=('Arial', 10, 'underline'))
        style.map("Link.TButton", foreground=[("active", "dark blue"), ("pressed", "blue")])

    def show_graph_text_settings(self):
        text_settings = tk.Toplevel(self.root)
        text_settings.title("Graph Text Customization")
        text_settings.attributes("-topmost", True)
        text_settings.resizable(False, False)

        ttk.Label(text_settings, text="Title Font Size:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        font_size_entry = ttk.Entry(text_settings)
        font_size_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        font_size_entry.insert(0, str(self.graph_title_font_size))

        btn_text_color = ttk.Button(text_settings, text="Change Title Color", command=self.change_graph_text_color)
        btn_text_color.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        apply_font_size_btn = ttk.Button(text_settings, text="Apply Font Size Setting", command=lambda: self.set_graph_text("", font_size_entry.get()))
        apply_font_size_btn.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

    def change_graph_text_color(self):
        color = colorchooser.askcolor(title="Choose Graph Text Color")
        if color[1]:
            self.graph_text_color = color[1]
            self.update_graph()

    def set_graph_text(self, title, font_size):
        try:
            self.graph_title_font_size = int(font_size)
        except ValueError:
            messagebox.showerror("Invalid Font Size", "Font size must be an integer.")
        self.update_graph()

    def update_graph(self):
        self.ax.clear()
        if self.graph_type == "Line":
            self.ax.plot(range(len(self.download_data)), list(self.download_data), color=self.colors['download'], label="Download")
            self.ax.plot(range(len(self.upload_data)), list(self.upload_data), color=self.colors['upload'], label="Upload")
        elif self.graph_type == "Bar":
            indices = list(range(len(self.download_data)))
            width = 0.4
            self.ax.bar([i - width/2 for i in indices], list(self.download_data), width=width, color=self.colors['download'], label="Download")
            self.ax.bar([i + width/2 for i in indices], list(self.upload_data), width=width, color=self.colors['upload'], label="Upload")
        self.ax.legend(loc="upper left", fontsize=8)
        self.ax.set_facecolor(self.colors['graph_bg'])
        self.fig.patch.set_facecolor(self.colors['graph_bg'])
        self.ax.relim()
        self.ax.autoscale_view()
        self.ax.set_title(self.graph_title, fontsize=self.graph_title_font_size, color=self.graph_text_color)
        self.ax.tick_params(axis='x', colors=self.colors['scale_color'])
        self.ax.tick_params(axis='y', colors=self.colors['scale_color'])

    def toggle_graph(self):
        if self.canvas.get_tk_widget().winfo_ismapped():
            self.root.geometry("300x30")
            self.canvas.get_tk_widget().pack_forget()
        else:
            self.root.geometry("300x250")
            self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def toggle_taskbar_display(self):
        if not self.taskbar_mode:
            self.taskbar_mode = True
            self.canvas.get_tk_widget().pack_forget()
            self.root.geometry(f"{self.minimal_width}x{self.minimal_height}")
            self.root.overrideredirect(True)

            self.down_label.config(font=("TkDefaultFont", self.minimal_font_size))
            self.up_label.config(font=("TkDefaultFont", self.minimal_font_size))
            self.title_label.grid_forget()

            self.root.bind("<ButtonPress-1>", self.minimal_start_drag)
            self.root.bind("<B1-Motion>", self.minimal_on_drag)

            if sys.platform == 'win32':
                self.force_show_in_taskbar()
        else:
            self.taskbar_mode = False
            self.root.geometry("300x250")
            self.down_label.config(font=self.normal_speed_font)
            self.up_label.config(font=self.normal_speed_font)
            self.title_label.grid(row=0, column=1, padx=5, sticky="w")
            self.root.unbind("<ButtonPress-1>")
            self.root.unbind("<B1-Motion>")
            self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def minimal_start_drag(self, event):
        if self.lock_movement.get():
            return
        self.drag_offset_x = event.x
        self.drag_offset_y = event.y

    def minimal_on_drag(self, event):
        if self.lock_movement.get():
            return
        new_x = event.x_root - self.drag_offset_x
        new_y = event.y_root - self.drag_offset_y
        self.root.geometry(f"+{new_x}+{new_y}")

    def force_show_in_taskbar(self):
        hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
        GWL_EXSTYLE = -20
        WS_EX_APPWINDOW = 0x00040000
        style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        style |= WS_EX_APPWINDOW
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)

    def setup_tray_icon(self):
        menu = (
            item('Show/Hide Graph', self.toggle_graph),
            item('Data Used', self.show_data_usage_window),
            item('Settings', self.show_settings),
            item('About', self.show_about),
            item('Exit', self.clean_exit)
        )
        try:
            tray_image = Image.open("icon.png").convert("RGBA")
        except:
            tray_image = Image.new('RGB', (64, 64), (42, 42, 42))
            draw = ImageDraw.Draw(tray_image)
            draw.ellipse((16, 16, 48, 48), fill=self.colors['download'])

        self.tray_icon = Icon("netspeed", tray_image, menu=menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def start_speed_thread(self):
        self.speed_thread = threading.Thread(target=self.measure_speeds, daemon=True)
        self.speed_thread.start()

    def measure_speeds(self):
        if self.selected_adapter == "All":
            old_stats = psutil.net_io_counters()
        else:
            try:
                old_stats = psutil.net_io_counters(pernic=True)[self.selected_adapter]
            except KeyError:
                old_stats = psutil.net_io_counters()
        while self.running:
            time.sleep(self.update_interval / 1000)
            if self.selected_adapter == "All":
                new_stats = psutil.net_io_counters()
            else:
                try:
                    new_stats = psutil.net_io_counters(pernic=True)[self.selected_adapter]
                except KeyError:
                    new_stats = psutil.net_io_counters()
            down_kbps = (new_stats.bytes_recv - old_stats.bytes_recv) / 1024
            up_kbps = (new_stats.bytes_sent - old_stats.bytes_sent) / 1024
            old_stats = new_stats

            current_datetime = datetime.datetime.now()
            current_day = current_datetime.date()
            current_hour = current_datetime.hour

            if current_day != self.last_data_update_day:
                self.save_daily_data()
                self.daily_download_bytes = 0
                self.daily_upload_bytes = 0
                self.last_data_update_day = current_day
                self.hourly_data.clear()

            interval_seconds = self.update_interval / 1000.0
            download_bytes_interval = int(down_kbps * 1024 * interval_seconds)
            upload_bytes_interval = int(up_kbps * 1024 * interval_seconds)

            self.daily_download_bytes += download_bytes_interval
            self.daily_upload_bytes += upload_bytes_interval

            self.hourly_data[current_hour]['download'] += download_bytes_interval
            self.hourly_data[current_hour]['upload'] += upload_bytes_interval


            self.speed_queue.put((down_kbps, up_kbps))

    def start_ui_update(self):
        self.update_labels()
        self.root.after(self.ui_refresh_rate, self.start_ui_update)

    def update_labels(self):
        try:
            while True:
                down, up = self.speed_queue.get_nowait()
                if down >= 1024:
                    down_disp = down / 1024
                    down_unit = "MB/s"
                else:
                    down_disp = down
                    down_unit = "KB/s"
                if up >= 1024:
                    up_disp = up / 1024
                    up_unit = "MB/s"
                else:
                    up_disp = up
                    up_unit = "KB/s"
                self.down_label.config(text=f"▼ {down_disp:.2f} {down_unit}")
                self.up_label.config(text=f"▲ {up_disp:.2f} {up_unit}")
                self.download_data.append(down)
                self.upload_data.append(up)
        except:
            pass
        self.update_graph()
        self.canvas.draw_idle()
        self.update_daily_usage_display()

    def minimize_to_tray(self):
        self.root.withdraw()

    def clean_exit(self):
        self.running = False
        self.tray_icon.stop()
        self.save_daily_data()
        self.save_hourly_data()
        self.root.destroy()
        os._exit(0)

    def format_bytes(self, bytes_val):
        units = ['Bytes', 'KB', 'MB', 'GB', 'TB']
        unit_index = 0
        while bytes_val >= 1024 and unit_index < len(units) - 1:
            bytes_val /= 1024
            unit_index += 1
        return f"{bytes_val:.2f} {units[unit_index]}"

    def update_daily_usage_display(self):
        if self.data_usage_window and tk.Toplevel.winfo_exists(self.data_usage_window): # Check if window and labels exist
            if hasattr(self, 'daily_down_usage_label') and hasattr(self, 'daily_up_usage_label'): # Double check labels exist
                self.daily_down_usage_label.config(text=f"Download: {self.format_bytes(self.daily_download_bytes)}")
                self.daily_up_usage_label.config(text=f"Upload: {self.format_bytes(self.daily_upload_bytes)}")

    def load_daily_data(self):
        if os.path.exists(self.data_usage_file):
            try:
                with open(self.data_usage_file, 'r', newline='') as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        file_date = datetime.datetime.strptime(row['Date'], '%Y-%m-%d').date()
                        if file_date == datetime.date.today():
                            self.daily_download_bytes = int(row['DownloadBytes'])
                            self.daily_upload_bytes = int(row['UploadBytes'])
                            self.last_data_update_day = file_date
                            return
            except Exception as e:
                print(f"Error loading daily data: {e}")

        self.daily_download_bytes = 0
        self.daily_upload_bytes = 0
        self.last_data_update_day = datetime.date.today()
        self.hourly_data.clear()
        self.load_hourly_data()


    def save_daily_data(self):
        try:
            with open(self.data_usage_file, 'w', newline='') as csvfile:
                fieldnames = ['Date', 'DownloadBytes', 'UploadBytes']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerow({
                    'Date': self.last_data_update_day.strftime('%Y-%m-%d'),
                    'DownloadBytes': str(self.daily_download_bytes),
                    'UploadBytes': str(self.daily_upload_bytes)
                })
        except Exception as e:
            print(f"Error saving daily data: {e}")

    def show_data_usage_window(self):
        if self.data_usage_window is not None and tk.Toplevel.winfo_exists(self.data_usage_window):
            self.data_usage_window.lift()
            return

        self.data_usage_window = tk.Toplevel(self.root)
        self.data_usage_window.title("Data Usage")
        self.data_usage_window.attributes("-topmost", True)
        self.data_usage_window.resizable(True, True) # <-- Data Usage window is now resizable

        # Timeframe Selection
        ttk.Label(self.data_usage_window, text="Timeframe:").pack(pady=5)
        timeframe_options = ["Hourly", "Daily", "Weekly", "Monthly"]
        self.timeframe_var = tk.StringVar(value="Daily") # Initialize timeframe_var here again, ensuring default value
        timeframe_cb = ttk.Combobox(self.data_usage_window, values=timeframe_options, textvariable=self.timeframe_var, state="readonly")
        timeframe_cb.pack(pady=5)
        timeframe_cb.bind("<<ComboboxSelected>>", self.update_data_usage_graph)

        # Usage Labels (Create them here)
        self.daily_down_usage_label = ttk.Label(self.data_usage_window, text="Download: ") # Initialize labels here
        self.daily_down_usage_label.pack()
        self.daily_up_usage_label = ttk.Label(self.data_usage_window, text="Upload: ") # Initialize labels here
        self.daily_up_usage_label.pack()

        # Graph Area
        self.usage_fig = plt.Figure(figsize=(6, 4), facecolor=self.colors['graph_bg'])
        self.usage_ax = self.usage_fig.add_subplot(111, facecolor=self.colors['graph_bg'])
        self.usage_canvas = FigureCanvasTkAgg(self.usage_fig, self.data_usage_window)
        self.usage_canvas_widget = self.usage_canvas.get_tk_widget()
        self.usage_canvas_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.update_data_usage_graph()
        self.data_usage_window.protocol("WM_DELETE_WINDOW", self.close_data_usage_window)

    def close_data_usage_window(self):
        if self.data_usage_window:
            self.data_usage_window.destroy()
            self.data_usage_window = None

    def update_data_usage_graph(self, event=None):
        timeframe = self.timeframe_var.get()
        self.usage_ax.clear()

        if timeframe == "Hourly":
            self.plot_hourly_usage()
        elif timeframe == "Daily":
            self.plot_daily_usage_graph()
        elif timeframe == "Weekly":
            self.plot_weekly_usage()
        elif timeframe == "Monthly":
            self.plot_monthly_usage()

        self.usage_ax.set_facecolor(self.colors['graph_bg'])
        self.usage_fig.patch.set_facecolor(self.colors['graph_bg'])
        self.usage_ax.tick_params(axis='x', colors=self.colors['scale_color'])
        self.usage_ax.tick_params(axis='y', colors=self.colors['scale_color'])
        self.usage_ax.set_title(f"{timeframe} Data Usage", fontsize=self.graph_title_font_size, color=self.graph_text_color)
        self.usage_ax.set_xlabel("Time", color=self.colors['text'])
        self.usage_ax.set_ylabel("Data Usage", color=self.colors['text'])
        self.usage_ax.legend(loc="upper right", fontsize=8, facecolor=self.colors['graph_bg'], edgecolor=self.colors['text'], labelcolor=self.colors['text'])
        self.usage_canvas.draw_idle()

    def plot_daily_usage_graph(self):
        dates = []
        download_values = []
        upload_values = []
        today = datetime.date.today()
        for i in range(7):
            date = today - datetime.timedelta(days=i)
            dates.append(date.strftime('%Y-%m-%d'))
            daily_data = self.load_data_for_date(date)
            download_values.append(daily_data['download'])
            upload_values.append(daily_data['upload'])
        dates.reverse()
        download_values.reverse()
        upload_values.reverse()

        x_positions = range(len(dates))
        width = 0.35

        self.usage_ax.bar([pos - width/2 for pos in x_positions], download_values, width, label='Download', color=self.colors['download'])
        self.usage_ax.bar([pos + width/2 for pos in x_positions], upload_values, width, label='Upload', color=self.colors['upload'])

        self.usage_ax.set_xticks(x_positions)
        self.usage_ax.set_xticklabels([date.split('-')[2] for date in dates])

    def load_data_for_date(self, date):
        download_bytes = 0
        upload_bytes = 0
        date_str = date.strftime('%Y-%m-%d')
        if os.path.exists(self.data_usage_file):
            try:
                with open(self.data_usage_file, 'r', newline='') as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        if row['Date'] == date_str:
                            download_bytes = int(row['DownloadBytes'])
                            upload_bytes = int(row['UploadBytes'])
                            break
            except Exception as e:
                print(f"Error loading data for date {date_str}: {e}")
        return {'download': download_bytes / (1024*1024), 'upload': upload_bytes / (1024*1024)}

    def plot_hourly_usage(self):
        hours = list(range(24))
        download_usage = [self.hourly_data[hour]['download'] / (1024*1024) for hour in hours]
        upload_usage = [self.hourly_data[hour]['upload'] / (1024*1024) for hour in hours]

        x_positions = range(len(hours))
        width = 0.35

        self.usage_ax.bar([pos - width/2 for pos in x_positions], download_usage, width, label='Download', color=self.colors['download'])
        self.usage_ax.bar([pos + width/2 for pos in x_positions], upload_usage, width, label='Upload', color=self.colors['upload'])

        self.usage_ax.set_xticks(x_positions)
        self.usage_ax.set_xticklabels([f"{hour:02d}:00" for hour in hours])

    def plot_weekly_usage(self):
        messagebox.showinfo("Info", "Weekly data usage graph will be implemented in a future update.")

    def plot_monthly_usage(self):
        messagebox.showinfo("Info", "Monthly data usage graph will be implemented in a future update.")

    def load_hourly_data(self):
        hourly_data_file = "hourly_usage.csv"
        if os.path.exists(hourly_data_file):
            try:
                with open(hourly_data_file, 'r', newline='') as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        hour = int(row['Hour'])
                        download_bytes = int(row['DownloadBytes'])
                        upload_bytes = int(row['UploadBytes'])
                        self.hourly_data[hour]['download'] = download_bytes
                        self.hourly_data[hour]['upload'] = upload_bytes
            except Exception as e:
                print(f"Error loading hourly data: {e}")

    def save_hourly_data(self):
        hourly_data_file = "hourly_usage.csv"
        try:
            with open(hourly_data_file, 'w', newline='') as csvfile:
                fieldnames = ['Hour', 'DownloadBytes', 'UploadBytes']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                for hour in range(24):
                    writer.writerow({
                        'Hour': hour,
                        'DownloadBytes': str(self.hourly_data[hour]['download']),
                        'UploadBytes': str(self.hourly_data[hour]['upload'])
                    })
        except Exception as e:
            print(f"Error saving hourly data: {e}")


if __name__ == "__main__":
    if sys.platform == 'win32':
        import ctypes
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    SmoothNetMonitor()