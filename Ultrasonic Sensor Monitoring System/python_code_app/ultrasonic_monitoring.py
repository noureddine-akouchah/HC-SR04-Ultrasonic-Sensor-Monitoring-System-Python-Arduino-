# Projet réalisé par Noreddine Akouchah

import serial
import serial.tools.list_ports
import customtkinter as ctk
from tkinter import messagebox, filedialog
import tkinter as tk
from tkinter import ttk
import threading
import time
import json
import csv
from datetime import datetime
from pathlib import Path
import logging
import re

# Set appearance mode and color theme
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class ModernArduinoInterface:
    def __init__(self):
        # Set UTF-8 encoding for the application
        import sys
        if sys.platform.startswith('win'):
            try:
                # Try to set console to UTF-8 on Windows
                import os
                os.system("chcp 65001 > nul")
            except:
                pass
        
        # Connection properties
        self.arduino = None
        self.is_running = False
        self.port = None
        self.baudrate = 9600
        
        # Statistics
        self.conforme_count = 0
        self.non_conforme_count = 0
        self.session_start_time = None
        self.test_history = []
        
        # Ultrasonic sensor data
        self.current_distance = 0.0
        self.distance_history = []
        self.max_distance_history = 100
        
        # Configuration
        self.config_file = Path("config.json")
        self.auto_reconnect = False
        self.sound_enabled = True
        
        # Setup logging
        self.setup_logging()
        
        # Load configuration
        self.load_config()
        
        # Setup GUI
        self.setup_gui()
        
        # Auto-connect if configured
        if hasattr(self, 'last_port') and self.last_port:
            self.port_var.set(self.last_port)

    def setup_logging(self):
        # Create a custom formatter that handles Unicode properly
        class UnicodeFormatter(logging.Formatter):
            def format(self, record):
                # Remove emoji characters for console logging to avoid encoding issues
                msg = super().format(record)
                # Remove emoji characters for file logging on Windows
                import re
                # Remove all emoji characters
                emoji_pattern = re.compile("["
                    "\U0001F600-\U0001F64F"  # emoticons
                    "\U0001F300-\U0001F5FF"  # symbols & pictographs
                    "\U0001F680-\U0001F6FF"  # transport & map symbols
                    "\U0001F1E0-\U0001F1FF"  # flags (iOS)
                    "\U00002702-\U000027B0"
                    "\U000024C2-\U0001F251"
                    "]+", flags=re.UNICODE)
                return emoji_pattern.sub('', msg)
        
        # Configure logging with UTF-8 support
        formatter = UnicodeFormatter('%(asctime)s - %(levelname)s - %(message)s')
        
        # File handler with UTF-8 encoding
        try:
            file_handler = logging.FileHandler('arduino_control.log', encoding='utf-8')
        except:
            # Fallback for systems that don't support UTF-8
            file_handler = logging.FileHandler('arduino_control.log')
        file_handler.setFormatter(formatter)
        
        # Console handler with UTF-8 encoding
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        
        # Configure logger
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)
        
        # Clear existing handlers
        self.logger.handlers.clear()
        
        # Add handlers
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
        
        # Prevent propagation to avoid duplicate logs
        self.logger.propagate = False

    def load_config(self):
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.last_port = config.get('last_port', '')
                    self.baudrate = config.get('baudrate', 9600)
                    self.auto_reconnect = config.get('auto_reconnect', False)
                    self.sound_enabled = config.get('sound_enabled', True)
                    
                    # Log without emojis
                    clean_msg = "Configuration loaded successfully"
                    self.logger.info(clean_msg)
        except Exception as e:
            clean_msg = f"Error loading config: {e}"
            self.logger.error(clean_msg)
            self.last_port = ''

    def save_config(self):
        try:
            config = {
                'last_port': self.port_var.get(),
                'baudrate': self.baudrate,
                'auto_reconnect': self.auto_reconnect,
                'sound_enabled': self.sound_enabled
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            
            # Log without emojis to avoid encoding issues
            clean_msg = "Configuration saved successfully"
            self.logger.info(clean_msg)
        except Exception as e:
            clean_msg = f"Error saving config: {e}"
            self.logger.error(clean_msg)

    def setup_gui(self):
        self.root = ctk.CTk()
        self.root.title(" Ultrasonic-monitoring System v1.0")
        self.root.geometry("1000x800")
        self.root.minsize(900, 700)
        
        # Set gradient background colors
        self.root.configure(fg_color=("#f0f0f0", "#0a0a0a"))

        # Main container with gradient effect
        self.main_container = ctk.CTkFrame(
            self.root,
            fg_color="transparent",
            corner_radius=0
        )
        self.main_container.pack(fill="both", expand=True)

        # Header with logo and title
        self.setup_header()

        # Create main content area
        self.content_area = ctk.CTkFrame(
            self.main_container,
            fg_color="transparent",
            corner_radius=0
        )
        self.content_area.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Create tabview with modern styling
        self.setup_tabview()

        # Setup tabs
        self.setup_main_tab()
        self.setup_stats_tab()
        self.setup_settings_tab()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.log_message("🎉 Interface moderne initialisée - Version 3.0")

    def setup_header(self):
        # Header frame with gradient
        header_frame = ctk.CTkFrame(
            self.main_container,
            height=120,
            fg_color=("#1e40af", "#1e3a8a"),
            corner_radius=0
        )
        header_frame.pack(fill="x", pady=(0, 20))
        header_frame.pack_propagate(False)

        # Logo and title container
        logo_title_frame = ctk.CTkFrame(
            header_frame,
            fg_color="transparent"
        )
        logo_title_frame.pack(side="left", fill="y", padx=30, pady=20)

        # Logo (using Unicode symbol as placeholder)
        logo_label = ctk.CTkLabel(
            logo_title_frame,
            text="⚡",
            font=ctk.CTkFont(size=60),
            text_color="#fbbf24"
        )
        logo_label.pack(side="left", padx=(0, 20))

        # Title and subtitle
        title_container = ctk.CTkFrame(
            logo_title_frame,
            fg_color="transparent"
        )
        title_container.pack(side="left", fill="y")

        main_title = ctk.CTkLabel(
            title_container,
            text="ultrasonic monitoring System",
            font=ctk.CTkFont(size=32, weight="bold"),
            text_color="white"
        )
        main_title.pack(anchor="w", pady=(5, 0))

        subtitle = ctk.CTkLabel(
            title_container,
            text="🔬 Compliance Control System v3.0",
            font=ctk.CTkFont(size=16),
            text_color="#e5e7eb"
        )
        subtitle.pack(anchor="w", pady=(0, 5))

        # Status panel on the right
        status_panel = ctk.CTkFrame(
            header_frame,
            fg_color=("#ffffff", "#1f2937"),
            corner_radius=15
        )
        status_panel.pack(side="right", padx=30, pady=20, fill="y")

        # Connection status
        self.connection_indicator = ctk.CTkLabel(
            status_panel,
            text="●",
            font=ctk.CTkFont(size=30),
            text_color="#ef4444"
        )
        self.connection_indicator.pack(pady=(15, 5))

        self.connection_status_label = ctk.CTkLabel(
            status_panel,
            text="DISCONNECTED",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#ef4444"
        )
        self.connection_status_label.pack(pady=(0, 10))

        # Live time display
        self.time_label = ctk.CTkLabel(
            status_panel,
            text="",
            font=ctk.CTkFont(size=11),
            text_color=("#374151", "#9ca3af")
        )
        self.time_label.pack(pady=(0, 15))
        
        # Start time update
        self.update_time()

    def update_time(self):
        current_time = datetime.now().strftime("%H:%M:%S")
        self.time_label.configure(text=f"🕐 {current_time}")
        self.root.after(1000, self.update_time)

    def setup_tabview(self):
        # Custom tabview with modern styling
        self.tabview = ctk.CTkTabview(
            self.content_area,
            width=950,
            height=600,
            fg_color=("#ffffff", "#1a1a1a"),
            segmented_button_fg_color=("#e5e7eb", "#374151"),
            segmented_button_selected_color=("#3b82f6", "#1d4ed8"),
            segmented_button_selected_hover_color=("#2563eb", "#1e40af"),
            text_color=("#1f2937", "#f9fafb"),
            text_color_disabled=("#9ca3af", "#6b7280"),
            corner_radius=15
        )
        self.tabview.pack(fill="both", expand=True)

        # Add tabs with icons
        self.tab_main = self.tabview.add("🎛️ Contrôle Principal")
        self.tab_stats = self.tabview.add("📊 Statistiques")
        self.tab_settings = self.tabview.add("⚙️ Paramètres")

    def setup_main_tab(self):
        # Main tab with scroll
        main_scroll = ctk.CTkScrollableFrame(
            self.tab_main,
            fg_color="transparent"
        )
        main_scroll.pack(fill="both", expand=True, padx=15, pady=15)

        # Connection section
        self.setup_connection_section(main_scroll)

        # Distance monitoring section
        self.setup_distance_section(main_scroll)

        # Test controls section
        self.setup_test_controls_section(main_scroll)

        # Stats cards section
        self.setup_stats_cards(main_scroll)

        # Log section
        self.setup_log_section(main_scroll)

    def setup_connection_section(self, parent):
        # Connection card with glass effect
        conn_card = ctk.CTkFrame(
            parent,
            fg_color=("#f8fafc", "#111827"),
            corner_radius=20,
            border_width=1,
            border_color=("#e5e7eb", "#374151")
        )
        conn_card.pack(fill="x", pady=(0, 25))

        # Card header
        header = ctk.CTkFrame(
            conn_card,
            fg_color=("#3b82f6", "#1e40af"),
            corner_radius=15,
            height=60
        )
        header.pack(fill="x", padx=15, pady=15)
        header.pack_propagate(False)

        title_label = ctk.CTkLabel(
            header,
            text="🔌 Connection Configuration",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=15)

        # Connection controls grid
        controls_grid = ctk.CTkFrame(conn_card, fg_color="transparent")
        controls_grid.pack(fill="x", padx=20, pady=(0, 20))

        # Port selection row
        port_row = ctk.CTkFrame(controls_grid, fg_color="transparent")
        port_row.pack(fill="x", pady=10)

        ctk.CTkLabel(
            port_row,
            text="📡 COM Port:",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(side="left", padx=(0, 15))
        
        self.port_var = ctk.StringVar()
        self.port_combo = ctk.CTkComboBox(
            port_row,
            variable=self.port_var,
            width=250,
            state="readonly",
            border_color=("#3b82f6", "#1e40af"),
            button_color=("#3b82f6", "#1e40af"),
            button_hover_color=("#2563eb", "#1d4ed8")
        )
        self.port_combo.pack(side="left", padx=5)

        refresh_btn = ctk.CTkButton(
            port_row,
            text="🔄 Refresh",
            command=self.refresh_ports,
            width=120,
            height=35,
            fg_color=("#10b981", "#059669"),
            hover_color=("#047857", "#065f46"),
            font=ctk.CTkFont(weight="bold")
        )
        refresh_btn.pack(side="left", padx=15)

        # Baudrate and connect row
        connect_row = ctk.CTkFrame(controls_grid, fg_color="transparent")
        connect_row.pack(fill="x", pady=10)

        ctk.CTkLabel(
            connect_row,
            text="⚡ Baudrate:",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(side="left", padx=(0, 15))
        
        self.baudrate_var = ctk.StringVar(value=str(self.baudrate))
        baudrate_combo = ctk.CTkComboBox(
            connect_row,
            variable=self.baudrate_var,
            values=["9600", "19200", "38400", "57600", "115200"],
            width=130,
            state="readonly",
            border_color=("#3b82f6", "#1e40af")
        )
        baudrate_combo.pack(side="left", padx=5)

        self.connect_btn = ctk.CTkButton(
            connect_row,
            text="🚀 CONNECT",
            command=self.toggle_connection,
            width=200,
            height=45,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=("#3b82f6", "#1e40af"),
            hover_color=("#2563eb", "#1d4ed8")
        )
        self.connect_btn.pack(side="right", padx=15)

        # Auto-reconnect option
        self.auto_reconnect_var = ctk.BooleanVar(value=self.auto_reconnect)
        auto_reconnect_cb = ctk.CTkCheckBox(
            connect_row,
            text="🔄 Auto Reconnect",
            variable=self.auto_reconnect_var,
            command=self.toggle_auto_reconnect,
            font=ctk.CTkFont(size=12),
            checkbox_width=20,
            checkbox_height=20
        )
        auto_reconnect_cb.pack(side="right", padx=(0, 30))

        self.refresh_ports()

    def setup_distance_section(self, parent):
        # Distance monitoring card
        distance_card = ctk.CTkFrame(
            parent,
            fg_color=("#f8fafc", "#111827"),
            corner_radius=20,
            border_width=1,
            border_color=("#e5e7eb", "#374151")
        )
        distance_card.pack(fill="x", pady=(0, 25))

        # Card header
        header = ctk.CTkFrame(
            distance_card,
            fg_color=("#8b5cf6", "#7c3aed"),
            corner_radius=15,
            height=60
        )
        header.pack(fill="x", padx=15, pady=15)
        header.pack_propagate(False)

        title_label = ctk.CTkLabel(
            header,
            text="📏 Ultrasonic Sensor",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=15)

        # Threshold configuration
        threshold_frame = ctk.CTkFrame(
            distance_card,
            fg_color=("#f1f5f9", "#1e293b"),
            corner_radius=10
        )
        threshold_frame.pack(fill="x", padx=20, pady=(0, 15))

        ctk.CTkLabel(
            threshold_frame,
            text="⚙️ Threshold Settings:",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=(15, 10))

        threshold_controls = ctk.CTkFrame(threshold_frame, fg_color="transparent")
        threshold_controls.pack(pady=(0, 15))

        # Min threshold
        ctk.CTkLabel(threshold_controls, text="Min:", font=ctk.CTkFont(size=14)).pack(side="left", padx=(20, 5))
        self.min_entry = ctk.CTkEntry(
            threshold_controls,
            width=80,
            placeholder_text="10",
            border_color=("#8b5cf6", "#7c3aed"),
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.min_entry.pack(side="left", padx=5)
        self.min_entry.insert(0, "10")

        # Max threshold
        ctk.CTkLabel(threshold_controls, text="Max:", font=ctk.CTkFont(size=14)).pack(side="left", padx=(15, 5))
        self.max_entry = ctk.CTkEntry(
            threshold_controls,
            width=80,
            placeholder_text="50",
            border_color=("#8b5cf6", "#7c3aed"),
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.max_entry.pack(side="left", padx=5)
        self.max_entry.insert(0, "50")

        ctk.CTkLabel(threshold_controls, text="cm", font=ctk.CTkFont(size=14)).pack(side="left", padx=(5, 15))

        apply_btn = ctk.CTkButton(
            threshold_controls,
            text="✅ Apply",
            command=self.apply_thresholds,
            width=120,
            height=35,
            fg_color=("#10b981", "#059669"),
            hover_color=("#047857", "#065f46"),
            font=ctk.CTkFont(weight="bold")
        )
        apply_btn.pack(side="left", padx=15)

        # Distance display with neon effect
        display_frame = ctk.CTkFrame(
            distance_card,
            fg_color=("#1e293b", "#0f172a"),
            corner_radius=15
        )
        display_frame.pack(fill="x", padx=20, pady=(0, 15))

        self.distance_label = ctk.CTkLabel(
            display_frame,
            text="-- cm",
            font=ctk.CTkFont(size=48, weight="bold"),
            text_color="#06b6d4"
        )
        self.distance_label.pack(pady=25)

        self.conformity_label = ctk.CTkLabel(
            display_frame,
            text="",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        self.conformity_label.pack(pady=(0, 20))

        # Distance statistics row
        stats_row = ctk.CTkFrame(distance_card, fg_color="transparent")
        stats_row.pack(fill="x", padx=20, pady=(0, 20))

        # Stats cards
        for i, (label, color) in enumerate([("Min", "#ef4444"), ("Max", "#f59e0b"), ("Avg", "#06b6d4")]):
            stat_card = ctk.CTkFrame(
                stats_row,
                fg_color=("#ffffff", "#1f2937"),
                corner_radius=10,
                border_width=2,
                border_color=(color, color)
            )
            stat_card.pack(side="left", fill="x", expand=True, padx=5)

            if i == 0:
                self.min_distance_label = ctk.CTkLabel(
                    stat_card,
                    text="Min: -- cm",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=color
                )
                self.min_distance_label.pack(pady=10)
            elif i == 1:
                self.max_distance_label = ctk.CTkLabel(
                    stat_card,
                    text="Max: -- cm",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=color
                )
                self.max_distance_label.pack(pady=10)
            else:
                self.avg_distance_label = ctk.CTkLabel(
                    stat_card,
                    text="Avg: -- cm",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=color
                )
                self.avg_distance_label.pack(pady=10)

        # Reset button
        reset_distance_btn = ctk.CTkButton(
            stats_row,
            text="🔄 Reset",
            command=self.reset_distance_stats,
            width=100,
            height=35,
            fg_color=("#ef4444", "#dc2626"),
            hover_color=("#dc2626", "#b91c1c"),
            font=ctk.CTkFont(weight="bold")
        )
        reset_distance_btn.pack(side="right", padx=15)

        # Initialize thresholds
        self.min_threshold = 10.0
        self.max_threshold = 50.0

    def setup_test_controls_section(self, parent):
        # Test controls card
        test_card = ctk.CTkFrame(
            parent,
            fg_color=("#f8fafc", "#111827"),
            corner_radius=20,
            border_width=1,
            border_color=("#e5e7eb", "#374151")
        )
        test_card.pack(fill="x", pady=(0, 25))

        # Card header
        header = ctk.CTkFrame(
            test_card,
            fg_color=("#f59e0b", "#d97706"),
            corner_radius=15,
            height=60
        )
        header.pack(fill="x", padx=15, pady=15)
        header.pack_propagate(False)

        title_label = ctk.CTkLabel(
            header,
            text="🧪 Test Controls",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=15)

        # Status display
        status_frame = ctk.CTkFrame(
            test_card,
            fg_color=("#1e293b", "#0f172a"),
            corner_radius=10
        )
        status_frame.pack(fill="x", padx=20, pady=(0, 20))

        self.status_label = ctk.CTkLabel(
            status_frame,
            text="Waiting for connection...",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#94a3b8"
        )
        self.status_label.pack(pady=20)

        # LED indicators with glow effect
        led_container = ctk.CTkFrame(test_card, fg_color="transparent")
        led_container.pack(pady=(0, 20))

        self.led_conforme = ctk.CTkButton(
            led_container,
            text="✅ PASS",
            width=180,
            height=70,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=("#374151", "#374151"),
            hover_color=("#374151", "#374151"),
            state="disabled",
            corner_radius=15
        )
        self.led_conforme.pack(side="left", padx=15)

        self.led_non_conforme = ctk.CTkButton(
            led_container,
            text="❌ FAIL",
            width=180,
            height=70,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=("#374151", "#374151"),
            hover_color=("#374151", "#374151"),
            state="disabled",
            corner_radius=15
        )
        self.led_non_conforme.pack(side="left", padx=15)

        # Manual test buttons with hover effects
        manual_frame = ctk.CTkFrame(test_card, fg_color="transparent")
        manual_frame.pack(pady=(0, 20))

        conforme_btn = ctk.CTkButton(
            manual_frame,
            text="✅ Pass Test",
            command=lambda: self.manual_test(True),
            width=200,
            height=55,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=("#10b981", "#059669"),
            hover_color=("#047857", "#065f46"),
            corner_radius=15
        )
        conforme_btn.pack(side="left", padx=20)

        non_conforme_btn = ctk.CTkButton(
            manual_frame,
            text="❌ Fail Test",
            command=lambda: self.manual_test(False),
            width=200,
            height=55,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=("#ef4444", "#dc2626"),
            hover_color=("#dc2626", "#b91c1c"),
            corner_radius=15
        )
        non_conforme_btn.pack(side="left", padx=20)

    def setup_stats_cards(self, parent):
        # Quick stats cards
        stats_container = ctk.CTkFrame(
            parent,
            fg_color=("#f8fafc", "#111827"),
            corner_radius=20,
            border_width=1,
            border_color=("#e5e7eb", "#374151")
        )
        stats_container.pack(fill="x", pady=(0, 25))

        # Header
        header = ctk.CTkFrame(
            stats_container,
            fg_color=("#06b6d4", "#0891b2"),
            corner_radius=15,
            height=60
        )
        header.pack(fill="x", padx=15, pady=15)
        header.pack_propagate(False)

        title_label = ctk.CTkLabel(
            header,
            text="📊 Real-Time Statistics",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=15)

        # Stats grid
        stats_grid = ctk.CTkFrame(stats_container, fg_color="transparent")
        stats_grid.pack(fill="x", padx=20, pady=(0, 20))

        # Create stat cards
        stat_configs = [
            ("Passes", "#10b981", "self.stats_conforme_label"),
            ("Fails", "#ef4444", "self.stats_non_conforme_label"),
            ("Total Tests", "#8b5cf6", "self.stats_total_label"),
            ("Success Rate", "#06b6d4", "self.success_rate_label")
        ]

        for i, (title, color, attr_name) in enumerate(stat_configs):
            stat_card = ctk.CTkFrame(
                stats_grid,
                fg_color=("#ffffff", "#1f2937"),
                corner_radius=15,
                border_width=2,
                border_color=(color, color)
            )
            stat_card.grid(row=0, column=i, padx=10, pady=10, sticky="ew")
            stats_grid.grid_columnconfigure(i, weight=1)

            # Card icon and title
            ctk.CTkLabel(
                stat_card,
                text=title,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=("#64748b", "#94a3b8")
            ).pack(pady=(15, 5))

            # Value label
            label = ctk.CTkLabel(
                stat_card,
                text="0" if "Taux" not in title else "0%",
                font=ctk.CTkFont(size=24, weight="bold"),
                text_color=color
            )
            label.pack(pady=(0, 15))
            
            # Set the attribute dynamically
            if attr_name == "self.stats_conforme_label":
                self.stats_conforme_label = label
            elif attr_name == "self.stats_non_conforme_label":
                self.stats_non_conforme_label = label
            elif attr_name == "self.stats_total_label":
                self.stats_total_label = label
            elif attr_name == "self.success_rate_label":
                self.success_rate_label = label

        # Reset button
        reset_stats_btn = ctk.CTkButton(
            stats_grid,
            text="🔄 Reset",
            command=self.reset_stats,
            width=150,
            height=40,
            fg_color=("#ef4444", "#dc2626"),
            hover_color=("#dc2626", "#b91c1c"),
            font=ctk.CTkFont(weight="bold"),
            corner_radius=10
        )
        reset_stats_btn.grid(row=0, column=4, padx=20, pady=10)

    def setup_log_section(self, parent):
        # Log section card
        log_card = ctk.CTkFrame(
            parent,
            fg_color=("#f8fafc", "#111827"),
            corner_radius=20,
            border_width=1,
            border_color=("#e5e7eb", "#374151")
        )
        log_card.pack(fill="both", expand=True)

        # Header
        header = ctk.CTkFrame(
            log_card,
            fg_color=("#64748b", "#475569"),
            corner_radius=15,
            height=60
        )
        header.pack(fill="x", padx=15, pady=15)
        header.pack_propagate(False)

        title_label = ctk.CTkLabel(
            header,
            text="📋 Event Log",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=15)

        # Log controls
        controls_frame = ctk.CTkFrame(log_card, fg_color="transparent")
        controls_frame.pack(fill="x", padx=20, pady=(0, 15))

        clear_btn = ctk.CTkButton(
            controls_frame,
            text="🗑️ Clear",
            command=self.clear_log,
            width=100,
            height=35,
            fg_color=("#ef4444", "#dc2626"),
            hover_color=("#dc2626", "#b91c1c"),
            font=ctk.CTkFont(weight="bold")
        )
        clear_btn.pack(side="left", padx=5)

        save_btn = ctk.CTkButton(
            controls_frame,
            text="💾 Save",
            command=self.save_log,
            width=120,
            height=35,
            fg_color=("#10b981", "#059669"),
            hover_color=("#047857", "#065f46"),
            font=ctk.CTkFont(weight="bold")
        )
        save_btn.pack(side="left", padx=5)

        self.auto_scroll_var = ctk.BooleanVar(value=True)
        auto_scroll_cb = ctk.CTkCheckBox(
            controls_frame,
            text="📜 Auto Scroll",
            variable=self.auto_scroll_var,
            font=ctk.CTkFont(size=12),
            checkbox_width=20,
            checkbox_height=20
        )
        auto_scroll_cb.pack(side="right", padx=15)

        # Log text area with modern styling
        self.log_text = ctk.CTkTextbox(
            log_card,
            height=250,
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=("#1e293b", "#0f172a"),
            text_color=("#e2e8f0", "#cbd5e1"),
            corner_radius=10,
            border_width=1,
            border_color=("#334155", "#1e293b")
        )
        self.log_text.pack(fill="both", expand=True, padx=20, pady=(0, 20))

    def setup_stats_tab(self):
        # Stats tab with advanced tables
        stats_scroll = ctk.CTkScrollableFrame(
            self.tab_stats,
            fg_color="transparent"
        )
        stats_scroll.pack(fill="both", expand=True, padx=20, pady=20)

        # Session info card
        session_card = ctk.CTkFrame(
            stats_scroll,
            fg_color=("#f8fafc", "#111827"),
            corner_radius=20,
            border_width=1,
            border_color=("#e5e7eb", "#374151")
        )
        session_card.pack(fill="x", pady=(0, 25))

        # Header
        header = ctk.CTkFrame(
            session_card,
            fg_color=("#3b82f6", "#1e40af"),
            corner_radius=15,
            height=60
        )
        header.pack(fill="x", padx=15, pady=15)
        header.pack_propagate(False)

        title_label = ctk.CTkLabel(
            header,
            text="📈 Session Information",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=15)

        # Session details
        session_details = ctk.CTkFrame(session_card, fg_color="transparent")
        session_details.pack(fill="x", padx=20, pady=(0, 20))

        self.session_start_label = ctk.CTkLabel(
            session_details,
            text="📅 Session not started",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=("#64748b", "#94a3b8")
        )
        self.session_start_label.pack(anchor="w", pady=5)

        self.session_duration_label = ctk.CTkLabel(
            session_details,
            text="⏱️ Durée: --",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=("#64748b", "#94a3b8")
        )
        self.session_duration_label.pack(anchor="w", pady=5)

        # Test history table
        history_card = ctk.CTkFrame(
            stats_scroll,
            fg_color=("#f8fafc", "#111827"),
            corner_radius=20,
            border_width=1,
            border_color=("#e5e7eb", "#374151")
        )
        history_card.pack(fill="both", expand=True)

        # Table header
        table_header = ctk.CTkFrame(
            history_card,
            fg_color=("#8b5cf6", "#7c3aed"),
            corner_radius=15,
            height=60
        )
        table_header.pack(fill="x", padx=15, pady=15)
        table_header.pack_propagate(False)

        title_label = ctk.CTkLabel(
            table_header,
            text="📋 Detailed Test History",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=15)

        # Create modern table using tkinter Treeview with custom styling
        table_container = ctk.CTkFrame(
            history_card,
            fg_color=("#ffffff", "#1f2937"),
            corner_radius=10
        )
        table_container.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Configure treeview style for dark theme
        style = ttk.Style()
        style.theme_use("clam")
        
        # Configure colors for dark theme
        style.configure("Custom.Treeview",
                       background="#1f2937",
                       foreground="#e5e7eb",
                       fieldbackground="#1f2937",
                       borderwidth=0,
                       relief="flat")
        
        style.configure("Custom.Treeview.Heading",
                       background="#374151",
                       foreground="#f9fafb",
                       borderwidth=1,
                       relief="solid",
                       font=('Segoe UI', 10, 'bold'))
        
        style.map("Custom.Treeview",
                 background=[('selected', '#3b82f6')],
                 foreground=[('selected', '#ffffff')])

        # Create treeview
        columns = ('🕐 Time', '✅ Result', '📏 Distance', '⏱️ Duration')
        self.history_tree = ttk.Treeview(
            table_container,
            columns=columns,
            show='headings',
            height=12,
            style="Custom.Treeview"
        )
        
        # Configure column headings and widths
        self.history_tree.heading('🕐 Time', text='🕐 Time', anchor='center')
        self.history_tree.heading('✅ Result', text='✅ Result', anchor='center')
        self.history_tree.heading('📏 Distance', text='📏 Distance (cm)', anchor='center')
        self.history_tree.heading('⏱️ Duration', text='⏱️ Duration (ms)', anchor='center')
        
        self.history_tree.column('🕐 Time', width=120, anchor='center')
        self.history_tree.column('✅ Result', width=150, anchor='center')
        self.history_tree.column('📏 Distance', width=120, anchor='center')
        self.history_tree.column('⏱️ Duration', width=120, anchor='center')

        # Add scrollbar
        tree_scroll = ctk.CTkScrollbar(
            table_container,
            orientation="vertical",
            command=self.history_tree.yview
        )
        self.history_tree.configure(yscrollcommand=tree_scroll.set)

        # Pack treeview and scrollbar
        self.history_tree.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        tree_scroll.pack(side="right", fill="y", pady=10, padx=(0, 10))

    def setup_settings_tab(self):
        # Settings tab with modern cards
        settings_scroll = ctk.CTkScrollableFrame(
            self.tab_settings,
            fg_color="transparent"
        )
        settings_scroll.pack(fill="both", expand=True, padx=20, pady=20)

        # Appearance settings card
        appearance_card = ctk.CTkFrame(
            settings_scroll,
            fg_color=("#f8fafc", "#111827"),
            corner_radius=20,
            border_width=1,
            border_color=("#e5e7eb", "#374151")
        )
        appearance_card.pack(fill="x", pady=(0, 25))

        # Header
        header = ctk.CTkFrame(
            appearance_card,
            fg_color=("#8b5cf6", "#7c3aed"),
            corner_radius=15,
            height=60
        )
        header.pack(fill="x", padx=15, pady=15)
        header.pack_propagate(False)

        title_label = ctk.CTkLabel(
            header,
            text="🎨 Apparence et Thème",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=15)

        # Theme controls
        theme_frame = ctk.CTkFrame(appearance_card, fg_color="transparent")
        theme_frame.pack(fill="x", padx=20, pady=(0, 20))

        ctk.CTkLabel(
            theme_frame,
            text="🌙 Mode d'apparence:",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(side="left", padx=(0, 15))
        
        self.appearance_var = ctk.StringVar(value="Dark")
        appearance_combo = ctk.CTkComboBox(
            theme_frame,
            variable=self.appearance_var,
            values=["Light", "Dark", "System"],
            command=self.change_appearance,
            width=150,
            border_color=("#8b5cf6", "#7c3aed")
        )
        appearance_combo.pack(side="left", padx=5)

        # Color theme selector
        ctk.CTkLabel(
            theme_frame,
            text="🎨 Thème de couleur:",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(side="left", padx=(30, 15))
        
        self.color_theme_var = ctk.StringVar(value="blue")
        color_combo = ctk.CTkComboBox(
            theme_frame,
            variable=self.color_theme_var,
            values=["blue", "green", "dark-blue"],
            command=self.change_color_theme,
            width=150,
            border_color=("#8b5cf6", "#7c3aed")
        )
        color_combo.pack(side="left", padx=5)

        # Audio settings card
        audio_card = ctk.CTkFrame(
            settings_scroll,
            fg_color=("#f8fafc", "#111827"),
            corner_radius=20,
            border_width=1,
            border_color=("#e5e7eb", "#374151")
        )
        audio_card.pack(fill="x", pady=(0, 25))

        # Header
        header = ctk.CTkFrame(
            audio_card,
            fg_color=("#f59e0b", "#d97706"),
            corner_radius=15,
            height=60
        )
        header.pack(fill="x", padx=15, pady=15)
        header.pack_propagate(False)

        title_label = ctk.CTkLabel(
            header,
            text="🔊 Paramètres Audio",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=15)

        # Audio controls
        audio_controls = ctk.CTkFrame(audio_card, fg_color="transparent")
        audio_controls.pack(fill="x", padx=20, pady=(0, 20))

        self.sound_var = ctk.BooleanVar(value=self.sound_enabled)
        sound_cb = ctk.CTkCheckBox(
            audio_controls,
            text="🔔 Activer les notifications sonores",
            variable=self.sound_var,
            command=self.toggle_sound,
            font=ctk.CTkFont(size=14, weight="bold"),
            checkbox_width=25,
            checkbox_height=25
        )
        sound_cb.pack(anchor="w", padx=20, pady=15)

        # Export settings card
        export_card = ctk.CTkFrame(
            settings_scroll,
            fg_color=("#f8fafc", "#111827"),
            corner_radius=20,
            border_width=1,
            border_color=("#e5e7eb", "#374151")
        )
        export_card.pack(fill="x", pady=(0, 25))

        # Header
        header = ctk.CTkFrame(
            export_card,
            fg_color=("#10b981", "#059669"),
            corner_radius=15,
            height=60
        )
        header.pack(fill="x", padx=15, pady=15)
        header.pack_propagate(False)

        title_label = ctk.CTkLabel(
            header,
            text="📤 Data Export",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=15)

        # Export buttons
        export_frame = ctk.CTkFrame(export_card, fg_color="transparent")
        export_frame.pack(fill="x", padx=20, pady=(0, 20))

        csv_btn = ctk.CTkButton(
            export_frame,
            text="📊 Export CSV",
            command=self.export_to_csv,
            width=180,
            height=50,
            fg_color=("#10b981", "#059669"),
            hover_color=("#047857", "#065f46"),
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=15
        )
        csv_btn.pack(side="left", padx=15)

        json_btn = ctk.CTkButton(
            export_frame,
            text="📋 Export JSON",
            command=self.export_to_json,
            width=180,
            height=50,
            fg_color=("#3b82f6", "#1e40af"),
            hover_color=("#2563eb", "#1d4ed8"),
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=15
        )
        json_btn.pack(side="left", padx=15)

        # Advanced settings card
        advanced_card = ctk.CTkFrame(
            settings_scroll,
            fg_color=("#f8fafc", "#111827"),
            corner_radius=20,
            border_width=1,
            border_color=("#e5e7eb", "#374151")
        )
        advanced_card.pack(fill="x")

        # Header
        header = ctk.CTkFrame(
            advanced_card,
            fg_color=("#ef4444", "#dc2626"),
            corner_radius=15,
            height=60
        )
        header.pack(fill="x", padx=15, pady=15)
        header.pack_propagate(False)

        title_label = ctk.CTkLabel(
            header,
            text="⚙️ Paramètres Avancés",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=15)

        # Advanced controls
        advanced_frame = ctk.CTkFrame(advanced_card, fg_color="transparent")
        advanced_frame.pack(fill="x", padx=20, pady=(0, 20))

        reset_btn = ctk.CTkButton(
            advanced_frame,
            text="🔄 Reset Configuration",
            command=self.reset_config,
            width=220,
            height=45,
            fg_color=("#ef4444", "#dc2626"),
            hover_color=("#dc2626", "#b91c1c"),
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=15
        )
        reset_btn.pack(side="left", padx=15)

        log_folder_btn = ctk.CTkButton(
            advanced_frame,
            text="📁 Open Logs Folder",
            command=self.open_log_folder,
            width=200,
            height=45,
            fg_color=("#64748b", "#475569"),
            hover_color=("#475569", "#334155"),
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=15
        )
        log_folder_btn.pack(side="left", padx=15)

    def change_appearance(self, new_appearance):
        ctk.set_appearance_mode(new_appearance)

    def change_color_theme(self, new_theme):
        ctk.set_default_color_theme(new_theme)
        messagebox.showinfo("Theme Change", "Restart the application to apply the new color theme.")

    # Connection methods (keeping the original logic with modern UI updates)
    def refresh_ports(self):
        ports = serial.tools.list_ports.comports()
        port_list = [f"{port.device} - {port.description}" for port in ports]
        port_devices = [port.device for port in ports]
        
        self.port_combo.configure(values=port_list)
        self.port_devices = port_devices
        
        if port_devices:
            current_selection = self.port_var.get().split(' - ')[0] if ' - ' in self.port_var.get() else self.port_var.get()
            if current_selection in port_devices:
                idx = port_devices.index(current_selection)
                self.port_var.set(port_list[idx])
            else:
                self.port_var.set(port_list[0])
        else:
            self.port_var.set("")
            self.log_message("⚠️ No serial port detected")

    def toggle_connection(self):
        if self.is_running:
            self.disconnect()
        else:
            self.connect()

    def connect(self):
        selected_port_display = self.port_var.get()
        if not selected_port_display:
            messagebox.showwarning("Missing Port", "Please select a COM port.")
            return

        selected_port = selected_port_display.split(' - ')[0]
        self.baudrate = int(self.baudrate_var.get())

        try:
            self.arduino = serial.Serial(selected_port, self.baudrate, timeout=1)
            time.sleep(2)
            
            self.is_running = True
            self.port = selected_port
            self.session_start_time = datetime.now()
            
            self.update_connection_ui(True)
            self.start_reading_thread()
            self.start_connection_timer()
            
            self.log_message("🚀 Connected to {} at {} baud".format(selected_port, self.baudrate))
            self.save_config()
            
        except serial.SerialException as e:
            messagebox.showerror("Connection Error", 
                               "Unable to connect to port {}\n\nError: {}".format(selected_port, str(e)))
            self.update_connection_ui(False)
            self.log_message("❌ Connection failed on {}: {}".format(selected_port, str(e)))

    def disconnect(self):
        self.is_running = False
        if self.arduino and self.arduino.is_open:
            try:
                self.arduino.close()
            except Exception as e:
                self.logger.error(f"Error closing serial connection: {e}")
        
        self.update_connection_ui(False)
        self.reset_status()
        self.log_message("🔌 Disconnected")

    def update_connection_ui(self, connected):
        if connected:
            self.connection_status_label.configure(text="CONNECTED", text_color="#10b981")
            self.connection_indicator.configure(text_color="#10b981")
            self.connect_btn.configure(
                text="🔌 DISCONNECT",
                fg_color=("#ef4444", "#dc2626"),
                hover_color=("#dc2626", "#b91c1c")
            )
            self.port_combo.configure(state="disabled")
            self.status_label.configure(text="✅ Waiting for data...", text_color="#10b981")
        else:
            self.connection_status_label.configure(text="DISCONNECTED", text_color="#ef4444")
            self.connection_indicator.configure(text_color="#ef4444")
            self.connect_btn.configure(
                text="🚀 CONNECT",
                fg_color=("#3b82f6", "#1e40af"),
                hover_color=("#2563eb", "#1d4ed8")
            )
            self.port_combo.configure(state="readonly")
            self.status_label.configure(text="⏳ Waiting for connection...", text_color="#94a3b8")

    def start_connection_timer(self):
        def update_timer():
            if self.is_running and self.session_start_time:
                duration = datetime.now() - self.session_start_time
                hours, remainder = divmod(int(duration.total_seconds()), 3600)
                minutes, seconds = divmod(remainder, 60)
                time_str = f"⏱️ Duration: {hours:02d}:{minutes:02d}:{seconds:02d}"
                self.session_duration_label.configure(text=time_str)
                self.session_start_label.configure(
                    text=f"📅 Started: {self.session_start_time.strftime('%H:%M:%S')}"
                )
                self.root.after(1000, update_timer)
        
        self.root.after(1000, update_timer)

    def start_reading_thread(self):
        threading.Thread(target=self.read_arduino, daemon=True).start()

    def read_arduino(self):
        errors = 0
        max_errors = 5
        
        while self.is_running and errors < max_errors:
            try:
                if self.arduino and self.arduino.in_waiting > 0:
                    data = self.arduino.readline().decode('utf-8', errors='ignore').strip()
                    
                    if data:
                        self.process_arduino_data(data)
                        errors = 0
                        
                time.sleep(0.05)
                

                    
            except Exception as e:
                errors += 1
                clean_msg = "Unexpected error reading Arduino data: {}".format(str(e))
                self.logger.error(clean_msg)
                if errors >= max_errors:
                    self.log_message("❌ Multiple errors: {}".format(str(e)))
                    self.root.after(0, self.disconnect)
                    break

    def process_arduino_data(self, data):
        try:
            # Process distance data first
            if data.startswith("DIST:"):
                distance = float(data.replace("DIST:", "").strip())
                self.process_distance_data(distance)
                return
            
            # Check for pure numeric distance values
            if data.replace('.', '').replace('-', '').isdigit():
                distance = float(data)
                self.process_distance_data(distance)
                return
            
            # Process test results
            if data.upper() == "OK":
                self.root.after(0, lambda: self.process_result(conforme=True))
                return
            elif data.upper() == "NON":
                self.root.after(0, lambda: self.process_result(conforme=False))
                return
            
            # Try to parse distance patterns
            if self.parse_arduino_data(data):
                return
            
            # Log other messages
            if data:
                self.log_message("📡 Arduino: {}".format(data))
                
        except Exception as e:
            clean_msg = "Error processing Arduino data '{}': {}".format(data, str(e))
            self.logger.error(clean_msg)
            self.log_message("⚠️ Processing error: {}".format(data))

    def parse_arduino_data(self, data):
        try:
            if data.startswith("DIST:"):
                distance = float(data.replace("DIST:", "").strip())
                self.update_distance(distance)
                return True
            
            elif data.replace('.', '').replace('-', '').isdigit():
                distance = float(data)
                if 0 <= distance <= 400:
                    self.update_distance(distance)
                    return True
            
            distance_patterns = [
                r'Distance:\s*(\d+\.?\d*)\s*cm',
                r'(\d+\.?\d*)\s*cm',

            ]
            
            for pattern in distance_patterns:
                match = re.search(pattern, data, re.IGNORECASE)
                if match:
                    distance = float(match.group(1))
                    if 0 <= distance <= 400:
                        self.update_distance(distance)
                        return True
            
            return False
        except (ValueError, AttributeError) as e:
            self.logger.error(f"Error parsing distance data '{data}': {e}")
            return False

    def process_distance_data(self, distance):
        if not (0 <= distance <= 400):
            self.log_message("⚠️ Distance out of range: {:.1f}cm".format(distance))
            return
            
        self.current_distance = distance
        
        self.distance_history.append(distance)
        if len(self.distance_history) > self.max_distance_history:
            self.distance_history.pop(0)
        
        self.root.after(0, self.update_distance_display)
        
        self.log_message("📏 Measurement: {:.1f} cm".format(distance))

    def update_distance(self, distance):
        if not (0 <= distance <= 400):
            self.log_message("⚠️ Suspicious distance: {:.1f}cm".format(distance))
            return
            
        self.current_distance = distance
        
        self.distance_history.append(distance)
        if len(self.distance_history) > self.max_distance_history:
            self.distance_history.pop(0)
        
        self.root.after(0, self.update_distance_display)
        
        self.log_message("📏 Distance: {:.1f} cm".format(distance))

    def apply_thresholds(self):
        """Apply the min/max threshold values"""
        try:
            min_val = float(self.min_entry.get())
            max_val = float(self.max_entry.get())
            
            if min_val >= max_val:
                messagebox.showerror("Error", "Minimum value must be less than maximum value")
                return
                
            if min_val < 0 or max_val < 0:
                messagebox.showerror("Error", "Values must be positive")
                return
                
            self.min_threshold = min_val
            self.max_threshold = max_val
            
            self.log_message("⚙️ Seuils appliqués: Min={:.1f}cm, Max={:.1f}cm".format(min_val, max_val))
            
            # Re-evaluate current distance if available
            if self.current_distance > 0:
                self.check_conformity(self.current_distance)
                
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numeric values")

    def check_conformity(self, distance):
        """Check if distance is within thresholds and update conformity status"""
        if hasattr(self, 'min_threshold') and hasattr(self, 'max_threshold'):
            if self.min_threshold <= distance <= self.max_threshold:
                self.conformity_label.configure(text="✅ PASS", text_color="#10b981")
                # Automatically trigger pass result
                self.root.after(100, lambda: self.process_result(True))
            else:
                self.conformity_label.configure(text="❌ FAIL", text_color="#ef4444")
                # Automatically trigger fail result
                self.root.after(100, lambda: self.process_result(False))
        else:
            self.conformity_label.configure(text="")
            
    def update_distance_display(self):
        try:
            self.distance_label.configure(text="{:.1f} cm".format(self.current_distance))
            
            if self.current_distance < 10:
                color = "#ef4444"
            elif self.current_distance < 20:
                color = "#f59e0b"
            elif self.current_distance < 100:
                color = "#10b981"
            else:
                color = "#06b6d4"
                
            self.distance_label.configure(text_color=color)
            
            # Check conformity automatically
            self.check_conformity(self.current_distance)
            
            if self.distance_history:
                min_dist = min(self.distance_history)
                max_dist = max(self.distance_history)
                avg_dist = sum(self.distance_history) / len(self.distance_history)
                
                self.min_distance_label.configure(text="Min: {:.1f} cm".format(min_dist))
                self.max_distance_label.configure(text="Max: {:.1f} cm".format(max_dist))
                self.avg_distance_label.configure(text="Avg: {:.1f} cm".format(avg_dist))
                
        except Exception as e:
            clean_msg = "Error updating distance display: {}".format(str(e))
            self.logger.error(clean_msg)
            self.distance_label.configure(text="-- cm", text_color="#06b6d4")

    def reset_distance_stats(self):
        if messagebox.askyesno("Confirmation", "Are you sure you want to reset the distance statistics?"):
            self.distance_history.clear()
            self.current_distance = 0.0
            
            self.distance_label.configure(text="-- cm")
            self.min_distance_label.configure(text="Min: -- cm")
            self.max_distance_label.configure(text="Max: -- cm")
            self.avg_distance_label.configure(text="Avg: -- cm")
            self.conformity_label.configure(text="")
            
            self.log_message("🔄 Distance statistics reset")

    def attempt_reconnection(self):
        max_attempts = 3
        for attempt in range(max_attempts):
            self.log_message("🔄 Reconnection attempt {}/{}".format(attempt + 1, max_attempts))
            time.sleep(2)
            try:
                if self.port:
                    self.arduino = serial.Serial(self.port, self.baudrate, timeout=1)
                    time.sleep(2)
                    self.is_running = True
                    self.start_reading_thread()
                    self.log_message("✅ Reconnection successful")
                    return
            except Exception as e:
                clean_msg = "Reconnection attempt {} failed: {}".format(attempt + 1, str(e))
                self.logger.error(clean_msg)
        
        self.log_message("❌ Reconnection failed")
        self.disconnect()

    def process_result(self, conforme):
        timestamp = datetime.now()
        
        if conforme:
            self.conforme_count += 1
            self.update_status("✅ PASS", "#10b981")
            self.log_message("✅ Result: PASS")
            result_text = "PASS"
        else:
            self.non_conforme_count += 1
            self.update_status("❌ FAIL", "#ef4444")
            self.log_message("❌ Result: FAIL")
            result_text = "FAIL"
        
        self.test_history.append({
            'timestamp': timestamp,
            'result': result_text,
            'conforme': conforme,
            'distance': self.current_distance
        })
        
        # Add to history tree with colors
        item_id = self.history_tree.insert('', 0, values=(
            timestamp.strftime("%H:%M:%S"),
            f"{'✅' if conforme else '❌'} {result_text}",
            f"{self.current_distance:.1f}",
            "< 100"
        ))
        
        # Color coding for tree items
        if conforme:
            self.history_tree.set(item_id, '✅ Result', '✅ PASS')
        else:
            self.history_tree.set(item_id, '✅ Result', '❌ FAIL')
        
        self.update_stats()
        self.play_notification_sound(conforme)

    def manual_test(self, conforme):
        if not self.session_start_time:
            self.session_start_time = datetime.now()
            self.start_connection_timer()
        
        if self.current_distance == 0.0:
            import random
            self.current_distance = round(random.uniform(10.0, 50.0), 1)
            self.distance_history.append(self.current_distance)
            self.update_distance_display()
        
        self.process_result(conforme)

    def update_status(self, text, color):
        def task():
            self.status_label.configure(text=text, text_color=color)
            
            if color == "#10b981":  # Green for conforme
                self.led_conforme.configure(
                    fg_color=("#10b981", "#059669"),
                    hover_color=("#10b981", "#059669"),
                    text_color="white"
                )
                self.led_non_conforme.configure(
                    fg_color=("#374151", "#374151"),
                    hover_color=("#374151", "#374151"),
                    text_color=("#9ca3af", "#6b7280")
                )
            elif color == "#ef4444":  # Red for non conforme
                self.led_non_conforme.configure(
                    fg_color=("#ef4444", "#dc2626"),
                    hover_color=("#ef4444", "#dc2626"),
                    text_color="white"
                )
                self.led_conforme.configure(
                    fg_color=("#374151", "#374151"),
                    hover_color=("#374151", "#374151"),
                    text_color=("#9ca3af", "#6b7280")
                )
            
            self.root.after(3000, self.reset_status)
            
        self.root.after(0, task)

    def reset_status(self):
        if self.is_running:
            self.status_label.configure(text="⏳ Waiting...", text_color="#94a3b8")
        else:
            self.status_label.configure(text="⏳ Waiting for connection...", text_color="#94a3b8")
        
        self.led_conforme.configure(
            fg_color=("#374151", "#374151"),
            hover_color=("#374151", "#374151"),
            text_color=("#9ca3af", "#6b7280")
        )
        self.led_non_conforme.configure(
            fg_color=("#374151", "#374151"),
            hover_color=("#374151", "#374151"),
            text_color=("#9ca3af", "#6b7280")
        )

    def update_stats(self):
        total = self.conforme_count + self.non_conforme_count
        success_rate = (self.conforme_count / total * 100) if total > 0 else 0
        
        self.stats_conforme_label.configure(text=str(self.conforme_count))
        self.stats_non_conforme_label.configure(text=str(self.non_conforme_count))
        self.stats_total_label.configure(text=str(total))
        self.success_rate_label.configure(text="{:.1f}%".format(success_rate))

    def reset_stats(self):
        if messagebox.askyesno("Confirmation", "Are you sure you want to reset all statistics?"):
            self.conforme_count = 0
            self.non_conforme_count = 0
            self.test_history.clear()
            
            # Clear tree
            for item in self.history_tree.get_children():
                self.history_tree.delete(item)
            
            self.update_stats()
            self.log_message("🔄 Statistics reset")

    def new_test_session(self):
        if messagebox.askyesno("New session", "Start a new test session?"):
            self.reset_stats()
            self.reset_distance_stats()
            self.session_start_time = datetime.now() if self.is_running else None
            self.log_message("🎉 New test session started")

    def toggle_auto_reconnect(self):
        self.auto_reconnect = self.auto_reconnect_var.get()
        self.save_config()

    def toggle_sound(self):
        self.sound_enabled = self.sound_var.get()
        self.save_config()

    def play_notification_sound(self, conforme):
        if self.sound_enabled:
            try:
                self.root.bell()
            except Exception:
                pass

    def test_connection(self):
        if not self.is_running:
            messagebox.showwarning("Test unavailable", "Please connect to the Arduino first.")
            return
        
        try:
            self.arduino.write(b"TEST\n")
            self.log_message("🧪 Test command sent")
        except Exception as e:
            messagebox.showerror("Test Error", "Unable to send the test command:\n{}".format(str(e)))

    def send_custom_command(self):
        if not self.is_running:
            messagebox.showwarning("Connection required", "Please connect to the Arduino first.")
            return
        
        # Create custom dialog for command input
        dialog = ctk.CTkInputDialog(text="Enter the command to send:", title="Custom Command")
        command = dialog.get_input()
        
        if command:
            try:
                self.arduino.write("{}\n".format(command).encode())
                self.log_message("📤 Command sent: {}".format(command))
            except Exception as e:
                messagebox.showerror("Error", "Unable to send the command:\n{}".format(str(e)))

    def export_to_csv(self):
        if not self.test_history:
            messagebox.showwarning("No data", "No tests to export.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Export to CSV"
        )
        
        if filename:
            try:
                with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                    fieldnames = ['Timestamp', 'Result', 'Conforme', 'Distance_cm']
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                    
                    writer.writeheader()
                    for test in self.test_history:
                        writer.writerow({
                            'Timestamp': test['timestamp'].strftime('%Y-%m-%d %H:%M:%S'),
                            'Result': test['result'],
                            'Conforme': test['conforme'],
                            'Distance_cm': test.get('distance', 0.0)
                        })
                
                self.log_message("📊 Data exported to: {}".format(filename))
                messagebox.showinfo("Export successful", "Data exported to:\n{}".format(filename))
            except Exception as e:
                messagebox.showerror("Export Error", "Unable to export data:\n{}".format(str(e)))

    def export_to_json(self):
        if not self.test_history:
            messagebox.showwarning("No data", "No tests to export.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Export to JSON"
        )
        
        if filename:
            try:
                export_data = {
                    'session_info': {
                        'start_time': self.session_start_time.isoformat() if self.session_start_time else None,
                        'export_time': datetime.now().isoformat(),
                        'total_tests': len(self.test_history),
                        'conforme_count': self.conforme_count,
                        'non_conforme_count': self.non_conforme_count,
                        'success_rate': (self.conforme_count / len(self.test_history) * 100) if self.test_history else 0
                    },
                    'distance_stats': {
                        'current_distance': self.current_distance,
                        'min_distance': min(self.distance_history) if self.distance_history else 0,
                        'max_distance': max(self.distance_history) if self.distance_history else 0,
                        'avg_distance': sum(self.distance_history) / len(self.distance_history) if self.distance_history else 0,
                        'thresholds': {
                            'min': getattr(self, 'min_threshold', 10.0),
                            'max': getattr(self, 'max_threshold', 50.0)
                        }
                    },
                    'tests': [
                        {
                            'timestamp': test['timestamp'].isoformat(),
                            'result': test['result'],
                            'conforme': test['conforme'],
                            'distance_cm': test.get('distance', 0.0)
                        }
                        for test in self.test_history
                    ]
                }
                
                with open(filename, 'w', encoding='utf-8') as jsonfile:
                    json.dump(export_data, jsonfile, indent=2, ensure_ascii=False)
                
                self.log_message("📋 Data exported to: {}".format(filename))
                messagebox.showinfo("Export successful", "Data exported to:\n{}".format(filename))
            except Exception as e:
                messagebox.showerror("Export Error", "Unable to export data:\n{}".format(str(e)))

    def save_log(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title="Save Log"
        )
        
        if filename:
            try:
                log_content = self.log_text.get("1.0", "end")
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(log_content)
                messagebox.showinfo("Save successful", "Log saved to:\n{}".format(filename))
            except Exception as e:
                messagebox.showerror("Error", "Unable to save the log:\n{}".format(str(e)))

    def reset_config(self):
        if messagebox.askyesno("Reset", "Are you sure you want to reset the configuration?"):
            try:
                if self.config_file.exists():
                    self.config_file.unlink()
                
                self.baudrate = 9600
                self.auto_reconnect = False
                self.sound_enabled = True
                
                self.baudrate_var.set("9600")
                self.auto_reconnect_var.set(False)
                self.sound_var.set(True)
                
                self.log_message("🔄 Configuration reset")
                messagebox.showinfo("Reset", "Configuration reset successfully.")
            except Exception as e:
                messagebox.showerror("Error", "Unable to reset configuration:\n{}".format(str(e)))

    def open_log_folder(self):
        import os
        import subprocess
        
        try:
            log_dir = Path.cwd()
            if os.name == 'nt':
                os.startfile(log_dir)
            elif os.name == 'posix':
                subprocess.run(['open' if 'darwin' in os.uname().sysname.lower() else 'xdg-open', log_dir])
        except Exception as e:
            messagebox.showerror("Error", "Unable to open folder:\n{}".format(str(e)))

    def show_help(self):
        help_text = """
    🚀 USER GUIDE - AZURA QUALITY CONTROL v3.0

    🔧 CONNECTION:
    - Select your Arduino COM port from the dropdown
    - Choose the correct baudrate (usually 9600)
    - Click "CONNECT" to establish the link
    - The status indicator turns green when connected

    📏 ULTRASONIC SENSOR:
    - Configure Min/Max thresholds in centimeters
    - Click "Apply" to save thresholds
    - Distance displays in real-time with color coding
    - Conformity tests run automatically based on thresholds

    🧪 MANUAL TESTS:
    - Use "Pass Test" and "Fail Test" buttons
    - Virtual LEDs light according to the result
    - History is recorded automatically

    📊 STATISTICS:
    - View real-time stats in the colored cards
    - See detailed history in the Statistics tab
    - Export your data to CSV or JSON for analysis

    ⚙️ SETTINGS:
    - Change appearance (Light/Dark/System)
    - Change color theme (blue/green/dark-blue)
    - Enable/disable notification sounds
    - Manage advanced configuration

    🔄 ADVANCED FEATURES:
    - Auto-reconnect on connection loss
    - Configuration auto-save
    - Full event logging
    - Modern responsive UI

    For more help, consult the full documentation.
        """
        
        help_window = ctk.CTkToplevel(self.root)
        help_window.title("📚 User Guide")
        help_window.geometry("700x600")
        help_window.transient(self.root)
        help_window.grab_set()
        
        # Header
        header_frame = ctk.CTkFrame(help_window, fg_color=("#3b82f6", "#1e40af"), height=80)
        header_frame.pack(fill="x", padx=20, pady=20)
        header_frame.pack_propagate(False)
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="📚 Complete User Guide",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=20)
        
        # Content
        text_widget = ctk.CTkTextbox(
            help_window,
            wrap="word",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=("#f8fafc", "#1f2937")
        )
        text_widget.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        text_widget.insert("1.0", help_text)
        text_widget.configure(state="disabled")
        
        # Close button
        close_btn = ctk.CTkButton(
            help_window,
            text="✅ Close",
            command=help_window.destroy,
            width=120,
            height=40,
            font=ctk.CTkFont(weight="bold")
        )
        close_btn.pack(pady=(0, 20))

    def show_about(self):
        about_text = """
    ⚡ ULTRASONIC MONITORING SYSTEM
    Version 1.0 - Modern Interface

    🎯 DEVELOPMENT:
    Project by Noreddine Akouchah

    🛠 TECHNOLOGIES:
    • Python 3.8+
    • CustomTkinter (Modern UI)
    • PySerial (Serial communication)
    • Threading (Real-time processing)

    🎨 FEATURES:
    • Modern dark interface with gradients
    • Colorful cards and smooth animations
    • Styled transparent tables
    • Virtual LED indicators
    • Real-time statistics
    • Advanced data export

    📅 LAST UPDATED:
    December 2025

    © 2025 - Ultrasonic Monitoring Quality Control System
    All rights reserved
        """
        
        # Create custom about dialog
        about_window = ctk.CTkToplevel(self.root)
        about_window.title("ℹ️ About")
        about_window.geometry("500x400")
        about_window.transient(self.root)
        about_window.grab_set()
        
        # Logo and title
        header_frame = ctk.CTkFrame(about_window, fg_color=("#8b5cf6", "#7c3aed"), height=100)
        header_frame.pack(fill="x", padx=20, pady=20)
        header_frame.pack_propagate(False)
        
        logo_label = ctk.CTkLabel(header_frame, text="⚡", font=ctk.CTkFont(size=40))
        logo_label.pack(side="left", padx=20, pady=30)
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="AZURA QUALITY CONTROL",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        title_label.pack(side="left", pady=30)
        
        # Content
        content_label = ctk.CTkLabel(
            about_window,
            text=about_text,
            font=ctk.CTkFont(size=12),
            justify="left"
        )
        content_label.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        # Close button
        close_btn = ctk.CTkButton(
            about_window,
            text="✅ Close",
            command=about_window.destroy,
            width=120,
            height=40,
            font=ctk.CTkFont(weight="bold")
        )
        close_btn.pack(pady=(0, 20))

    def clean_message_for_logging(self, msg):
        """Remove emoji characters from messages for logging compatibility"""
        import re
        # Remove all emoji characters that cause encoding issues
        emoji_pattern = re.compile("["
            "\U0001F600-\U0001F64F"  # emoticons
            "\U0001F300-\U0001F5FF"  # symbols & pictographs
            "\U0001F680-\U0001F6FF"  # transport & map symbols
            "\U0001F1E0-\U0001F1FF"  # flags (iOS)
            "\U00002702-\U000027B0"
            "\U000024C2-\U0001F251"
            "⚠️⚡🔧🎉📏🔄✅❌🧪📊🔌📡📤💾🗑️📋📁📚ℹ️📅⏱️🚀🎯🛠️🎨📱🔵🟣🟡🔵🟢🔴📜🔔🔊🌙🕐"
            "]+", flags=re.UNICODE)
        return emoji_pattern.sub('', msg).strip()

    def log_message(self, msg):
        def task():
            timestamp = datetime.now().strftime("%H:%M:%S")
            log_entry = f"[{timestamp}] {msg}\n"
            self.log_text.insert("end", log_entry)
            
            if self.auto_scroll_var.get():
                self.log_text.see("end")
            
            # Clean message for logging to avoid Unicode errors
            clean_msg = self.clean_message_for_logging(msg)
            if clean_msg:  # Only log if there's content after cleaning
                self.logger.info(clean_msg)
        
        self.root.after(0, task)

    def clear_log(self):
        if messagebox.askyesno("Confirmation", "Are you sure you want to clear the log?"):
            self.log_text.delete("1.0", "end")
            self.log_message("🗑️ Log cleared")

    def on_closing(self):
        if self.is_running:
            if messagebox.askyesno("Close", "A connection is active. Are you sure you want to exit?"):
                self.disconnect()
                self.save_config()
                self.root.destroy()
        else:
            self.save_config()
            self.root.destroy()

    def run(self):
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            self.logger.info("Application interrupted by user")
            self.on_closing()
        except Exception as e:
            clean_msg = "Unexpected error: {}".format(str(e))
            self.logger.error(clean_msg)
            messagebox.showerror("Critical Error", "An unexpected error occurred:\n{}".format(str(e)))

if __name__ == "__main__":
    try:
        app = ModernArduinoInterface()
        app.run()
    except Exception as e:
        print("Failed to start application: {}".format(str(e)))
        input("Press Enter to exit...")
