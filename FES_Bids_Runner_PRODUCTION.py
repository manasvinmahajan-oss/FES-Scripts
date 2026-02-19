#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
FES Bid Automation - Daily Runner GUI - PRODUCTION VERSION
- Uses production file paths
- Uses production SQL tables (Bids_Murley_D_Minus_1, Bids_SU_D_Minus_1)
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime, timedelta
import sys
import os
from pathlib import Path
from datetime import datetime, timedelta

# Import the PRODUCTION master script classes
try:
    from FES_MasterScript_PRODUCTION import grab_forecast_data, MurleyGUCompiler, SupplyUnitCompiler
except ImportError:
    print("ERROR: Cannot import FES_MasterScript_PRODUCTION. Make sure FES_MasterScript_PRODUCTION.py is in the same directory.")
    sys.exit(1)


class FESBidApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FES Bid Automation - PRODUCTION Runner")
        self.root.geometry("900x900")
        self.root.resizable(True, True)
        self.root.minsize(800, 700)  # Minimum size

        # Apply modern styling
        style = ttk.Style()
        style.theme_use('clam')

        # Configure colors - Professional blue theme
        self.colors = {
            'header': '#2c3e50',  # Dark blue-grey
            'bg': '#f5f6fa',  # Light grey-blue background
            'accent': '#3498db',  # Professional blue
            'success': '#27ae60',  # Green
            'warning': '#f39c12',  # Orange (not scary red)
            'text_dark': '#2c3e50',
            'text_light': '#7f8c8d',
            'border': '#dfe4ea'  # Light border
        }

        # Create main container with scrollbar
        main_container = tk.Frame(root)
        main_container.pack(fill="both", expand=True)

        # Canvas for scrolling
        canvas = tk.Canvas(main_container, bg=self.colors['bg'])
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors['bg'])

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # ==========================================
        # HEADER
        # ==========================================
        header_frame = tk.Frame(scrollable_frame, bg=self.colors['header'], height=100)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)

        tk.Label(
            header_frame,
            text="FES Bids Manager",
            font=("Segoe UI", 24, "bold"),
            bg=self.colors['header'],
            fg="white"
        ).pack(pady=(20, 5))

        tk.Label(
            header_frame,
            text="Day-Ahead & Intraday Bid Compilation",
            font=("Segoe UI", 11),
            bg=self.colors['header'],
            fg="#95a5a6"
        ).pack()

        # ==========================================
        # MAIN CONTENT AREA
        # ==========================================
        main_frame = tk.Frame(scrollable_frame, bg=self.colors['bg'], padx=50, pady=30)
        main_frame.pack(fill="both", expand=True)

        # ==========================================
        # DATE SELECTION SECTION
        # ==========================================
        date_section = tk.LabelFrame(
            main_frame,
            text="Trading Day Selection",
            font=("Segoe UI", 13, "bold"),
            bg=self.colors['bg'],
            fg=self.colors['text_dark'],
            padx=30,
            pady=25,
            relief="groove",
            borderwidth=2
        )
        date_section.pack(pady=(0, 25), fill="x")

        # Date input frame
        date_input_frame = tk.Frame(date_section, bg=self.colors['bg'])
        date_input_frame.pack(pady=15)

        # Day
        tk.Label(
            date_input_frame,
            text="Day:",
            font=("Segoe UI", 12, "bold"),
            bg=self.colors['bg'],
            fg=self.colors['text_dark']
        ).grid(row=0, column=0, padx=10, sticky="e")

        self.day_var = tk.StringVar(value=self.get_tomorrow()[0])
        day_spinbox = ttk.Spinbox(
            date_input_frame,
            from_=1,
            to=31,
            textvariable=self.day_var,
            width=8,
            font=("Segoe UI", 13)
        )
        day_spinbox.grid(row=0, column=1, padx=10)

        # Month
        tk.Label(
            date_input_frame,
            text="Month:",
            font=("Segoe UI", 12, "bold"),
            bg=self.colors['bg'],
            fg=self.colors['text_dark']
        ).grid(row=0, column=2, padx=10, sticky="e")

        self.month_var = tk.StringVar(value=self.get_tomorrow()[1])
        month_spinbox = ttk.Spinbox(
            date_input_frame,
            from_=1,
            to=12,
            textvariable=self.month_var,
            width=8,
            font=("Segoe UI", 13)
        )
        month_spinbox.grid(row=0, column=3, padx=10)

        # Year
        tk.Label(
            date_input_frame,
            text="Year:",
            font=("Segoe UI", 12, "bold"),
            bg=self.colors['bg'],
            fg=self.colors['text_dark']
        ).grid(row=0, column=4, padx=10, sticky="e")

        self.year_var = tk.StringVar(value=self.get_tomorrow()[2])
        year_spinbox = ttk.Spinbox(
            date_input_frame,
            from_=2024,
            to=2030,
            textvariable=self.year_var,
            width=7,
            font=("Segoe UI", 12)
        )
        year_spinbox.grid(row=0, column=5, padx=5)

        # ==========================================
        # BID TYPE SELECTION
        # ==========================================
        bid_type_section = tk.LabelFrame(
            main_frame,
            text="Select Process",
            font=("Segoe UI", 13, "bold"),
            bg=self.colors['bg'],
            fg=self.colors['text_dark'],
            padx=30,
            pady=25,
            relief="groove",
            borderwidth=2
        )
        bid_type_section.pack(pady=(0, 25), fill="x")
        
        self.bid_type_var = tk.StringVar(value="D-X")
        
        radio_frame = tk.Frame(bid_type_section, bg=self.colors['bg'])
        radio_frame.pack(pady=15)
        
        tk.Radiobutton(
            radio_frame,
            text="D-X Bids (Change date to adjust lags)",
            variable=self.bid_type_var,
            value="D-X",
            font=("Segoe UI", 11),
            bg=self.colors['bg'],
            fg=self.colors['text_dark'],
            selectcolor=self.colors['bg'],
            activebackground=self.colors['bg'],
            cursor="hand2"
        ).pack(anchor="w", pady=5)
        
        tk.Radiobutton(
            radio_frame,
            text="IDA-1 Adjustment (Generation Forecast Change)",
            variable=self.bid_type_var,
            value="IDA-1",
            font=("Segoe UI", 11),
            bg=self.colors['bg'],
            fg=self.colors['text_dark'],
            selectcolor=self.colors['bg'],
            activebackground=self.colors['bg'],
            cursor="hand2"
        ).pack(anchor="w", pady=5)
        
        # Info label for IDA-1
        ida_info = tk.Label(
            bid_type_section,
            text="ℹ️ IDA-1: Calculates adjustment bids based on difference between D-1 and IDA-1 forecasts",
            font=("Segoe UI", 9, "italic"),
            bg=self.colors['bg'],
            fg=self.colors['text_light']
        )
        ida_info.pack(pady=(0, 5))

        # ==========================================
        # UPLOAD OPTIONS SECTION
        # ==========================================
        upload_section = tk.LabelFrame(
            main_frame,
            text="Database Upload Options",
            font=("Segoe UI", 13, "bold"),
            bg=self.colors['bg'],
            fg=self.colors['text_dark'],
            padx=30,
            pady=25,
            relief="groove",
            borderwidth=2
        )
        upload_section.pack(pady=(0, 25), fill="x")

        self.upload_sql_var = tk.BooleanVar(value=False)
        self.create_ppt_var = tk.BooleanVar(value=False)
        self.friday_mode_var = tk.BooleanVar(value=False)

        upload_check = tk.Checkbutton(
            upload_section,
            text="Upload to Fabric SQL Database",
            variable=self.upload_sql_var,
            font=("Segoe UI", 11, "bold"),
            bg=self.colors['bg'],
            fg=self.colors['accent'],
            selectcolor=self.colors['bg'],
            activebackground=self.colors['bg'],
            activeforeground=self.colors['accent']
        )
        upload_check.pack(anchor="w", pady=(0, 12))
        
        ppt_check = tk.Checkbutton(
            upload_section,
            text="Generate PowerPoint & Send Email",
            variable=self.create_ppt_var,
            font=("Segoe UI", 11, "bold"),
            bg=self.colors['bg'],
            fg=self.colors['text_dark'],
            selectcolor=self.colors['bg'],
            activebackground=self.colors['bg'],
            activeforeground=self.colors['text_dark']
        )
        ppt_check.pack(anchor="w", pady=(0, 8))
        
        # Friday mode checkbox (only shown when PPT is checked)
        friday_check = tk.Checkbutton(
            upload_section,
            text="   → Include Weekend Forecasts (Friday Mode)",
            variable=self.friday_mode_var,
            font=("Segoe UI", 10, "italic"),
            bg=self.colors['bg'],
            fg=self.colors['text_dark'],
            selectcolor=self.colors['bg'],
            activebackground=self.colors['bg'],
            activeforeground=self.colors['success']
        )
        friday_check.pack(anchor="w", pady=(0, 15))

        # Warning label
        warning_frame = tk.Frame(upload_section, bg="#ffe6e6", relief="solid", borderwidth=1)
        warning_frame.pack(fill="x", pady=(0, 15))
        
        info_frame = tk.Frame(upload_section, bg="#e3f2fd", relief="solid", borderwidth=1)
        info_frame.pack(fill="x", pady=(0, 15))
        
        info_label_header = tk.Label(
            info_frame,
            text="SQL Upload Target Tables:",
            font=("Segoe UI", 9, "bold"),
            bg="#e3f2fd",
            fg=self.colors['accent'],
            justify="left",
            padx=15
        )
        info_label_header.pack(anchor="w", pady=(10, 5))
        
        info_label_content = tk.Label(
            info_frame,
            text="• Bids_Murley_D_Minus_1 / D_Minus_X\n"
                 "• Bids_SU_D_Minus_1 / D_Minus_X\n"
                 "• Generation_D_Minus_1 / D_Minus_X",
            font=("Segoe UI", 9),
            bg="#e3f2fd",
            fg=self.colors['text_dark'],
            justify="left",
            padx=15
        )
        info_label_content.pack(anchor="w", pady=(0, 10))

        note_label = tk.Label(
            upload_section,
            text="Note: Files are always saved to I: drive. Uncheck boxes to skip SQL upload or PPT/email generation.",
            font=("Segoe UI", 9, "italic"),
            bg=self.colors['bg'],
            fg=self.colors['text_light'],
            justify="left"
        )
        note_label.pack(anchor="w")

        # ==========================================
        # RUN BUTTON
        # ==========================================
        button_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        button_frame.pack(pady=20)

        self.run_button = tk.Button(
            button_frame,
            text="RUN COMPILATION",
            command=self.run_workflow,
            bg=self.colors['accent'],
            fg="white",
            font=("Segoe UI", 13, "bold"),
            padx=50,
            pady=15,
            relief="flat",
            cursor="hand2",
            borderwidth=0
        )
        self.run_button.pack()

        # ==========================================
        # STATUS LOG
        # ==========================================
        log_section = tk.LabelFrame(
            main_frame,
            text="Execution Log",
            font=("Segoe UI", 13, "bold"),
            bg=self.colors['bg'],
            fg=self.colors['text_dark'],
            padx=30,
            pady=25,
            relief="groove",
            borderwidth=2
        )
        log_section.pack(pady=(0, 25), fill="both", expand=True)

        self.status_text = scrolledtext.ScrolledText(
            log_section,
            width=90,
            height=20,
            font=("Consolas", 10),
            bg="#2c3e50",
            fg="#ecf0f1",
            insertbackground="white",
            relief="flat",
            wrap="word"
        )
        self.status_text.pack(fill="both", expand=True, padx=5, pady=5)

        # Add initial message
        self.log_status("Ready to compile bids.")
        self.log_status("Select trading date and bid type, then click RUN COMPILATION.")
        self.log_status("SQL upload is disabled by default - check box to enable.")
        self.status_text.config(state="disabled")

        # ==========================================
        # FOOTER
        # ==========================================
        footer = tk.Frame(scrollable_frame, bg=self.colors['header'], height=35)
        footer.pack(side="bottom", fill="x")
        footer.pack_propagate(False)

        tk.Label(
            footer,
            text="Flogas Ireland · FES Bid Compilation System",
            font=("Segoe UI", 9),
            bg=self.colors['header'],
            fg="#95a5a6"
        ).pack(pady=8)

    def get_tomorrow(self):
        """Get tomorrow's date for D-1 bidding"""
        tomorrow = datetime.now() + timedelta(days=1)
        return str(tomorrow.day).zfill(2), str(tomorrow.month).zfill(2), str(tomorrow.year)

    def set_tomorrow(self):
        """Set date to tomorrow (D-1)"""
        day, month, year = self.get_tomorrow()
        self.day_var.set(day)
        self.month_var.set(month)
        self.year_var.set(year)

    def set_today(self):
        """Set date to today"""
        today = datetime.now()
        self.day_var.set(str(today.day).zfill(2))
        self.month_var.set(str(today.month).zfill(2))
        self.year_var.set(str(today.year))

    def log_status(self, message):
        """Add message to status log"""
        self.status_text.config(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.insert("end", f"[{timestamp}] {message}\n")
        self.status_text.see("end")
        self.status_text.config(state="disabled")
        self.root.update()

    def clear_status(self):
        """Clear status log"""
        self.status_text.config(state="normal")
        self.status_text.delete("1.0", "end")
        self.status_text.config(state="disabled")

    def validate_date(self):
        """Validate and format the selected date"""
        try:
            day = int(self.day_var.get())
            month = int(self.month_var.get())
            year = int(self.year_var.get())

            if not (1 <= day <= 31):
                raise ValueError("Day must be between 1 and 31")
            if not (1 <= month <= 12):
                raise ValueError("Month must be between 1 and 12")
            if not (2024 <= year <= 2030):
                raise ValueError("Year must be between 2024 and 2030")

            date_obj = datetime(year, month, day)
            formatted_date = date_obj.strftime("%d/%m/%Y")
            return formatted_date

        except ValueError as e:
            messagebox.showerror("Invalid Date", f"Please enter a valid date.\n\nError: {str(e)}")
            return None

    def run_workflow(self):
        """Main workflow execution"""
        input_date = self.validate_date()
        if not input_date:
            return

        upload_sql = self.upload_sql_var.get()
        create_ppt = self.create_ppt_var.get()
        friday_mode = self.friday_mode_var.get()
        bid_type = self.bid_type_var.get()  # Get bid type (D-1 or IDA-1)
        
        upload_status = "ENABLED" if upload_sql else "DISABLED"
        ppt_status = "ENABLED" if create_ppt else "DISABLED"
        friday_status = " (Friday Mode: Include Weekend)" if (create_ppt and friday_mode) else ""

        # Build confirmation message based on bid type
        confirm_msg = f"Confirm Bid Compilation\n\n"
        confirm_msg += f"Trading Day: {input_date}\n"
        confirm_msg += f"Bid Type: {bid_type}\n"
        
        if bid_type == "D-X":
            confirm_msg += f"SQL Upload: {upload_status}\n"
            confirm_msg += f"PowerPoint: {ppt_status}{friday_status}\n\n"
            confirm_msg += "Steps:\n"
            confirm_msg += "  1. Generate Generation Forecast\n"
            confirm_msg += "  2. Compile Murley GU Bids\n"
            confirm_msg += "  3. Compile Supply Unit Bids\n"
            confirm_msg += "  4. Save all files to I: drive\n"

            if create_ppt:
                if friday_mode:
                    confirm_msg += "  5. Create PowerPoint (with weekend forecasts) & Send email\n"
                else:
                    confirm_msg += "  5. Create PowerPoint & Send email\n"

            if upload_sql:
                confirm_msg += "\nSQL Upload Enabled:\n"
                confirm_msg += "  • Generation_D_Minus_1 / D_Minus_X\n"
                confirm_msg += "  • Bids_Murley_D_Minus_1 / D_Minus_X\n"
                confirm_msg += "  • Bids_SU_D_Minus_1 / D_Minus_X\n"
            else:
                confirm_msg += "\nSQL upload disabled (files only)\n"
        
        else:  # IDA-1
            confirm_msg += f"SQL Upload: {upload_status}\n\n"
            confirm_msg += "IDA-1 Workflow:\n"
            confirm_msg += "  1. Load D-1 forecast (morning)\n"
            confirm_msg += "  2. Load IDA-1 forecast (evening)\n"
            confirm_msg += "  3. Calculate adjustment = IDA-1 - D-1\n"
            confirm_msg += "  4. Generate IDA-1 adjustment bids\n"
            confirm_msg += "  5. Create comparison Excel & charts\n"
            confirm_msg += "  6. Save to IDA-1 bid directories\n"
            confirm_msg += "\n⚠️ Ensure IDA-1 forecast file exists!\n"
            confirm_msg += "(evening generation forecast update)\n"

        confirm_msg += "\nProceed with execution?"

        # Show confirmation
        confirm = messagebox.askyesno(
            "Confirm Execution", 
            confirm_msg,
            icon='question'
        )
        if not confirm:
            return

        # Disable button and change text
        self.run_button.config(state="disabled", text="RUNNING...", bg="#7f8c8d")
        self.clear_status()

        try:
            # ==========================================
            # ROUTE TO CORRECT WORKFLOW
            # ==========================================
            if bid_type == "IDA-1":
                # IDA-1 WORKFLOW
                self.run_ida1_workflow(input_date, upload_sql)
            else:
                # D-1 WORKFLOW (original)
                self.run_d1_workflow(input_date, upload_sql, create_ppt, friday_mode)
            
            # SUCCESS
            self.log_status("")
            self.log_status("=" * 70)
            self.log_status("EXECUTION COMPLETED SUCCESSFULLY")
            self.log_status("=" * 70)
            messagebox.showinfo("Success", f"{bid_type} bid compilation completed successfully!")

        except Exception as e:
            self.log_status("")
            self.log_status("=" * 70)
            self.log_status(f"EXECUTION FAILED")
            self.log_status(f"Error: {str(e)}")
            self.log_status("=" * 70)
            messagebox.showerror("Execution Failed", f"Error during {bid_type} compilation:\n\n{str(e)}")

        finally:
            # Re-enable button
            self.run_button.config(state="normal", text="RUN COMPILATION", bg=self.colors['accent'])
    
    def run_d1_workflow(self, input_date, upload_sql, create_ppt, friday_mode):
        """Execute D-1 bid workflow"""
        upload_status = "ENABLED" if upload_sql else "DISABLED"
        
        # Calculate lag dynamically based on trading date
        from datetime import datetime, timedelta
        trading_date = datetime.strptime(input_date, "%d/%m/%Y")
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        days_ahead = (trading_date - today).days
        lag = f"D-{days_ahead}"
        
        try:
            # ==========================================
            # STEP 1: Generation Forecast
            # ==========================================
            self.log_status("=" * 70)
            self.log_status("STEP 1/3: GENERATION FORECAST")
            self.log_status("=" * 70)
            self.log_status(f"Trading Day: {input_date}")
            self.log_status(f"Lag: {lag} ({days_ahead} days ahead)")
            self.log_status(f"SQL Upload: {upload_status}")

            try:
                forecast_df, lag, gen_upload_success = grab_forecast_data(input_date, upload_to_sql=upload_sql)
                self.log_status(f"[OK] Generation forecast complete: {len(forecast_df)} periods")
                if upload_sql:
                    table_name = "Generation_D_Minus_1" if lag == "D-1" else "Generation_D_Minus_X"
                    if gen_upload_success:
                        self.log_status(f"[OK] Uploaded to {table_name} (PRODUCTION)")
                    else:
                        self.log_status(f"[ERROR] FAILED to upload to {table_name}")
                else:
                    self.log_status("[SKIP] File saved to I: drive (SQL upload disabled)")
            except Exception as e:
                self.log_status(f"[ERROR] FAILED: {str(e)}")
                raise

            # ==========================================
            # STEP 2: Murley GU Bids
            # ==========================================
            self.log_status("")
            self.log_status("=" * 70)
            self.log_status("STEP 2/3: MURLEY GU BIDS")
            self.log_status("=" * 70)

            try:
                gu_compiler = MurleyGUCompiler()
                agg_gu, dam_gu, ets_gu, gu_upload_success = gu_compiler.run(
                    bid_date=input_date,
                    lag=lag,
                    upload_sql=upload_sql,
                    use_production=True  # ALWAYS use production table
                )
                self.log_status(f"[OK] GU bids compiled: {len(agg_gu)} periods")
                if upload_sql:
                    table_name = "Bids_Murley_D_Minus_1" if lag == "D-1" else "Bids_Murley_D_Minus_X"
                    if gu_upload_success:
                        self.log_status(f"[OK] Uploaded to {table_name} (PRODUCTION)")
                    else:
                        self.log_status(f"[ERROR] FAILED to upload to {table_name}")
                else:
                    self.log_status("[SKIP] Files saved to I: drive (SQL upload disabled)")
            except Exception as e:
                self.log_status(f"[ERROR] FAILED: {str(e)}")
                raise

            # ==========================================
            # STEP 3: Supply Unit Bids
            # ==========================================
            self.log_status("")
            self.log_status("=" * 70)
            self.log_status("STEP 3/3: SUPPLY UNIT BIDS")
            self.log_status("=" * 70)

            try:
                su_compiler = SupplyUnitCompiler()
                agg_su, dam_su, ets_su, su_upload_success = su_compiler.run(
                    bid_date=input_date,
                    lag=lag,
                    upload_sql=upload_sql,
                    use_production=True  # ALWAYS use production table
                )
                self.log_status(f"[OK] SU bids compiled: {len(agg_su)} periods")
                if upload_sql:
                    table_name = "Bids_SU_D_Minus_1" if lag == "D-1" else "Bids_SU_D_Minus_X"
                    if su_upload_success:
                        self.log_status(f"[OK] Uploaded to {table_name} (PRODUCTION)")
                    else:
                        self.log_status(f"[ERROR] FAILED to upload to {table_name}")
                else:
                    self.log_status("[SKIP] Files saved to I: drive (SQL upload disabled)")
            except Exception as e:
                self.log_status(f"[ERROR] FAILED: {str(e)}")
                raise

            # ==========================================
            # STEP 4: PowerPoint Presentation (Optional)
            # ==========================================
            if create_ppt:
                friday_mode = self.friday_mode_var.get()
                
                self.log_status("")
                self.log_status("=" * 70)
                if friday_mode:
                    self.log_status("STEP 4/4: POWERPOINT & EMAIL (FRIDAY MODE)")
                else:
                    self.log_status("STEP 4/4: POWERPOINT & EMAIL")
                self.log_status("=" * 70)

                try:
                    from FES_MasterScript_PRODUCTION import create_forecast_presentation
                    
                    # Get chart paths from the compilers
                    gu_chart_path = Path.cwd() / "output" / f"GU_504260_Bid_Chart_{datetime.strptime(input_date, '%d/%m/%Y').strftime('%d.%m.%Y')}.png"
                    su_chart_path = Path.cwd() / "output" / f"SU_400130_Bid_Chart_{datetime.strptime(input_date, '%d/%m/%Y').strftime('%d.%m.%Y')}.png"
                    
                    ppt_path = create_forecast_presentation(
                        input_date,
                        gu_chart_path if gu_chart_path.exists() else None,
                        su_chart_path if su_chart_path.exists() else None,
                        send_email=True,  # Always send email when PPT is generated
                        force_friday_mode=friday_mode  # Pass Friday mode setting
                    )
                    
                    if ppt_path:
                        self.log_status(f"[OK] PowerPoint created: {ppt_path}")
                        if friday_mode:
                            self.log_status(f"[OK] Friday mode: Weekend forecasts (Sat/Sun/Mon) included")
                        self.log_status(f"[OK] Email sent to isemtrading@flogas.ie")
                    else:
                        self.log_status("[ERROR] PowerPoint/Email generation failed (check logs)")
                except Exception as e:
                    self.log_status(f"[ERROR] PowerPoint/Email failed: {str(e)}")
                    # Don't raise - PPT/Email is optional
        
        except Exception as e:
            # Re-raise to be caught by main workflow handler
            raise
    
    def run_ida1_workflow(self, input_date, upload_sql):
        """Execute IDA-1 bid workflow"""
        
        upload_status = "ENABLED" if upload_sql else "DISABLED"
        
        try:
            # Clear cached IDA-1 module to ensure we get the latest version
            import sys
            if 'FES_IDA1_Compiler' in sys.modules:
                del sys.modules['FES_IDA1_Compiler']
            
            # Import IDA-1 compiler
            from FES_IDA1_Compiler import compile_ida1_bids
            
            # ==========================================
            # IDA-1 COMPILATION
            # ==========================================
            self.log_status("=" * 70)
            self.log_status("IDA-1 BID COMPILATION WORKFLOW")
            self.log_status("=" * 70)
            self.log_status(f"Trading Day: {input_date}")
            self.log_status(f"SQL Upload: {upload_status}")
            self.log_status("")
            
            if upload_sql:
                self.log_status("[STEP 1/5] Downloading IDA-1 forecast from Meteologica...")
                self.log_status("[STEP 2/5] Loading D-1 forecast (from morning)...")
                self.log_status("[STEP 3/5] Calculating adjustment (D-1 - IDA-1)...")
                self.log_status("[STEP 4/5] Uploading IDA-1 bids to SQL (dbo.ida1_bids)...")
                self.log_status("[STEP 5/5] Creating IDA Excel with all sheets and charts...")
            else:
                self.log_status("[STEP 1/4] Downloading IDA-1 forecast from Meteologica...")
                self.log_status("[STEP 2/4] Loading D-1 forecast (from morning)...")
                self.log_status("[STEP 3/4] Calculating adjustment (D-1 - IDA-1)...")
                self.log_status("[STEP 4/4] Creating IDA Excel with all sheets and charts...")
            
            # Run IDA-1 compilation (returns single Excel file)
            ida_excel = compile_ida1_bids(input_date, upload_sql=upload_sql)
            
            self.log_status("")
            self.log_status("File created:")
            self.log_status(f"  OK IDA Excel: {ida_excel.name}")
            self.log_status("")
            if upload_sql:
                self.log_status("SQL Upload:")
                self.log_status(f"  OK Uploaded to dbo.ida1_bids and Generation_D_Minus_X")
                self.log_status("")
            else:
                self.log_status("SQL Upload: DISABLED (files saved to I: drive only)")
                self.log_status("")
            self.log_status("Excel contains:")
            self.log_status("  - Sheet 1: D-1 Forecast (morning)")
            self.log_status("  - Sheet 2: IDA-1 Forecast (evening)")
            self.log_status("  - Sheet 3: Adjustment Calculation")
            self.log_status("  - Sheet 4: IDA-1 Bids")
            self.log_status("  - Sheet 5: Charts (D-1 vs IDA-1 + Adjustment)")
            
        except Exception as e:
            self.log_status(f"X IDA-1 FAILED: {str(e)}")
            raise


def main():
    """Launch the GUI application"""
    root = tk.Tk()
    app = FESBidApp(root)

    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    root.mainloop()


if __name__ == "__main__":
    main()