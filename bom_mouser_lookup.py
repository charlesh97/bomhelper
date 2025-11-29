#!/usr/bin/env python3
"""Main GUI application for BOM Mouser part lookup."""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from typing import Dict, Any, List, Optional
import webbrowser
import csv
from pathlib import Path
import logging
import json
import time
import re
from datetime import datetime
import google.generativeai as genai

from bom_parser import BOMParser
# SpecParser is deprecated, functionality moved here
from mouser_api import MouserAPI
from part_ranker import RankingEngine
from config import Config

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),  # Log to console/terminal
        logging.FileHandler("bom_helper.log")  # Log to file
    ]
)
logger = logging.getLogger(__name__)



class BOMMouserLookupApp:
    """Main application window."""
    
    def __init__(self, root):
        """Initialize the application."""
        logger.info("Initializing BOM Mouser Lookup App")
        self.root = root
        self.root.title("BOM Mouser Part Lookup Tool")
        self.root.geometry("1400x800")
        
        # Initialize components
        self.config = Config()
        self.bom_parser = BOMParser()
        # self.spec_parser = SpecParser(self.config)  # Deprecated
        
        # Initialize Gemini
        self.gemini_model = None
        if self.config.get_gemini_api_key():
            try:
                genai.configure(api_key=self.config.get_gemini_api_key())
                self.gemini_model = genai.GenerativeModel('gemini-2.5-flash')
                logger.info("Gemini initialized for search term generation")
            except Exception as e:
                logger.error(f"Failed to initialize Gemini: {e}")
        else:
            logger.warning("Gemini API key not found - keyword search will be limited")
        
        try:
            self.mouser_api = MouserAPI(self.config)
        except ValueError as e:
            logger.error(f"Configuration error: {e}")
            messagebox.showerror("Configuration Error", 
                               f"Mouser API key not configured: {e}\n\nPlease check keys.txt file.")
            self.mouser_api = None
        
        self.ranker = RankingEngine(
            stock_weight=0.3,
            price_weight=0.5,
            lifecycle_weight=0.1,
            package_match_weight=0.1
        )
        
        # Data storage
        self.components = []
        self.consolidated_parts = []
        self.column_mapping = {}  # Maps normalized column names to original BOM column names
        self.selected_parts = {}  # Maps part key to selected Mouser part
        self.current_search_results = {}  # Maps part key to list of Mouser parts
        self.current_search_index = {}  # Maps part key to current displayed index
        self.checkbox_vars = {}  # Maps part_key to list of checkbox variables for that part
        self.radio_vars = {}  # Maps part_key to radio button variable (StringVar or IntVar)
        self.result_frames = {}  # Maps (part_key, index) to frame widget for visual updates
        self.part_selected = {}  # Maps part_key to boolean (checkbox state)
        self.editing_cell = None  # Track currently editing cell (item_id, column)
        # Batch navigation
        self.current_batch_index = 0
        self.batch_part_keys = []
        self.batch_results = {}
        # Sort preference (persists across searches)
        self.sort_preference = tk.StringVar(value='stock')  # 'stock' or 'price'
        self.current_displayed_results = {}  # Store currently displayed results for re-sorting
        # Index-based part_key mapping
        self.part_key_to_index = {}  # Maps part_key to index in consolidated_parts
        # Store last search keyword used for each part (to show in custom search)
        self.last_search_keywords = {}  # Maps part_key to last keyword used
        
        # Build UI
        self.build_ui()
    
    def build_ui(self):
        """Build the user interface."""
        # Menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open BOM...", command=self.open_bom)
        file_menu.add_separator()
        file_menu.add_command(label="Save BOM State...", command=self.save_bom_state)
        file_menu.add_command(label="Load BOM State...", command=self.load_bom_state)
        file_menu.add_separator()
        file_menu.add_command(label="Export BOM...", command=self.export_bom)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="API Keys...", command=self.show_api_keys_dialog)
        
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)  # PanedWindow column (resizable)
        main_frame.columnconfigure(1, weight=0)   # Options column (fixed)
        
        # Create PanedWindow for resizable BOM and Results columns
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        # Left panel: Consolidated parts table
        left_frame = ttk.LabelFrame(paned, text="BOM Parts", padding="5")
        paned.add(left_frame, weight=1)
        left_frame.columnconfigure(0, weight=1)
        left_frame.columnconfigure(1, weight=0)  # Scrollbar column
        left_frame.rowconfigure(1, weight=1)  # Table row
        left_frame.rowconfigure(2, weight=0)  # Horizontal scrollbar row
        
        # Add hint label about double-click editing
        hint_label = ttk.Label(left_frame, text="💡 Tip: Double-click any cell to edit", 
                              font=('TkDefaultFont', 8), foreground='gray')
        hint_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 2))
        
        # Parts treeview - start blank, will be populated when BOM is loaded
        # First column will be checkbox, then data columns
        # Use 'tree headings' to show both the checkbox column (#0) and data column headings
        self.parts_tree = ttk.Treeview(left_frame, columns=(), show='tree headings', height=20)
        
        parts_v_scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.parts_tree.yview)
        parts_h_scrollbar = ttk.Scrollbar(left_frame, orient=tk.HORIZONTAL, command=self.parts_tree.xview)
        self.parts_tree.configure(yscrollcommand=parts_v_scrollbar.set, xscrollcommand=parts_h_scrollbar.set)
        
        self.parts_tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        parts_v_scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        parts_h_scrollbar.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        # Removed <<TreeviewSelect>> binding - no longer clears results on selection
        # self.parts_tree.bind('<<TreeviewSelect>>', self.on_part_selected)
        self.parts_tree.bind('<Double-1>', self.on_cell_double_click)
        self.parts_tree.bind('<Button-1>', self.on_cell_click)
        
        # Center panel: Search results
        right_frame = ttk.LabelFrame(paned, text="Mouser Search Results", padding="5")
        paned.add(right_frame, weight=2)
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)  # Results canvas
        right_frame.rowconfigure(1, weight=0)  # Get More Parts button
        right_frame.rowconfigure(2, weight=0)  # OK button
        
        # Results container (scrollable)
        results_canvas = tk.Canvas(right_frame)
        results_scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=results_canvas.yview)
        self.results_frame = ttk.Frame(results_canvas)
        
        results_canvas.create_window((0, 0), window=self.results_frame, anchor=tk.NW)
        results_canvas.configure(yscrollcommand=results_scrollbar.set)
        
        results_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        results_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configure scrolling
        def update_scrollregion(event=None):
            results_canvas.configure(scrollregion=results_canvas.bbox('all'))
        
        def configure_canvas_width(event):
            canvas_width = event.width
            results_canvas.itemconfig('window', width=canvas_width)
        
        self.results_frame.bind('<Configure>', update_scrollregion)
        results_canvas.bind('<Configure>', configure_canvas_width)
        
        # Mouse wheel scrolling (works on Windows and macOS)
        def on_mousewheel(event):
            # macOS uses delta differently
            if event.delta:
                delta = -1 * (event.delta / 120)
            else:
                # Some systems use different event format
                delta = -1 if event.num == 4 else 1
            results_canvas.yview_scroll(int(delta), "units")
        
        # Bind mouse wheel events
        results_canvas.bind("<MouseWheel>", on_mousewheel)
        results_canvas.bind("<Button-4>", on_mousewheel)  # Linux/Mac scroll up
        results_canvas.bind("<Button-5>", on_mousewheel)  # Linux/Mac scroll down
        
        # Store canvas reference for scroll region updates
        self.results_canvas = results_canvas
        
        # Button frame for Get More Parts and Custom Search buttons
        button_frame = ttk.Frame(right_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=5, sticky=(tk.W, tk.E))
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        
        # Get more parts button (initially hidden)
        self.more_parts_btn = ttk.Button(button_frame, text="Get More Parts", 
                                        command=self.get_more_parts, state=tk.DISABLED)
        self.more_parts_btn.grid(row=0, column=0, padx=(0, 2), sticky=(tk.W, tk.E))
        
        # Custom search button (initially hidden)
        self.custom_search_btn = ttk.Button(button_frame, text="Custom Search", 
                                           command=self.show_custom_search_dialog, state=tk.DISABLED)
        self.custom_search_btn.grid(row=0, column=1, padx=(2, 0), sticky=(tk.W, tk.E))
        
        # OK button at bottom (initially disabled)
        self.confirm_part_btn = ttk.Button(right_frame, text="OK - This is the part I want", 
                                          command=self.confirm_selected_part, state=tk.DISABLED)
        self.confirm_part_btn.grid(row=2, column=0, columnspan=2, pady=5, sticky=(tk.W, tk.E))
        
        # Right panel: Search filters and options (fixed width, smaller)
        options_frame = ttk.LabelFrame(main_frame, text="Search Options", padding="5")
        options_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        options_frame.columnconfigure(0, minsize=180)  # Fixed width of ~180 pixels
        
        # Filter checkboxes
        self.in_stock_var = tk.BooleanVar(value=True)
        self.active_only_var = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(options_frame, text="In Stock Only", 
                       variable=self.in_stock_var).grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Checkbutton(options_frame, text="Active Parts Only", 
                       variable=self.active_only_var).grid(row=1, column=0, sticky=tk.W, pady=5)
        
        # Two search buttons
        ttk.Button(options_frame, text="Get Selected Parts", 
                  command=self.search_selected_parts).grid(row=2, column=0, pady=10, sticky=(tk.W, tk.E))
        ttk.Button(options_frame, text="Get Whole BOM", 
                  command=self.search_whole_bom).grid(row=3, column=0, pady=5, sticky=(tk.W, tk.E))
        
        # Separator
        ttk.Separator(options_frame, orient=tk.HORIZONTAL).grid(row=4, column=0, sticky=(tk.W, tk.E), pady=10)
        
        # Export buttons
        ttk.Button(options_frame, text="Preview BOM", 
                  command=self.preview_bom).grid(row=5, column=0, pady=5, sticky=(tk.W, tk.E))
        ttk.Button(options_frame, text="Export BOM", 
                  command=self.export_bom).grid(row=6, column=0, pady=5, sticky=(tk.W, tk.E))
        
        options_frame.columnconfigure(0, weight=1)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready - Open a BOM file to begin")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
    
    def open_bom(self):
        """Open and parse BOM file."""
        file_path = filedialog.askopenfilename(
            title="Select BOM File",
            filetypes=[("BOM files", "*.xlsx *.xls *.csv"), ("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if not file_path:
            logger.info("BOM open cancelled by user")
            return
            
        logger.info(f"Opening BOM file: {file_path}")
        
        try:
            self.status_var.set("Parsing BOM file...")
            self.root.update()
            
            # Parse BOM - returns (components, column_mapping)
            self.components, self.column_mapping = self.bom_parser.parse(file_path)
            logger.info(f"Parsed {len(self.components)} raw components")
            logger.info(f"Column mapping: {self.column_mapping}")
            
            self.consolidated_parts = self.bom_parser.get_consolidated_parts(self.components)
            logger.info(f"Consolidated into {len(self.consolidated_parts)} unique parts")
            
            # Clear previous data
            self.selected_parts = {}
            self.current_search_results = {}
            self.current_search_index = {}
            
            # Collect ALL columns from ALL parts (both raw components and consolidated)
            # This ensures we capture every field that might be used for keyword generation
            all_columns = set()
            
            # Get columns from raw components (original BOM data)
            for component in self.components:
                all_columns.update(component.keys())
            
            # Get columns from consolidated parts
            for part in self.consolidated_parts:
                all_columns.update(part.keys())
            
            # Only filter out truly internal/technical keys
            # Keep everything else, including all BOM fields
            internal_keys = {'refdes_list'}  # Only exclude internal consolidation keys
            
            # Preserve original column order from BOM file instead of alphabetical sorting
            # Define preferred order for standard columns
            preferred_order = ['refdes', 'quantity', 'value', 'package', 'mpn', 'manufacturer', 
                             'description', 'voltage', 'power', 'tolerance', 'footprint', 'datasheet']
            
            # Get columns in original file order from column_mapping
            ordered_cols = []
            seen_cols = set()
            
            # First, add preferred columns in order (if they exist)
            for pref_col in preferred_order:
                if pref_col in all_columns and pref_col not in internal_keys:
                    ordered_cols.append(pref_col)
                    seen_cols.add(pref_col)
            
            # Then add remaining columns in original file order (from column_mapping keys)
            for norm_col in self.column_mapping.keys():
                if norm_col in all_columns and norm_col not in seen_cols and norm_col not in internal_keys:
                    ordered_cols.append(norm_col)
                    seen_cols.add(norm_col)
            
            # Finally, add any columns not in mapping (shouldn't happen, but safety)
            for col in all_columns:
                if col not in seen_cols and col not in internal_keys:
                    ordered_cols.append(col)
            
            display_cols = ordered_cols
            logger.info(f"Found {len(display_cols)} normalized columns in order: {', '.join(display_cols)}")
            
            # Add checkbox as first column (but don't include it in the columns list)
            # Treeview columns are data columns only, checkbox will be in #0
            self.parts_tree['columns'] = display_cols
            
            # Configure the tree column (#0) for checkbox
            self.parts_tree.heading('#0', text='✓')
            self.parts_tree.column('#0', width=30, minwidth=30, stretch=False, anchor=tk.CENTER)
            
            # Clear existing headings and columns for data columns
            for col in display_cols:
                self.parts_tree.heading(col, text='')
                self.parts_tree.column(col, width=0)
            
            # Set up all data columns using ORIGINAL column names from the BOM file
            for col in display_cols:
                # Use original column name from mapping, or fallback to formatted normalized name
                original_name = self.column_mapping.get(col, col.replace('_', ' ').title())
                self.parts_tree.heading(col, text=original_name)
                # Set reasonable column widths based on original name length
                col_width = max(100, len(original_name) * 8)
                self.parts_tree.column(col, width=col_width, minwidth=50, stretch=False)
            
            # Store mapping of item_id to part_key for editing
            self.item_to_part_key = {}
            
            # Clear and populate parts table with ALL data
            self.parts_tree.delete(*self.parts_tree.get_children())
            for idx, part in enumerate(self.consolidated_parts):
                # Create part key using index
                part_key = self._generate_part_key(idx)
                
                # Initialize checkbox state if not exists
                if part_key not in self.part_selected:
                    self.part_selected[part_key] = False
                
                # Get value for each column, preserving all data
                values = []
                for col in display_cols:
                    value = part.get(col, '')
                    # Convert lists/dicts to string representation if needed
                    if isinstance(value, (list, dict)):
                        value = str(value)
                    # Convert None to empty string
                    elif value is None:
                        value = ''
                    values.append(str(value))
                
                # Insert row with checkbox in #0 column and store mapping
                checkbox_text = '✓' if self.part_selected.get(part_key, False) else ''
                item_id = self.parts_tree.insert('', tk.END, text=checkbox_text, values=values)
                self.item_to_part_key[item_id] = part_key
                
                # Set row color if selected (check if it's N/A)
                if self.part_selected.get(part_key, False):
                    # Check if this part has N/A selection
                    selected_part = self.selected_parts.get(part_key, {})
                    is_na = selected_part.get('mpn') == 'NA'
                    if is_na:
                        self.parts_tree.item(item_id, tags=('na_selected',))
                    else:
                        self.parts_tree.item(item_id, tags=('selected',))
                else:
                    self.parts_tree.item(item_id, tags=())
            
            # Configure tag colors
            self.parts_tree.tag_configure('selected', background='lightgreen')
            self.parts_tree.tag_configure('na_selected', background='lightgray')
            
            self.status_var.set(f"Loaded {len(self.consolidated_parts)} unique parts from BOM")
            
            # Build part_key to index mapping
            self.part_key_to_index = {}
            for idx, part in enumerate(self.consolidated_parts):
                part_key = self._generate_part_key(idx)
                self.part_key_to_index[part_key] = idx
            
        except Exception as e:
            logger.error(f"Error parsing BOM file: {e}", exc_info=True)
            messagebox.showerror("Error", f"Failed to load BOM file:\n{e}")
            self.status_var.set("Error loading BOM file")
    
    def populate_parts_table(self):
        """Populate the parts table with consolidated_parts data. Reusable for both open_bom and load_bom_state."""
        if not self.consolidated_parts:
            return
        
        # Collect ALL columns from ALL parts
        all_columns = set()
        
        # Get columns from consolidated parts
        for part in self.consolidated_parts:
            all_columns.update(part.keys())
        
        # Only filter out truly internal/technical keys
        internal_keys = {'refdes_list'}  # Only exclude internal consolidation keys
        
        # Preserve original column order from BOM file instead of alphabetical sorting
        # Define preferred order for standard columns
        preferred_order = ['refdes', 'quantity', 'value', 'package', 'mpn', 'manufacturer', 
                         'description', 'voltage', 'power', 'tolerance', 'footprint', 'datasheet']
        
        # Get columns in original file order from column_mapping
        ordered_cols = []
        seen_cols = set()
        
        # First, add preferred columns in order (if they exist)
        for pref_col in preferred_order:
            if pref_col in all_columns and pref_col not in internal_keys:
                ordered_cols.append(pref_col)
                seen_cols.add(pref_col)
        
        # Then add remaining columns in original file order (from column_mapping keys)
        for norm_col in self.column_mapping.keys():
            if norm_col in all_columns and norm_col not in seen_cols and norm_col not in internal_keys:
                ordered_cols.append(norm_col)
                seen_cols.add(norm_col)
        
        # Finally, add any columns not in mapping (shouldn't happen, but safety)
        for col in all_columns:
            if col not in seen_cols and col not in internal_keys:
                ordered_cols.append(col)
        
        display_cols = ordered_cols
        logger.info(f"Found {len(display_cols)} normalized columns in order: {', '.join(display_cols)}")
        
        # Add checkbox as first column (but don't include it in the columns list)
        # Treeview columns are data columns only, checkbox will be in #0
        self.parts_tree['columns'] = display_cols
        
        # Configure the tree column (#0) for checkbox
        self.parts_tree.heading('#0', text='✓')
        self.parts_tree.column('#0', width=30, minwidth=30, stretch=False, anchor=tk.CENTER)
        
        # Clear existing headings and columns for data columns
        for col in display_cols:
            self.parts_tree.heading(col, text='')
            self.parts_tree.column(col, width=0)
        
        # Set up all data columns using ORIGINAL column names from the BOM file
        for col in display_cols:
            # Use original column name from mapping, or fallback to formatted normalized name
            original_name = self.column_mapping.get(col, col.replace('_', ' ').title())
            self.parts_tree.heading(col, text=original_name)
            # Set reasonable column widths based on original name length
            col_width = max(100, len(original_name) * 8)
            self.parts_tree.column(col, width=col_width, minwidth=50, stretch=False)
        
        # Store mapping of item_id to part_key for editing
        self.item_to_part_key = {}
        
        # Clear and populate parts table with ALL data
        self.parts_tree.delete(*self.parts_tree.get_children())
        for idx, part in enumerate(self.consolidated_parts):
            # Create part key using index
            part_key = self._generate_part_key(idx)
            
            # Initialize checkbox state if not exists
            if part_key not in self.part_selected:
                self.part_selected[part_key] = False
            
            # Get value for each column, preserving all data
            values = []
            for col in display_cols:
                value = part.get(col, '')
                # Convert lists/dicts to string representation if needed
                if isinstance(value, (list, dict)):
                    value = str(value)
                # Convert None to empty string
                elif value is None:
                    value = ''
                values.append(str(value))
            
            # Insert row with checkbox in #0 column and store mapping
            checkbox_text = '✓' if self.part_selected.get(part_key, False) else ''
            item_id = self.parts_tree.insert('', tk.END, text=checkbox_text, values=values)
            self.item_to_part_key[item_id] = part_key
            
            # Set row color if selected (check if it's N/A)
            if self.part_selected.get(part_key, False):
                # Check if this part has N/A selection
                selected_part = self.selected_parts.get(part_key, {})
                is_na = selected_part.get('mpn') == 'NA'
                if is_na:
                    self.parts_tree.item(item_id, tags=('na_selected',))
                else:
                    self.parts_tree.item(item_id, tags=('selected',))
            else:
                self.parts_tree.item(item_id, tags=())
        
        # Configure tag colors
        self.parts_tree.tag_configure('selected', background='lightgreen')
        self.parts_tree.tag_configure('na_selected', background='lightgray')
        
        # Build part_key to index mapping
        self.part_key_to_index = {}
        for idx, part in enumerate(self.consolidated_parts):
            part_key = self._generate_part_key(idx)
            self.part_key_to_index[part_key] = idx
        
        self.status_var.set(f"Loaded {len(self.consolidated_parts)} unique parts from BOM")
    
    def _generate_part_key(self, index: int) -> str:
        """Generate an index-based part key."""
        return f"part_{index}"
    
    def _get_part_by_key(self, part_key: str) -> Optional[Dict[str, Any]]:
        """Get part data from part_key using index lookup."""
        index = self.part_key_to_index.get(part_key)
        if index is not None and index < len(self.consolidated_parts):
            return self.consolidated_parts[index]
        return None
    
    def get_column_value(self, item_values: List[Any], col_name: str) -> str:
        """Helper to get value from treeview item values based on column name."""
        columns = self.parts_tree['columns']
        try:
            idx = list(columns).index(col_name)
            if idx < len(item_values):
                return str(item_values[idx])
        except ValueError:
            pass
        return ""

    def on_cell_click(self, event):
        """Handle cell click - checkbox column toggles checkbox, other columns start editing on single click."""
        region = self.parts_tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.parts_tree.identify_column(event.x)
            item = self.parts_tree.identify_row(event.y)
            
            if item and column:
                # Get column index (column is like '#0', '#1', '#2', etc.)
                col_index = int(column.replace('#', ''))
                columns = list(self.parts_tree['columns'])
                
                if col_index == 0:  # Checkbox column (#0)
                    part_key = self.item_to_part_key.get(item)
                    if part_key:
                        # Toggle checkbox only - don't do anything else
                        current_state = self.part_selected.get(part_key, False)
                        self.part_selected[part_key] = not current_state
                        self.update_row_checkbox(item, part_key)
                # Don't start editing on single click - only on double click
                # (Double-click handler in on_cell_double_click will handle editing)
    
    def on_cell_double_click(self, event):
        """Handle double-click to start editing a cell."""
        region = self.parts_tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.parts_tree.identify_column(event.x)
            item = self.parts_tree.identify_row(event.y)
            
            if item and column:
                # identify_column returns '#0' for tree column, '#1', '#2', etc. for data columns
                col_index = int(column.replace('#', ''))
                columns = list(self.parts_tree['columns'])
                
                if col_index == 0:  # Don't edit checkbox column (#0)
                    return
                
                # Data columns start at index 1, so subtract 1 to get data column index
                # col_index 1 -> data_col_index 0 (first data column)
                # col_index 2 -> data_col_index 1 (second data column)
                data_col_index = col_index - 1
                if data_col_index >= 0 and data_col_index < len(columns):
                    col_name = columns[data_col_index]
                    self.start_cell_edit(item, col_name)
    
    def start_cell_edit(self, item_id, col_name):
        """Start editing a cell."""
        # Cancel any existing edit
        if self.editing_cell:
            self.cancel_cell_edit()
        
        if not col_name or col_name == 'selected':  # Don't edit checkbox
            return
        
        # Get current value
        current_value = self.parts_tree.set(item_id, col_name)
        
        # Get cell bbox
        bbox = self.parts_tree.bbox(item_id, col_name)
        if not bbox:
            return
        
        # Create entry widget for editing
        self.editing_cell = (item_id, col_name)
        entry = tk.Entry(self.parts_tree, width=bbox[2])
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus()
        
        def save_edit(event=None):
            new_value = entry.get()
            self.save_cell_value(item_id, col_name, new_value)
            entry.destroy()
            self.editing_cell = None
        
        def cancel_edit(event=None):
            entry.destroy()
            self.editing_cell = None
        
        entry.bind('<Return>', save_edit)
        entry.bind('<FocusOut>', save_edit)
        entry.bind('<Escape>', cancel_edit)
        
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
    
    def cancel_cell_edit(self):
        """Cancel current cell edit."""
        if self.editing_cell:
            self.editing_cell = None
    
    def save_cell_value(self, item_id, col_name, value):
        """Save edited cell value back to data structure."""
        part_key = self.item_to_part_key.get(item_id)
        if not part_key:
            return
        
        # Get the part using index lookup
        part = self._get_part_by_key(part_key)
        if part:
            part[col_name] = value
            # Update treeview display
            self.parts_tree.set(item_id, col_name, value)
            logger.debug(f"Updated {col_name} for {part_key} to {value}")
    
    def update_row_checkbox(self, item_id, part_key, is_na: bool = False):
        """Update checkbox display and row color."""
        is_selected = self.part_selected.get(part_key, False)
        logger.debug(f"Updating row checkbox for {part_key}: selected={is_selected}, is_na={is_na}, item_id={item_id}")
        
        # Verify item_id exists
        try:
            # Check if item exists in tree
            if not self.parts_tree.exists(item_id):
                logger.warning(f"Item {item_id} does not exist in tree for part_key {part_key}")
                return
            
            if is_selected:
                self.parts_tree.item(item_id, text='✓')
                if is_na:
                    self.parts_tree.item(item_id, tags=('na_selected',))
                    logger.debug(f"Set item {item_id} to checked with grey background (N/A)")
                else:
                    self.parts_tree.item(item_id, tags=('selected',))
                    logger.debug(f"Set item {item_id} to checked with green background")
            else:
                self.parts_tree.item(item_id, text='', tags=())
                logger.debug(f"Set item {item_id} to unchecked")
            
            # Force UI update
            self.root.update_idletasks()
        except Exception as e:
            logger.error(f"Error updating row checkbox: {e}", exc_info=True)
    
    def get_selected_part_key(self) -> Optional[str]:
        """Get the key for the currently selected part in the table."""
        selection = self.parts_tree.selection()
        if not selection:
            return None
        
        item_id = selection[0]
        return self.item_to_part_key.get(item_id)
    
    def get_selected_part_data(self) -> Optional[Dict[str, Any]]:
        """Get the full data for the currently selected part."""
        part_key = self.get_selected_part_key()
        if not part_key:
            return None
        
        # Get part using index lookup
        return self._get_part_by_key(part_key)
    
    def on_part_selected(self, event):
        """Handle part selection in the table. Now a no-op to prevent clearing search results."""
        # This method is kept for backward compatibility but no longer clears results
        # Row selection changes should not affect search results or trigger searches
        pass
    
    def clear_results(self):
        """Clear the results panel."""
        for widget in self.results_frame.winfo_children():
            widget.destroy()
        self.more_parts_btn.config(state=tk.DISABLED)
        if hasattr(self, 'confirm_part_btn'):
            self.confirm_part_btn.config(state=tk.DISABLED)
        # Note: We keep radio_vars to track selections across "Get More Parts" calls
    
    def generate_search_term(self, component: Dict[str, Any]) -> str:
        """Generate a concise search term from component data using Gemini.
        
        Uses ALL fields from the component to ensure nothing is omitted.
        """
        try:
            # Build context for Gemini using ALL component fields
            # This ensures all BOM columns are available for keyword generation
            desc_parts = []
            footprint_package = None
            
            for key, val in component.items():
                # Skip empty values and internal keys
                if val and key != 'refdes_list':
                    # Format the key nicely
                    key_display = key.replace('_', ' ').title()
                    
                    # Special handling for footprint/package fields - extract package size
                    if key in ['footprint', 'package'] and isinstance(val, str):
                        # Look for package size patterns like 0402, 0603, 0805, etc.
                        import re
                        # Pattern to match 4-digit package codes (0402, 0603, 0805, 1206, etc.)
                        package_match = re.search(r'(\d{4})', val)
                        if package_match:
                            footprint_package = package_match.group(1)
                            desc_parts.append(f"{key_display}: {val} [PACKAGE_SIZE: {footprint_package}]")
                        else:
                            desc_parts.append(f"{key_display}: {val}")
                    else:
                        desc_parts.append(f"{key_display}: {val}")
            
            context = ", ".join(desc_parts)
            
            # Add explicit package size instruction if found
            package_instruction = ""
            if footprint_package:
                package_instruction = f"\nIMPORTANT: The package size is {footprint_package} (extracted from footprint/package field). Use this exact package size in your search phrase."
            
            prompt = f"""Create a concise search phrase for Mouser for this component.
Component Data: {context}{package_instruction}

Rules:
1. Extract the package size from footprint/package fields. Look for 4-digit codes like 0402, 0603, 0805, 1206, etc.
   - In "Resistor_SMD:R_0603_1608Metric", the package is "0603"
   - In "C_0402_1005Metric", the package is "0402"
   - If you see [PACKAGE_SIZE: XXXX] in the data, use that exact value
   - For ICs, the package size might be BGA or QFN or SOT etc. 
2. Use the component value (e.g., "0.1uF", "10k", "100ohm", "499k")
3. Identify component type from description or reference designator (e.g., "capacitor", "resistor", "inductor")
4. Format as: "[value] [type] [package]" (e.g., "0.1uF capacitor 0402" or "10k resistor 0603" or "499k resistor 0603")
5. If there is a manufacturer part number, start with that and add the package
6. Keep it under 40 characters - be concise!
7. Ignore library paths, symbols, and other non-essential text
8. Return ONLY the search phrase, no quotes or extra text.

Examples:
- Value: "0.1uF", Footprint: "Capacitor_SMD:C_0402_1005Metric" -> "0.1uF capacitor 0402"
- Value: "10k", Footprint: "Resistor_SMD:R_0603_1608Metric" -> "10k resistor 0603"
- Value: "499k", Footprint: "Resistor_SMD:R_0603_1608Metric" -> "499k resistor 0603"
"""
            response = self.gemini_model.generate_content(prompt)
            term = response.text.strip().strip('"').strip("'")
            logger.info(f"Gemini generated search term: '{term}' from '{context}'")
            return term
            
        except Exception as e:
            logger.error(f"Gemini generation failed: {e}")
            # Fallback
            return f"{component.get('value', '')} {component.get('package', '')} {component.get('description', '')}"

    def get_selected_parts(self) -> List[Dict[str, Any]]:
        """Get all parts that are currently selected (highlighted) in the BOM table."""
        selected_parts = []
        
        # Get selected item IDs from the treeview
        selected_items = self.parts_tree.selection()
        if not selected_items:
            return selected_parts
        
        # Map item IDs to part_keys and then to part data
        selected_part_keys = set()
        for item_id in selected_items:
            part_key = self.item_to_part_key.get(item_id)
            if part_key:
                selected_part_keys.add(part_key)
        
        # Get the actual part data for selected part_keys using lookup
        for part_key in selected_part_keys:
            part = self._get_part_by_key(part_key)
            if part:
                selected_parts.append(part)
        
        return selected_parts
    
    def search_selected_parts(self):
        """Search Mouser for all selected parts (those with checked checkboxes) using batch keyword generation."""
        if not self.mouser_api:
            logger.warning("Attempted search without Mouser API configuration")
            messagebox.showerror("Error", "Mouser API not configured")
            return
        
        if not self.consolidated_parts:
            messagebox.showwarning("No BOM", "Please load a BOM file first")
            return
        
        # Get selected parts
        selected_parts = self.get_selected_parts()
        if not selected_parts:
            logger.warning("Attempted search without any selected parts")
            messagebox.showwarning("No Selection", "Please select one or more rows in the BOM table")
            return
        
        logger.info(f"Starting batch Mouser search for {len(selected_parts)} selected parts")
        
        # Clear previous results
        self.clear_results()
        
        # Show loading
        loading_label = ttk.Label(self.results_frame, text=f"Generating keywords and searching Mouser for {len(selected_parts)} selected parts...")
        loading_label.pack(pady=20)
        self.root.update()
        
        # Run search in thread to avoid blocking UI
        def do_batch_search():
            try:
                # Prepare components with part_key
                # Find indices for selected parts in consolidated_parts
                components_with_keys = []
                for part in selected_parts:
                    # Find the index of this part in consolidated_parts
                    try:
                        idx = self.consolidated_parts.index(part)
                        part_key = self._generate_part_key(idx)
                        part_copy = part.copy()
                        part_copy['_part_key'] = part_key
                        components_with_keys.append(part_copy)
                    except ValueError:
                        logger.warning(f"Could not find part in consolidated_parts: {part.get('refdes', 'Unknown')}")
                        continue
                
                # Batch generate keywords
                logger.info("Batch generating keywords with Gemini...")
                keywords_dict = self.batch_generate_keywords(components_with_keys)
                logger.info(f"Generated {len(keywords_dict)} keywords")
                
                # Search Mouser for each selected part
                in_stock_only = self.in_stock_var.get()
                active_only = self.active_only_var.get()
                
                all_results = {}
                for part in selected_parts:
                    # Find the index of this part in consolidated_parts
                    try:
                        idx = self.consolidated_parts.index(part)
                        part_key = self._generate_part_key(idx)
                    except ValueError:
                        logger.warning(f"Could not find part in consolidated_parts: {part.get('refdes', 'Unknown')}")
                        continue
                    
                    search_term = keywords_dict.get(part_key, '')
                    if not search_term:
                        # Fallback if keyword not generated
                        search_term = f"{part.get('value', '')} {part.get('package', '')} {part.get('description', '')}"
                    
                    # Rate limiting: wait 1 second after every 10 calls
                    result_count = len(all_results)
                    if result_count > 0 and result_count % 10 == 0:
                        logger.info(f"Rate limiting: waiting 1 second after {result_count} API calls")
                        time.sleep(1.0)
                    
                    # Log with readable refdes for clarity
                    refdes_display = part.get('refdes', 'Unknown')[:50] + ('...' if len(part.get('refdes', '')) > 50 else '')
                    logger.info(f"Searching Mouser for {refdes_display} (key: {part_key}): {search_term}")
                    # Store the keyword used for this search (to show in custom search)
                    self.last_search_keywords[part_key] = search_term
                    spec = {'keyword': search_term}
                    results = self.mouser_api.search(part, spec, in_stock_only, active_only)
                    
                    # Rank results with current sort preference
                    target_package = part.get('package', '')
                    sort_by = self.sort_preference.get()
                    ranked_results = self.rank_parts_with_preference(results, target_package, sort_by)
                    
                    all_results[part_key] = ranked_results
                    self.current_search_results[part_key] = ranked_results
                    self.current_search_index[part_key] = 0
                
                # Update UI in main thread - show batch results with navigation
                self.root.after(0, lambda: self.display_batch_results(all_results))
                
            except Exception as e:
                error_msg = str(e)
                logger.error(f"Batch search thread error: {error_msg}", exc_info=True)
                self.root.after(0, lambda: self.show_search_error(error_msg))
        
        thread = threading.Thread(target=do_batch_search, daemon=True)
        thread.start()
    
    def batch_generate_keywords(self, components: List[Dict[str, Any]]) -> Dict[str, str]:
        """
        Generate search keywords for all components in a single Gemini API call.
        
        Args:
            components: List of component dictionaries with part_key stored in component
            
        Returns:
            Dictionary mapping part_key to search_term
        """
        if not self.gemini_model:
            # Fallback: generate individually
            logger.warning("Gemini not available, generating keywords individually")
            keywords = {}
            for comp in components:
                part_key = comp.get('_part_key', '')
                if part_key:
                    keywords[part_key] = self.generate_search_term(comp)
            return keywords
        
        try:
            # Build prompt with all components
            component_data = []
            for idx, comp in enumerate(components):
                part_key = comp.get('_part_key', f'part_{idx}')
                desc_parts = []
                footprint_package = None
                
                for key, val in comp.items():
                    if key != '_part_key' and key != 'refdes_list' and val:
                        key_display = key.replace('_', ' ').title()
                        
                        # Special handling for footprint/package fields - extract package size
                        if key in ['footprint', 'package'] and isinstance(val, str):
                            # Look for package size patterns like 0402, 0603, 0805, etc.
                            import re
                            # Pattern to match 4-digit package codes (0402, 0603, 0805, 1206, etc.)
                            package_match = re.search(r'(\d{4})', val)
                            if package_match:
                                footprint_package = package_match.group(1)
                                desc_parts.append(f"{key_display}: {val} [PACKAGE_SIZE: {footprint_package}]")
                            else:
                                desc_parts.append(f"{key_display}: {val}")
                        else:
                            desc_parts.append(f"{key_display}: {val}")
                
                context = ", ".join(desc_parts)
                package_note = f" [PACKAGE_SIZE: {footprint_package}]" if footprint_package else ""
                component_data.append(f"Component {idx + 1} (Key: {part_key}): {context}{package_note}")
            
            all_components_text = "\n\n".join(component_data)
            
            prompt = f"""Create concise search phrases for Mouser for each of these components.
Return the results as a JSON object where each key is the component key and the value is the search phrase.

Component Data:
{all_components_text}

Rules for each search phrase:
1. Extract the package size from footprint/package fields. Look for 4-digit codes like 0402, 0603, 0805, 1206, etc.
   - In "Resistor_SMD:R_0603_1608Metric", the package is "0603"
   - In "C_0402_1005Metric", the package is "0402"
   - If you see [PACKAGE_SIZE: XXXX] in the data, use that exact value
2. Use the component value (e.g., "0.1uF", "10k", "100ohm", "499k")
3. Identify component type from description or reference designator (e.g., "capacitor", "resistor", "inductor")
4. Format as: "[value] [type] [package]" (e.g., "0.1uF capacitor 0402" or "10k resistor 0603" or "499k resistor 0603")
5. If there is a manufacturer part number, start with that and add the package
6. Keep each phrase under 40 characters - be concise!
7. Ignore library paths, symbols, and other non-essential text
8. Return ONLY a JSON object in this format: {{"key1": "search phrase 1", "key2": "search phrase 2", ...}}
9. No quotes around the JSON object, no extra text.

Examples:
- Value: "0.1uF", Footprint: "Capacitor_SMD:C_0402_1005Metric" -> "0.1uF capacitor 0402"
- Value: "10k", Footprint: "Resistor_SMD:R_0603_1608Metric" -> "10k resistor 0603"
- Value: "499k", Footprint: "Resistor_SMD:R_0603_1608Metric" -> "499k resistor 0603"
"""
            logger.debug(f"Gemini Prompt: {prompt}")
            response = self.gemini_model.generate_content(prompt)
            response_text = response.text.strip()
            logger.debug(f"Gemini Response: {response_text}")
            
            # Try to parse JSON response
            # Remove markdown code blocks if present
            if response_text.startswith('```'):
                response_text = response_text.split('```')[1]
                if response_text.startswith('json'):
                    response_text = response_text[4:]
                response_text = response_text.strip()
            
            keywords = json.loads(response_text)
            return keywords
            
        except Exception as e:
            logger.error(f"Batch Gemini generation failed: {e}, falling back to individual generation")
            # Fallback to individual generation
            keywords = {}
            for comp in components:
                part_key = comp.get('_part_key', '')
                if part_key:
                    try:
                        keywords[part_key] = self.generate_search_term(comp)
                    except Exception as e2:
                        logger.error(f"Failed to generate keyword for {part_key}: {e2}")
                        keywords[part_key] = f"{comp.get('value', '')} {comp.get('package', '')}"
            return keywords
    
    def search_whole_bom(self):
        """Search Mouser for all parts in the BOM using batch keyword generation."""
        if not self.mouser_api:
            logger.warning("Attempted search without Mouser API configuration")
            messagebox.showerror("Error", "Mouser API not configured")
            return
        
        if not self.consolidated_parts:
            messagebox.showwarning("No BOM", "Please load a BOM file first")
            return
        
        logger.info(f"Starting batch Mouser search for {len(self.consolidated_parts)} parts")
        
        # Clear previous results
        self.clear_results()
        
        # Show loading
        loading_label = ttk.Label(self.results_frame, text=f"Generating keywords and searching Mouser for {len(self.consolidated_parts)} parts...")
        loading_label.pack(pady=20)
        self.root.update()
        
        # Run search in thread to avoid blocking UI
        def do_batch_search():
            try:
                # Prepare components with part_key using index
                components_with_keys = []
                for idx, part in enumerate(self.consolidated_parts):
                    part_key = self._generate_part_key(idx)
                    part_copy = part.copy()
                    part_copy['_part_key'] = part_key
                    components_with_keys.append(part_copy)
                
                # Batch generate keywords
                logger.info("Batch generating keywords with Gemini...")
                keywords_dict = self.batch_generate_keywords(components_with_keys)
                logger.info(f"Generated {len(keywords_dict)} keywords")
                
                # Search Mouser for each part
                in_stock_only = self.in_stock_var.get()
                active_only = self.active_only_var.get()
                
                all_results = {}
                for idx, part in enumerate(self.consolidated_parts):
                    part_key = self._generate_part_key(idx)
                    
                    search_term = keywords_dict.get(part_key, '')
                    if not search_term:
                        # Fallback if keyword not generated
                        search_term = f"{part.get('value', '')} {part.get('package', '')} {part.get('description', '')}"
                    
                    # Rate limiting: wait 1 second after every 10 calls
                    if idx > 0 and idx % 10 == 0:
                        logger.info(f"Rate limiting: waiting 1 second after {idx} API calls")
                        time.sleep(1.0)
                    
                    # Log with readable refdes for clarity
                    refdes_display = part.get('refdes', 'Unknown')[:50] + ('...' if len(part.get('refdes', '')) > 50 else '')
                    logger.info(f"Searching Mouser for {refdes_display} (key: {part_key}): {search_term}")
                    # Store the keyword used for this search (to show in custom search)
                    self.last_search_keywords[part_key] = search_term
                    spec = {'keyword': search_term}
                    results = self.mouser_api.search(part, spec, in_stock_only, active_only)
                    
                    # Rank results with current sort preference
                    target_package = part.get('package', '')
                    sort_by = self.sort_preference.get()
                    ranked_results = self.rank_parts_with_preference(results, target_package, sort_by)
                    
                    all_results[part_key] = ranked_results
                    self.current_search_results[part_key] = ranked_results
                    self.current_search_index[part_key] = 0
                
                # Update UI in main thread - show summary
                self.root.after(0, lambda: self.display_batch_results(all_results))
                
            except Exception as e:
                error_msg = str(e)
                logger.error(f"Batch search thread error: {error_msg}", exc_info=True)
                self.root.after(0, lambda: self.show_search_error(error_msg))
        
        thread = threading.Thread(target=do_batch_search, daemon=True)
        thread.start()
    
    def display_batch_results(self, all_results: Dict[str, List[Dict[str, Any]]]):
        """Display batch search results one at a time with navigation."""
        # Store batch results and initialize navigation
        self.batch_results = all_results
        self.batch_part_keys = list(all_results.keys())
        self.current_batch_index = 0
        
        # Display first part
        if self.batch_part_keys:
            self.display_single_batch_result()
        else:
            self.clear_results()
            no_results = ttk.Label(self.results_frame, text="No parts found in batch search")
            no_results.pack(pady=20)
        
        self.status_var.set(f"Batch search completed: {len(all_results)} parts - Review each part")
    
    def display_single_batch_result(self):
        """Display results for a single part in batch mode."""
        if not self.batch_part_keys or self.current_batch_index >= len(self.batch_part_keys):
            return
        
        part_key = self.batch_part_keys[self.current_batch_index]
        results = self.batch_results.get(part_key, [])
        total_count = len(self.batch_part_keys)
        
        # Display with navigation
        self.display_results(part_key, results, show_navigation=True, 
                           current_index=self.current_batch_index, total_count=total_count)
    
    def go_to_previous_part(self):
        """Navigate to previous part in batch review."""
        if self.current_batch_index > 0:
            self.current_batch_index -= 1
            self.display_single_batch_result()
    
    def go_to_next_part(self):
        """Navigate to next part in batch review."""
        if self.current_batch_index < len(self.batch_part_keys) - 1:
            self.current_batch_index += 1
            self.display_single_batch_result()
    
    def confirm_and_advance(self):
        """Confirm current selection and advance to next part."""
        # The confirmation is already done in confirm_selected_part
        # Just advance to next
        if self.current_batch_index < len(self.batch_part_keys) - 1:
            self.current_batch_index += 1
            self.display_single_batch_result()
        else:
            # Last part - show completion message
            messagebox.showinfo("Batch Complete", 
                              f"All {len(self.batch_part_keys)} parts have been reviewed.\n"
                              f"You can export the BOM or review parts using Previous/Next buttons.")
    
    def _extract_price(self, part: Dict[str, Any]) -> float:
        """Extract numeric price value from part for sorting. Uses unit price (quantity=1). Returns inf if no price."""
        price_breaks = part.get('price_breaks', [])
        if not price_breaks:
            logger.debug(f"Part {part.get('mpn', 'Unknown')}: No price_breaks found")
            return float('inf')
        
        # Price breaks are typically ordered by quantity (1, 10, 100, etc.)
        # For unit price sorting, we want the price for quantity=1 (unit price)
        # Price breaks usually start with quantity=1, which has the highest unit price
        
        unit_price = None
        first_price = None
        
        for pb in price_breaks:
            try:
                quantity = pb.get('quantity', 0)
                price = pb.get('price', '')
                
                if not price:
                    continue
                
                # Handle both string and numeric prices
                if isinstance(price, str):
                    price_str = price.replace('$', '').replace(',', '').strip()
                    price_float = float(price_str)
                else:
                    price_float = float(price)
                
                # Save first price as fallback (usually quantity=1)
                if first_price is None:
                    first_price = price_float
                
                # Prefer price at quantity=1 (unit price for single unit purchase)
                if quantity == 1:
                    unit_price = price_float
                    break  # Found unit price, we're done
                    
            except (ValueError, AttributeError, TypeError) as e:
                logger.debug(f"Failed to parse price break {pb}: {e}")
                continue
        
        # Use unit price if found, otherwise use first price break (which is usually qty=1 anyway)
        result = unit_price if unit_price is not None else first_price if first_price is not None else float('inf')
        
        mpn = part.get('mpn', 'Unknown')
        if result == float('inf'):
            logger.warning(f"Part {mpn}: Could not extract price from {len(price_breaks)} price breaks")
        else:
            logger.debug(f"Part {mpn}: Extracted unit price = ${result:.2f} from {len(price_breaks)} price breaks")
        return result
    
    def rank_parts_with_preference(self, results: List[Dict[str, Any]], 
                                   target_package: Optional[str], 
                                   sort_by: str = 'stock') -> List[Dict[str, Any]]:
        """
        Rank parts with adjustable weights based on sort preference.
        
        Args:
            results: List of part dictionaries from Mouser
            target_package: Target package/footprint for matching
            sort_by: 'stock' or 'price' to prioritize
            
        Returns:
            Sorted list of parts with 'score' field added
        """
        if sort_by == 'price':
            # Direct price sort: cheapest first
            # First, calculate scores for all parts (for display purposes)
            for part in results:
                part['score'] = self.ranker.calculate_score(part, target_package)
            
            # Extract prices for all parts for debugging
            price_data = []
            for p in results:
                price_val = self._extract_price(p)
                mpn = p.get('mpn', 'Unknown')
                # Get the displayed price string
                price_display = 'N/A'
                if p.get('price_breaks'):
                    first_pb = p['price_breaks'][0]
                    price_display = first_pb.get('price', 'N/A')
                price_data.append((price_val, mpn, price_display))
            
            # Sort directly by price (ascending), then by stock as tiebreaker
            # Create a copy to avoid modifying original list
            results_copy = list(results)
            sorted_parts = sorted(
                results_copy,
                key=lambda p: (self._extract_price(p), -p.get('stock', 0)),
                reverse=False  # Ascending for price (cheapest first)
            )
            
            for i, p in enumerate(sorted_parts[:10]):
                price_val = self._extract_price(p)
                mpn = p.get('mpn', 'Unknown')
                price_display = 'N/A'
                if p.get('price_breaks'):
                    first_pb = p['price_breaks'][0]
                    price_display = first_pb.get('price', 'N/A')
            
            logger.info(f"Sorted {len(sorted_parts)} parts by price (cheapest first)")
            return sorted_parts
        else:
            # Stock sort: use weighted scoring
            # Adjust weights based on sort preference
            stock_weight = 0.6
            price_weight = 0.2
            lifecycle_weight = 0.1
            package_match_weight = 0.1
            
            # Normalize weights
            total = stock_weight + price_weight + lifecycle_weight + package_match_weight
            stock_weight /= total
            price_weight /= total
            lifecycle_weight /= total
            package_match_weight /= total
            
            # Temporarily adjust ranker weights
            original_stock = self.ranker.stock_weight
            original_price = self.ranker.price_weight
            original_lifecycle = self.ranker.lifecycle_weight
            original_package = self.ranker.package_match_weight
            
            self.ranker.stock_weight = stock_weight
            self.ranker.price_weight = price_weight
            self.ranker.lifecycle_weight = lifecycle_weight
            self.ranker.package_match_weight = package_match_weight
            
            # Rank parts
            ranked = self.ranker.rank_parts(results, target_package)
            
            # Restore original weights
            self.ranker.stock_weight = original_stock
            self.ranker.price_weight = original_price
            self.ranker.lifecycle_weight = original_lifecycle
            self.ranker.package_match_weight = original_package
            
            logger.info(f"Ranked {len(ranked)} parts with preference: {sort_by} (stock_weight={stock_weight:.2f}, price_weight={price_weight:.2f})")
            
            return ranked
    
    def _suggest_keyword(self, part: Dict[str, Any]) -> str:
        """Suggest a keyword based on part data for the custom search input."""
        parts = []
        if part.get('value'):
            parts.append(part['value'])
        if part.get('package'):
            # Try to extract just the package size (e.g., 0603 from "Resistor_SMD:R_0603_1608Metric")
            package = part['package']
            match = re.search(r'(\d{4})', package)
            if match:
                parts.append(match.group(1))
            else:
                parts.append(package)
        if part.get('description'):
            # Take first few words from description
            desc_words = part['description'].split()[:3]
            parts.extend(desc_words)
        
        return ' '.join(parts[:5])  # Limit to 5 parts to keep it concise
    
    def search_with_custom_keyword(self, part_key: str, custom_keyword: str):
        """Search Mouser with a custom keyword provided by the user."""
        if not custom_keyword or not custom_keyword.strip():
            messagebox.showwarning("Empty Keyword", "Please enter a search keyword")
            return
        
        keyword = custom_keyword.strip()
        logger.info(f"Searching with custom keyword for part_key {part_key}: '{keyword}'")
        
        # Get the original part data
        part = self._get_part_by_key(part_key)
        if not part:
            logger.error(f"Could not find part data for {part_key}")
            messagebox.showerror("Error", "Could not find part data")
            return
        
        # Clear previous results and show loading
        self.clear_results()
        loading_label = ttk.Label(self.results_frame, text=f"Searching Mouser with keyword: '{keyword}'...")
        loading_label.pack(pady=20)
        self.root.update()
        
        # Run search in thread to avoid blocking UI
        def do_custom_search():
            try:
                # Get search options
                in_stock_only = self.in_stock_var.get()
                active_only = self.active_only_var.get()
                
                # Search with custom keyword and store it
                self.last_search_keywords[part_key] = keyword
                spec = {'keyword': keyword}
                results = self.mouser_api.search(part, spec, in_stock_only, active_only)
                
                # Rank results with current sort preference
                target_package = part.get('package', '')
                sort_by = self.sort_preference.get()
                ranked_results = self.rank_parts_with_preference(results, target_package, sort_by)
                
                # Store results
                self.current_search_results[part_key] = ranked_results
                
                # Update UI with results
                self.root.after(0, lambda: self._display_custom_search_results(part_key, ranked_results))
                
            except Exception as e:
                error_msg = str(e)
                logger.error(f"Custom search failed: {error_msg}")
                self.root.after(0, lambda: self.show_search_error(error_msg))
        
        # Start search thread
        thread = threading.Thread(target=do_custom_search, daemon=True)
        thread.start()
    
    def _display_custom_search_results(self, part_key: str, results: List[Dict[str, Any]]):
        """Display results from a custom keyword search."""
        # Check if we're in batch mode
        if self.batch_part_keys and part_key in self.batch_part_keys:
            # In batch mode - display with navigation
            current_index = self.batch_part_keys.index(part_key)
            total_count = len(self.batch_part_keys)
            self.display_results(part_key, results, show_navigation=True, 
                               current_index=current_index, total_count=total_count)
        else:
            # Single part mode - display first 3 results
            self.display_results(part_key, results[:3], show_navigation=False)
    
    def apply_sort_preference(self, part_key: str, results: List[Dict[str, Any]] = None):
        """Re-rank and re-display results based on sort preference change."""
        sort_by = self.sort_preference.get()
        logger.info(f"DEBUG: Applying sort preference '{sort_by}' for part_key {part_key}")
        
        # Get full results from stored results (not just displayed subset)
        if results is None:
            results = self.current_search_results.get(part_key, [])
            logger.info(f"DEBUG: Got {len(results)} results from current_search_results")
            if not results:
                # Fallback to displayed results if search results not available
                results = self.current_displayed_results.get(part_key, [])
                logger.info(f"DEBUG: Got {len(results)} results from current_displayed_results (fallback)")
        
        if not results:
            logger.warning(f"No results available for {part_key} to re-sort")
            return
        
        # Get the original part data for package matching using index lookup
        part = self._get_part_by_key(part_key)
        if not part:
            logger.warning(f"Could not find part data for {part_key}")
            return
        
        # Re-rank with new preference
        target_package = part.get('package', '')
        logger.info(f"DEBUG: Re-ranking {len(results)} results with sort_by='{sort_by}', target_package='{target_package}'")
        
        # Log prices before sorting for debugging
        if sort_by == 'price':
            logger.info("DEBUG: Prices BEFORE sorting:")
            for i, r in enumerate(results[:10]):
                price_val = self._extract_price(r)
                mpn = r.get('mpn', 'Unknown')
                logger.info(f"  {i+1}. {mpn}: ${price_val:.4f}")
        
        ranked_results = self.rank_parts_with_preference(results, target_package, sort_by)
        
        # Log prices after sorting for debugging
        if sort_by == 'price':
            logger.info("DEBUG: Prices AFTER sorting:")
            for i, r in enumerate(ranked_results[:10]):
                price_val = self._extract_price(r)
                mpn = r.get('mpn', 'Unknown')
                logger.info(f"  {i+1}. {mpn}: ${price_val:.4f}")
        
        # Update stored results
        self.current_search_results[part_key] = ranked_results
        self.current_displayed_results[part_key] = ranked_results
        
        # Check if we're in batch mode
        if self.batch_part_keys and part_key in self.batch_part_keys:
            # In batch mode - re-display current part with all results
            current_index = self.current_batch_index
            total_count = len(self.batch_part_keys)
            self.display_results(part_key, ranked_results, show_navigation=True, 
                               current_index=current_index, total_count=total_count)
        else:
            # Single part mode - re-display with first 3 results
            self.display_results(part_key, ranked_results[:3], show_navigation=False)
    
    def display_results(self, part_key: str, results: List[Dict[str, Any]], 
                       show_navigation: bool = False, current_index: int = 0, total_count: int = 0):
        """Display search results in the results panel with radio buttons."""
        self.clear_results()
        
        # Store results for re-sorting
        self.current_displayed_results[part_key] = results
        
        # Initialize radio button variable for this part_key (XOR selection)
        if part_key not in self.radio_vars:
            self.radio_vars[part_key] = tk.StringVar(value='')
        radio_var = self.radio_vars[part_key]
        
        # Clear result frames mapping for this part_key
        self.result_frames = {k: v for k, v in self.result_frames.items() if k[0] != part_key}
        
        # Get the original BOM part data to display at top
        bom_part = self._get_part_by_key(part_key)
        
        # Show BOM part info header at top (concatenated row values)
        if bom_part:
            # Build concatenated info string from key fields
            info_parts = []
            refdes = bom_part.get('refdes', '')
            if refdes:
                # Truncate refdes if too long
                refdes_display = refdes[:50] + ('...' if len(refdes) > 50 else '')
                info_parts.append(refdes_display)
            
            value = bom_part.get('value', '')
            if value:
                info_parts.append(value)
            
            package = bom_part.get('package', '')
            if package:
                info_parts.append(package)
            
            # Add other relevant fields
            voltage = bom_part.get('voltage', '')
            if voltage:
                info_parts.append(f"V:{voltage}")
            
            tolerance = bom_part.get('tolerance', '')
            if tolerance:
                info_parts.append(f"Tol:{tolerance}")
            
            power = bom_part.get('power', '')
            if power:
                info_parts.append(f"P:{power}")
            
            bom_info_text = " | ".join(info_parts) if info_parts else "Unknown part"
            bom_info_label = ttk.Label(self.results_frame, 
                                      text=bom_info_text,
                                      font=('TkDefaultFont', 9, 'bold'),
                                      foreground='darkblue')
            bom_info_label.pack(pady=(5, 0))
        
        # Show progress label if in batch mode (below BOM info)
        if show_navigation and total_count > 0:
            progress_label = ttk.Label(self.results_frame, 
                                      text=f"Part {current_index + 1} of {total_count}",
                                      font=('TkDefaultFont', 10, 'bold'))
            progress_label.pack(pady=(5, 10))
            
            # Navigation buttons and sort options
            nav_frame = ttk.Frame(self.results_frame)
            nav_frame.pack(pady=5)
            
            prev_btn = ttk.Button(nav_frame, text="Previous", 
                                 command=self.go_to_previous_part,
                                 state=tk.DISABLED if current_index == 0 else tk.NORMAL)
            prev_btn.pack(side=tk.LEFT, padx=5)
            
            next_btn = ttk.Button(nav_frame, text="Next", 
                                 command=self.go_to_next_part,
                                 state=tk.DISABLED if current_index >= total_count - 1 else tk.NORMAL)
            next_btn.pack(side=tk.LEFT, padx=5)
            
            # Sort preference radio buttons
            ttk.Label(nav_frame, text="Sort by:").pack(side=tk.LEFT, padx=(20, 5))
            # Pass part_key only - apply_sort_preference will get full results
            ttk.Radiobutton(nav_frame, text="Stock", variable=self.sort_preference, 
                          value='stock', command=lambda: self.apply_sort_preference(part_key)).pack(side=tk.LEFT, padx=2)
            ttk.Radiobutton(nav_frame, text="Price", variable=self.sort_preference, 
                          value='price', command=lambda: self.apply_sort_preference(part_key)).pack(side=tk.LEFT, padx=2)
        
        # Add sort options even when not in batch mode
        if not show_navigation:
            sort_frame = ttk.Frame(self.results_frame)
            sort_frame.pack(pady=5)
            ttk.Label(sort_frame, text="Sort by:").pack(side=tk.LEFT, padx=5)
            # Pass part_key only - apply_sort_preference will get full results
            ttk.Radiobutton(sort_frame, text="Stock", variable=self.sort_preference, 
                          value='stock', command=lambda: self.apply_sort_preference(part_key)).pack(side=tk.LEFT, padx=2)
            ttk.Radiobutton(sort_frame, text="Price", variable=self.sort_preference, 
                          value='price', command=lambda: self.apply_sort_preference(part_key)).pack(side=tk.LEFT, padx=2)
        
        if not results:
            no_results = ttk.Label(self.results_frame, text="No matching parts found")
            no_results.pack(pady=(20, 10))
            
            # Add custom keyword search option
            custom_search_frame = ttk.Frame(self.results_frame)
            custom_search_frame.pack(pady=10, padx=20, fill=tk.X)
            
            ttk.Label(custom_search_frame, text="Try a custom search keyword:", 
                     font=('TkDefaultFont', 9)).pack(anchor=tk.W, pady=(0, 5))
            
            # Text entry for custom keyword
            custom_keyword_var = tk.StringVar()
            custom_entry = ttk.Entry(custom_search_frame, textvariable=custom_keyword_var, width=40)
            custom_entry.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
            
            # Show the actual keyword that was sent to Mouser (if available), otherwise suggest one
            if part_key in self.last_search_keywords:
                actual_keyword = self.last_search_keywords[part_key]
                custom_keyword_var.set(actual_keyword)
                custom_entry.select_range(0, tk.END)  # Select all for easy editing
            elif bom_part:
                suggested_keyword = self._suggest_keyword(bom_part)
                if suggested_keyword:
                    custom_keyword_var.set(suggested_keyword)
                    custom_entry.select_range(0, tk.END)  # Select all for easy editing
            
            # Search button
            search_btn = ttk.Button(custom_search_frame, text="Search",
                                   command=lambda: self.search_with_custom_keyword(part_key, custom_keyword_var.get()))
            search_btn.pack(side=tk.LEFT, padx=5)
            
            # Allow Enter key to trigger search
            def on_enter(event):
                keyword = custom_keyword_var.get().strip()
                if keyword:
                    self.search_with_custom_keyword(part_key, keyword)
            
            custom_entry.bind('<Return>', on_enter)
            custom_entry.focus()
            
            # Still show N/A option
            self._add_na_option(part_key, radio_var)
        else:
            # Display each result with radio button
            for idx, part in enumerate(results):
                result_frame = ttk.Frame(self.results_frame, relief=tk.RIDGE, borderwidth=2)
                result_frame.pack(fill=tk.X, padx=5, pady=5)
                
                # Store frame reference for visual updates
                self.result_frames[(part_key, idx)] = result_frame
                
                # Make the entire frame clickable to select the radio button
                # Create click handler that selects this radio button
                click_handler = lambda event, k=part_key, i=idx: (radio_var.set(str(i)), self.on_radio_selected(k, i))
                result_frame.bind('<Button-1>', click_handler)
                
                # Check if package info is available
                package = part.get('package', '').strip() if part.get('package') else ''
                has_package = bool(package and package != '')
                # Rowspan will be calculated dynamically based on number of detail rows
                # Start with base: MPN (0), Description (1), Details (2+), Stock/Price (last)
                # We'll update rowspan after we know how many rows we have
                
                # Radio button for selection (XOR) - will adjust rowspan later
                radio_value = str(idx)
                radio_btn = ttk.Radiobutton(result_frame, variable=radio_var, value=radio_value,
                                           command=lambda k=part_key, i=idx: self.on_radio_selected(k, i))
                # Placeholder rowspan, will be updated
                radio_btn.grid(row=0, column=0, rowspan=6, padx=5, pady=5, sticky=tk.N)
                
                # Part number
                mpn = part.get('mpn', 'N/A')
                mouser_pn = part.get('mouser_part_number', '')
                part_num_text = f"MPN: {mpn}"
                if mouser_pn:
                    part_num_text += f"\nMouser: {mouser_pn}"
                
                part_num_label = ttk.Label(result_frame, text=part_num_text, font=('TkDefaultFont', 9, 'bold'))
                part_num_label.grid(row=0, column=1, sticky=tk.W, padx=5)
                part_num_label.bind('<Button-1>', click_handler)
                
                # Description
                desc = part.get('description', 'No description')
                desc_label = ttk.Label(result_frame, text=f"Description: {desc[:100]}", 
                                      wraplength=400, justify=tk.LEFT)
                desc_label.grid(row=1, column=1, sticky=tk.W, padx=5)
                desc_label.bind('<Button-1>', click_handler)
                
                # Additional details (below description, above price/stock)
                # Extract relevant specs from description or show what's available
                current_row = 2
                detail_parts = []
                
                # Package
                if has_package:
                    detail_parts.append(f"Package: {package}")
                
                # Try to extract specs from description using regex
                desc_upper = desc.upper()
                
                # Voltage pattern (e.g., "10V", "25VDC", "100V")
                voltage_match = re.search(r'\b(\d+\.?\d*)\s*V(?:DC|AC)?\b', desc_upper)
                if voltage_match:
                    detail_parts.append(f"Voltage: {voltage_match.group(1)}V")
                
                # Tolerance pattern (e.g., "5%", "10%", "±1%")
                tolerance_match = re.search(r'\b(\d+\.?\d*)\s*%|±\s*(\d+\.?\d*)\s*%', desc_upper)
                if tolerance_match:
                    tol_val = tolerance_match.group(1) or tolerance_match.group(2)
                    detail_parts.append(f"Tolerance: {tol_val}%")
                
                # Power/Wattage pattern (e.g., "0.1W", "1W", "1/4W")
                power_match = re.search(r'\b(\d+\.?\d*)\s*W|\b1/(\d+)\s*W', desc_upper)
                if power_match:
                    if power_match.group(2):
                        detail_parts.append(f"Power: 1/{power_match.group(2)}W")
                    else:
                        detail_parts.append(f"Power: {power_match.group(1)}W")
                
                # Temperature coefficient (e.g., "X7R", "X5R", "C0G", "NPO")
                temp_coef_match = re.search(r'\b(X7R|X5R|X6S|C0G|NPO|NP0)\b', desc_upper)
                if temp_coef_match:
                    detail_parts.append(f"Temp Coef: {temp_coef_match.group(1)}")
                
                # Manufacturer
                manufacturer = part.get('manufacturer', '')
                if manufacturer:
                    detail_parts.append(f"Mfr: {manufacturer}")
                
                # Add details row if we have any
                details_text = ""
                if detail_parts:
                    details_text = " | ".join(detail_parts[:5])  # Limit to first 5 details to avoid clutter
                    details_label = ttk.Label(result_frame, text=details_text, 
                                             font=('TkDefaultFont', 8),
                                             foreground='gray')
                    details_label.grid(row=current_row, column=1, sticky=tk.W, padx=5)
                    details_label.bind('<Button-1>', click_handler)
                    current_row += 1
                
                # Additional info (stock, price, lifecycle)
                info_parts = []
                if part.get('stock', 0) > 0:
                    info_parts.append(f"Stock: {part['stock']}")
                if part.get('price_breaks'):
                    price = part['price_breaks'][0].get('price', 'N/A')
                    # Remove $ if already present to avoid double $$
                    if isinstance(price, str):
                        price_clean = price.replace('$', '').strip()
                        info_parts.append(f"Price: ${price_clean}")
                    else:
                        info_parts.append(f"Price: ${price}")
                if part.get('lifecycle'):
                    info_parts.append(f"Status: {part['lifecycle']}")
                
                if info_parts:
                    info_label = ttk.Label(result_frame, text=" | ".join(info_parts))
                    info_label.grid(row=current_row, column=1, sticky=tk.W, padx=5)
                    info_label.bind('<Button-1>', click_handler)
                    current_row += 1
                
                # Update rowspan for radio button and link button based on actual rows used
                total_rows = current_row  # Total number of rows used
                radio_btn.grid_configure(rowspan=total_rows)
                
                # Link button
                button_frame = ttk.Frame(result_frame)
                button_frame.grid(row=0, column=2, rowspan=total_rows, padx=5, sticky=tk.E)
                
                product_url = part.get('product_url', '')
                if product_url:
                    link_btn = ttk.Button(button_frame, text="View on Mouser", 
                                         command=lambda url=product_url: webbrowser.open(url))
                    link_btn.pack(pady=2)
                
                result_frame.columnconfigure(1, weight=1)
            
            # Add N/A option
            self._add_na_option(part_key, radio_var, len(results))
        
        # Show "Get More Parts" button if there are more results
        all_results = self.current_search_results.get(part_key, [])
        current_idx = self.current_search_index.get(part_key, 0) + len(results)
        if current_idx < len(all_results):
            self.more_parts_btn.config(state=tk.NORMAL, 
                                      command=lambda: self.get_more_parts_for_key(part_key))
        else:
            self.more_parts_btn.config(state=tk.DISABLED)
        
        # Enable custom search button
        self.custom_search_btn.config(state=tk.NORMAL)
        
        # OK button at bottom (below Get More Parts button)
        if not hasattr(self, 'confirm_part_btn') or self.confirm_part_btn.winfo_exists() == 0:
            # Button will be created in right_frame
            pass
        
        self.status_var.set(f"Found {len(all_results)} matching parts")
        
        # Update scroll region after adding results
        if hasattr(self, 'results_canvas'):
            self.root.after_idle(lambda: self.results_canvas.configure(
                scrollregion=self.results_canvas.bbox('all')))
    
    def _add_na_option(self, part_key: str, radio_var: tk.StringVar, result_count: int = 0):
        """Add N/A option to result set."""
        na_frame = ttk.Frame(self.results_frame, relief=tk.RIDGE, borderwidth=2)
        na_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Store frame reference - use a special key to identify N/A frame
        # Use result_count as index, but we'll identify it differently in on_radio_selected
        na_index = result_count  # This will be the highest index
        self.result_frames[(part_key, na_index)] = na_frame
        
        # Make the entire frame clickable to select the radio button
        def on_na_frame_click(event):
            radio_var.set('NA')
            self.on_radio_selected(part_key, 'NA')
        na_frame.bind('<Button-1>', on_na_frame_click)
        
        # Radio button for N/A
        na_value = 'NA'
        radio_btn = ttk.Radiobutton(na_frame, variable=radio_var, value=na_value,
                                   command=lambda k=part_key: self.on_radio_selected(k, 'NA'))
        radio_btn.grid(row=0, column=0, rowspan=2, padx=5, pady=5, sticky=tk.N)
        
        # N/A label (also clickable)
        na_label = ttk.Label(na_frame, text="N/A - No part selected", 
                            font=('TkDefaultFont', 9, 'bold'), foreground='gray')
        na_label.grid(row=0, column=1, sticky=tk.W, padx=5)
        na_label.bind('<Button-1>', on_na_frame_click)
        
        desc_label = ttk.Label(na_frame, text="This part will be exported with MPN='NA'", 
                              foreground='gray')
        desc_label.grid(row=1, column=1, sticky=tk.W, padx=5)
        desc_label.bind('<Button-1>', on_na_frame_click)
        
        na_frame.columnconfigure(1, weight=1)
    
    def on_radio_selected(self, part_key: str, index):
        """Handle radio button selection - update visual appearance."""
        # Reset all frames to unselected appearance
        for (pk, idx), frame in self.result_frames.items():
            if pk == part_key:
                frame.config(relief=tk.RIDGE, borderwidth=2)
        
        # Highlight selected frame
        # Handle both int index and 'NA' string
        if isinstance(index, str) and index == 'NA':
            # Find the N/A frame - it's stored with result_count as index
            # Find the highest index for this part_key (N/A is always last)
            all_indices = [i for (p, i) in self.result_frames.keys() 
                          if p == part_key and isinstance(i, int)]
            if all_indices:
                na_index = max(all_indices)
                selected_frame = self.result_frames.get((part_key, na_index))
                logger.debug(f"Found N/A frame at index {na_index} for {part_key}")
            else:
                selected_frame = None
                logger.warning(f"Could not find N/A frame for {part_key}")
        else:
            # Regular part index
            try:
                idx = int(index) if isinstance(index, str) else index
                selected_frame = self.result_frames.get((part_key, idx))
            except (ValueError, TypeError):
                selected_frame = None
        
        if selected_frame:
            selected_frame.config(relief=tk.RAISED, borderwidth=3)
            logger.debug(f"Highlighted frame for {part_key} at index {index}")
        else:
            logger.warning(f"Could not find frame for {part_key} at index {index}")
        
        # Enable OK button
        if hasattr(self, 'confirm_part_btn'):
            self.confirm_part_btn.config(state=tk.NORMAL)
        
        self.status_var.set("Part selected - Click 'OK' to confirm")
    
    def get_more_parts_for_key(self, part_key: str):
        """Get more parts for a specific part key - refresh display with all results."""
        all_results = self.current_search_results.get(part_key, [])
        current_idx = self.current_search_index.get(part_key, 0)
        
        # Get next 3 parts
        next_results = all_results[current_idx:current_idx + 3]
        if next_results:
            self.current_search_index[part_key] = current_idx + len(next_results)
            # Re-display all results including new ones
            displayed_results = all_results[:self.current_search_index[part_key]]
            
            # Check if we're in batch mode
            if self.batch_part_keys and part_key in self.batch_part_keys:
                current_idx_batch = self.batch_part_keys.index(part_key)
                self.display_results(part_key, displayed_results, show_navigation=True,
                                   current_index=current_idx_batch, total_count=len(self.batch_part_keys))
            else:
                self.display_results(part_key, displayed_results)
            
            # Update button state
            if self.current_search_index[part_key] >= len(all_results):
                self.more_parts_btn.config(state=tk.DISABLED)
    
    def get_more_parts(self):
        """Get more parts for currently selected part."""
        part_key = self.get_selected_part_key()
        if part_key:
            self.get_more_parts_for_key(part_key)
    
    def show_custom_search_dialog(self):
        """Show a dialog for custom keyword search."""
        # Determine which part_key we're working with
        if self.batch_part_keys and self.current_batch_index < len(self.batch_part_keys):
            part_key = self.batch_part_keys[self.current_batch_index]
        else:
            part_key = self.get_selected_part_key()
            if not part_key:
                messagebox.showwarning("No Selection", "Please select a part from the BOM table first")
                return
        
        # Get the original BOM part data
        bom_part = self._get_part_by_key(part_key)
        if not bom_part:
            messagebox.showerror("Error", "Could not find part data")
            return
        
        # Create dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title("Custom Keyword Search")
        dialog.geometry("500x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Enter a custom search keyword:", 
                 font=('TkDefaultFont', 9)).pack(anchor=tk.W, pady=(0, 5))
        
        # Text entry for custom keyword
        custom_keyword_var = tk.StringVar()
        
        # Show the actual keyword that was sent to Mouser (if available), otherwise suggest one
        if part_key in self.last_search_keywords:
            actual_keyword = self.last_search_keywords[part_key]
            custom_keyword_var.set(actual_keyword)
        elif bom_part:
            suggested_keyword = self._suggest_keyword(bom_part)
            if suggested_keyword:
                custom_keyword_var.set(suggested_keyword)
        
        custom_entry = ttk.Entry(main_frame, textvariable=custom_keyword_var, width=60)
        custom_entry.pack(fill=tk.X, pady=(0, 10))
        custom_entry.select_range(0, tk.END)
        custom_entry.focus()
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        def do_search():
            keyword = custom_keyword_var.get().strip()
            if keyword:
                dialog.destroy()
                self.search_with_custom_keyword(part_key, keyword)
            else:
                messagebox.showwarning("Empty Keyword", "Please enter a search keyword")
        
        def cancel():
            dialog.destroy()
        
        ttk.Button(button_frame, text="Search", command=do_search).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=cancel).pack(side=tk.RIGHT)
        
        # Allow Enter key to trigger search
        custom_entry.bind('<Return>', lambda e: do_search())
    
    def on_part_selected_checkbox(self, part: Dict[str, Any], var: tk.BooleanVar, part_key: str):
        """Handle checkbox selection for a Mouser part (deprecated - now using radio buttons)."""
        # This method is kept for backward compatibility but radio buttons are used instead
        pass
    
    def confirm_selected_part(self):
        """Confirm the currently selected part from radio button."""
        # Find which part_key we're working with
        # Check if we're in batch mode
        if self.batch_part_keys and self.current_batch_index < len(self.batch_part_keys):
            part_key = self.batch_part_keys[self.current_batch_index]
        else:
            # Single part mode - get from current selection
            part_key = self.get_selected_part_key()
            if not part_key:
                messagebox.showwarning("No Selection", "Please select a part from the BOM table first")
                return
        
        # Get selected radio button value
        if part_key not in self.radio_vars:
            messagebox.showwarning("No Selection", "Please select a part option first")
            return
        
        radio_var = self.radio_vars[part_key]
        selected_value = radio_var.get()
        
        if not selected_value:
            messagebox.showwarning("No Selection", "Please select a part option first")
            return
        
        # Get the selected part
        if selected_value == 'NA':
            # N/A option selected
            mouser_part = {
                'mpn': 'NA',
                'mouser_part_number': 'NA',
                'description': 'N/A - No part selected',
                'manufacturer': 'N/A'
            }
        else:
            # Get part from results
            try:
                part_index = int(selected_value)
                all_results = self.current_search_results.get(part_key, [])
                if part_index < len(all_results):
                    mouser_part = all_results[part_index]
                else:
                    messagebox.showerror("Error", "Selected part index out of range")
                    return
            except (ValueError, IndexError) as e:
                logger.error(f"Error getting selected part: {e}")
                messagebox.showerror("Error", "Could not find selected part")
                return
        
        # Confirm the selection
        self.confirm_part_selection(part_key, mouser_part)
        
        # If in batch mode, advance to next part
        if self.batch_part_keys and self.current_batch_index < len(self.batch_part_keys) - 1:
            self.confirm_and_advance()
    
    def confirm_part_selection(self, part_key: str, mouser_part: Dict[str, Any]):
        """Confirm the selected Mouser part and update BOM table. Allows overwriting existing selections."""
        mpn = mouser_part.get('mpn', 'N/A')
        logger.info(f"Confirming part selection for {part_key}: {mpn}")
        
        # Check if this part_key already has a selection (allow overwrite)
        if part_key in self.selected_parts:
            logger.info(f"Overwriting existing selection for {part_key}: {self.selected_parts[part_key].get('mpn')} -> {mpn}")
        
        # Store the selected part (overwrites if exists)
        self.selected_parts[part_key] = mouser_part
        
        # Check the checkbox in the BOM table and turn row green
        self.part_selected[part_key] = True
        
        # Find the item_id for this part_key and update it
        item_id = None
        logger.debug(f"Looking for item_id for part_key {part_key}")
        logger.debug(f"item_to_part_key mapping has {len(self.item_to_part_key)} entries")
        
        for item, pk in self.item_to_part_key.items():
            if pk == part_key:
                item_id = item
                logger.debug(f"Found matching item_id {item_id} for part_key {part_key}")
                break
        
        if item_id:
            logger.info(f"Updating BOM table row {item_id} for part_key {part_key}")
            # Check if this is N/A selection
            is_na = (mpn == 'NA' or mouser_part.get('mpn') == 'NA')
            self.update_row_checkbox(item_id, part_key, is_na=is_na)
            if is_na:
                self.status_var.set(f"Confirmed: N/A - Row updated and marked grey")
            else:
                self.status_var.set(f"Confirmed: {mpn} - Row updated and checked")
            
            # Show confirmation message (only if not in batch mode to avoid spam)
            if not (self.batch_part_keys and self.current_batch_index < len(self.batch_part_keys)):
                messagebox.showinfo("Part Confirmed", 
                                  f"Part {mpn} has been added to your BOM selection.\n"
                                  f"The row has been marked green and will be included in export.")
        else:
            logger.error(f"Could not find item_id for part_key {part_key}")
            logger.debug(f"Available part_keys in mapping: {list(self.item_to_part_key.values())[:5]}...")
            self.status_var.set(f"Confirmed: {mpn} (table update failed - part not found)")
            if not (self.batch_part_keys and self.current_batch_index < len(self.batch_part_keys)):
                messagebox.showwarning("Update Warning", 
                                      f"Part {mpn} was selected but could not update the BOM table.\n"
                                      f"Please check the logs for details.")
    
    def show_search_error(self, error_msg: str):
        """Show search error message."""
        self.clear_results()
        error_label = ttk.Label(self.results_frame, text=f"Search error: {error_msg}", 
                               foreground='red')
        error_label.pack(pady=20)
        self.status_var.set("Search failed")
    
    def get_export_data(self) -> List[Dict[str, Any]]:
        """Get the export data for checked parts. Returns empty list if no parts checked."""
        # Filter to only include parts with checked checkboxes
        checked_parts = {k: v for k, v in self.selected_parts.items() 
                        if self.part_selected.get(k, False)}
        
        if not checked_parts:
            return []
        
        # Build export data
        export_data = []
        for part_key, mouser_part in checked_parts.items():
            # Get original part data using index lookup
            original_part = self._get_part_by_key(part_key)
            
            # Get package from Mouser part, leave blank if not available
            # Handle None, empty string, or missing field
            mouser_package_raw = mouser_part.get('package')
            if mouser_package_raw and isinstance(mouser_package_raw, str):
                mouser_package = mouser_package_raw.strip()
            else:
                mouser_package = ''
                # Debug logging for missing package
                mpn = mouser_part.get('mpn', 'Unknown')
                logger.debug(f"Package field missing or empty for part {mpn} (part_key: {part_key}). Available keys: {list(mouser_part.keys())}")
            
            row = {
                'REFDES': original_part.get('refdes', '') if original_part else '',
                'Quantity': original_part.get('quantity', '') if original_part else '',
                'Description': mouser_part.get('description', original_part.get('description', '') if original_part else ''),
                'Package': mouser_package,  # Use package from selected Mouser component
                'MPN': mouser_part.get('mpn', ''),
                'Mouser Part Number': mouser_part.get('mouser_part_number', ''),
                'Manufacturer': mouser_part.get('manufacturer', ''),
                'Value': original_part.get('value', '') if original_part else '',
                'Voltage': original_part.get('voltage', '') if original_part else '',
                'Stock': mouser_part.get('stock', 0),
                'Price': mouser_part.get('price_breaks', [{}])[0].get('price', '') if mouser_part.get('price_breaks') else '',
                'Lifecycle': mouser_part.get('lifecycle', ''),
                'Product URL': mouser_part.get('product_url', ''),
            }
            export_data.append(row)
        
        return export_data
    
    def preview_bom(self):
        """Preview the exported BOM in a dialog window."""
        export_data = self.get_export_data()
        
        if not export_data:
            logger.warning("Attempted preview with no checked parts")
            messagebox.showwarning("No Selections", "No parts have been checked for export. Please select Mouser parts and ensure the checkbox is checked.")
            return
        
        # Create preview window
        preview_window = tk.Toplevel(self.root)
        preview_window.title(f"Preview BOM - {len(export_data)} parts")
        preview_window.geometry("1200x600")
        
        # Create frame for treeview and scrollbars
        frame = ttk.Frame(preview_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Create treeview
        columns = list(export_data[0].keys())
        tree = ttk.Treeview(frame, columns=columns, show='headings', height=20)
        
        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, minwidth=100)
        
        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        h_scrollbar = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout
        tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        
        # Populate treeview
        for row in export_data:
            values = [str(row.get(col, '')) for col in columns]
            tree.insert('', tk.END, values=values)
        
        # Status label
        status_label = ttk.Label(preview_window, text=f"Previewing {len(export_data)} parts")
        status_label.pack(pady=5)
        
        logger.info(f"Previewing {len(export_data)} parts")
    
    def export_bom(self):
        """Export selected parts to CSV/Excel (only rows with checked checkboxes)."""
        export_data = self.get_export_data()
        
        if not export_data:
            logger.warning("Attempted export with no checked parts")
            messagebox.showwarning("No Selections", "No parts have been checked for export. Please select Mouser parts and ensure the checkbox is checked.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Export BOM",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not file_path:
            logger.info("Export cancelled by user")
            return
            
        logger.info(f"Exporting BOM to: {file_path}")
        
        try:
            
            # Write to file
            if file_path.endswith('.csv'):
                with open(file_path, 'w', newline='', encoding='utf-8') as f:
                    if export_data:
                        writer = csv.DictWriter(f, fieldnames=export_data[0].keys())
                        writer.writeheader()
                        writer.writerows(export_data)
            else:
                # Excel export
                try:
                    import openpyxl
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    
                    if export_data:
                        # Write headers
                        headers = list(export_data[0].keys())
                        ws.append(headers)
                        
                        # Write data
                        for row in export_data:
                            ws.append([row.get(h, '') for h in headers])
                    
                    wb.save(file_path)
                except ImportError:
                    messagebox.showerror("Error", "openpyxl required for Excel export")
                    return
            
            messagebox.showinfo("Success", f"Exported {len(export_data)} parts to {file_path}")
            self.status_var.set(f"Exported {len(export_data)} parts")
            logger.info(f"Successfully exported {len(export_data)} parts")
            
        except Exception as e:
            logger.error(f"Export failed: {e}", exc_info=True)
            messagebox.showerror("Error", f"Failed to export BOM:\n{e}")
    
    def save_bom_state(self):
        """Save all BOM state (parts, selections, search results, options) to a JSON file."""
        if not self.consolidated_parts:
            messagebox.showwarning("No BOM", "No BOM data to save. Please open a BOM file first.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Save BOM State",
            defaultextension=".bomhelper",
            filetypes=[("BOM Helper files", "*.bomhelper"), ("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if not file_path:
            logger.info("Save cancelled by user")
            return
        
        logger.info(f"Saving BOM state to: {file_path}")
        
        try:
            # Collect all state data
            state_data = {
                'version': '1.0',
                'save_date': datetime.now().isoformat(),
                
                # BOM Data
                'consolidated_parts': self.consolidated_parts,
                'column_mapping': self.column_mapping,
                'components': self.components,  # Optional but useful for completeness
                
                # User Selections
                'selected_parts': self.selected_parts,
                'part_selected': self.part_selected,
                
                # Search Options
                'in_stock_only': self.in_stock_var.get(),
                'active_only': self.active_only_var.get(),
                'sort_preference': self.sort_preference.get(),
                
                # Search Results
                'current_search_results': self.current_search_results,
                'batch_part_keys': self.batch_part_keys,
                'current_batch_index': self.current_batch_index,
            }
            
            # Convert any non-serializable types
            # tkinter variables are already converted above
            # Search results may contain complex nested structures, JSON should handle them
            
            # Serialize to JSON with proper formatting
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(state_data, f, indent=2, ensure_ascii=False)
            
            messagebox.showinfo("Success", f"BOM state saved to:\n{file_path}")
            self.status_var.set(f"BOM state saved")
            logger.info(f"Successfully saved BOM state with {len(self.consolidated_parts)} parts")
            
        except Exception as e:
            logger.error(f"Save failed: {e}", exc_info=True)
            messagebox.showerror("Error", f"Failed to save BOM state:\n{e}")
    
    def load_bom_state(self):
        """Load BOM state from a JSON file and restore all application state."""
        file_path = filedialog.askopenfilename(
            title="Load BOM State",
            defaultextension=".bomhelper",
            filetypes=[("BOM Helper files", "*.bomhelper"), ("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if not file_path:
            logger.info("Load cancelled by user")
            return
        
        logger.info(f"Loading BOM state from: {file_path}")
        
        try:
            # Read JSON file
            with open(file_path, 'r', encoding='utf-8') as f:
                state_data = json.load(f)
            
            # Validate file format
            if not isinstance(state_data, dict):
                raise ValueError("Invalid file format: expected a dictionary")
            
            version = state_data.get('version', 'unknown')
            logger.info(f"Loading BOM state version: {version}")
            
            # Restore BOM Data
            if 'consolidated_parts' in state_data:
                self.consolidated_parts = state_data['consolidated_parts']
                logger.info(f"Restored {len(self.consolidated_parts)} consolidated parts")
            else:
                messagebox.showerror("Error", "File does not contain BOM parts data")
                return
            
            # Restore column mapping
            if 'column_mapping' in state_data:
                self.column_mapping = state_data['column_mapping']
            
            # Restore original components (optional)
            if 'components' in state_data:
                self.components = state_data['components']
            
            # Rebuild part_key to index mapping
            self.part_key_to_index = {}
            for idx, part in enumerate(self.consolidated_parts):
                part_key = self._generate_part_key(idx)
                self.part_key_to_index[part_key] = idx
            
            # Redisplay BOM table (this will also rebuild item_to_part_key mapping)
            self.populate_parts_table()
            
            # Restore checkbox states and row colors
            if 'part_selected' in state_data:
                self.part_selected = state_data['part_selected']
                # Update checkboxes in table
                for item_id, part_key in self.item_to_part_key.items():
                    if part_key in self.part_selected:
                        self.update_row_checkbox(item_id, part_key)
            
            # Restore user selections (selected Mouser parts)
            if 'selected_parts' in state_data:
                self.selected_parts = state_data['selected_parts']
            
            # Restore search options
            if 'in_stock_only' in state_data:
                self.in_stock_var.set(state_data['in_stock_only'])
            if 'active_only' in state_data:
                self.active_only_var.set(state_data['active_only'])
            if 'sort_preference' in state_data:
                self.sort_preference.set(state_data['sort_preference'])
            
            # Restore search results
            if 'current_search_results' in state_data:
                self.current_search_results = state_data['current_search_results']
            
            # Restore batch navigation state
            if 'batch_part_keys' in state_data:
                self.batch_part_keys = state_data['batch_part_keys']
            if 'current_batch_index' in state_data:
                self.current_batch_index = state_data['current_batch_index']
                
                # If we have batch state, display the current part's results
                if self.batch_part_keys and 0 <= self.current_batch_index < len(self.batch_part_keys):
                    part_key = self.batch_part_keys[self.current_batch_index]
                    if part_key in self.current_search_results:
                        results = self.current_search_results[part_key]
                        self.display_results(part_key, results, show_navigation=True,
                                           current_index=self.current_batch_index,
                                           total_count=len(self.batch_part_keys))
            elif self.current_search_results:
                # If we have search results but no batch state, display first part's results
                first_part_key = next(iter(self.current_search_results))
                results = self.current_search_results[first_part_key]
                self.display_results(first_part_key, results[:3], show_navigation=False)
            
            save_date = state_data.get('save_date', 'unknown')
            messagebox.showinfo("Success", f"BOM state loaded successfully.\nSaved: {save_date}")
            self.status_var.set(f"BOM state loaded")
            logger.info(f"Successfully loaded BOM state from {file_path}")
            
        except FileNotFoundError:
            logger.error(f"File not found: {file_path}")
            messagebox.showerror("Error", f"File not found:\n{file_path}")
        except json.JSONDecodeError as e:
            logger.error(f"Invalid JSON file: {e}")
            messagebox.showerror("Error", f"Invalid JSON file:\n{e}")
        except Exception as e:
            logger.error(f"Load failed: {e}", exc_info=True)
            messagebox.showerror("Error", f"Failed to load BOM state:\n{e}")
    
    def show_api_keys_dialog(self):
        """Show API keys configuration dialog."""
        dialog = tk.Toplevel(self.root)
        dialog.title("API Keys Configuration")
        dialog.geometry("500x200")
        
        ttk.Label(dialog, text="API keys are loaded from keys.txt file.").pack(pady=10)
        ttk.Label(dialog, text=f"Mouser API Key: {'✓ Configured' if self.config.get_mouser_api_key() else '✗ Not configured'}").pack(pady=5)
        ttk.Label(dialog, text=f"Gemini API Key: {'✓ Configured' if self.config.get_gemini_api_key() else '✗ Not configured'}").pack(pady=5)
        
        ttk.Button(dialog, text="OK", command=dialog.destroy).pack(pady=20)


def main():
    """Main entry point."""
    root = tk.Tk()
    app = BOMMouserLookupApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

