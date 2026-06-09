import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import font as tkfont

class MicrosoftPowerPoint2007Ultimate:
    def __init__(self, root):
        self.root = root
        self.root.title("Microsoft PowerPoint 2007 - Ultimate Luna Blue Suite")
        
        # Exact Window Size configuration with standard 'x' notation to eliminate black screen
        self.root.geometry("1200x800")
        self.root.minsize(1000, 720)
        
        # =========================================================================
        # 1. GRAPHICS ENGINE & AUTOCRATIC COLOR MATRIX (OFFICE 2007 MASTER THEMES)
        # =========================================================================
        self.themes_registry = {
            "Luna Blue": {
                "outer_bg": "#2b3e5a", "apron_bg": "#BED4ED", "ribbon_body": "#D4E2F3",
                "text_deep": "#1E395B", "active_tab": "#F2F6FB", "sidebar_bg": "#D9E4F4",
                "accent_orange": "#E2A936", "hover_yellow": "#FFF4CE", "slide_bg": "#FFFFFF"
            },
            "Classic Silver": {
                "outer_bg": "#494949", "apron_bg": "#D1D3D4", "ribbon_body": "#E6E7E8",
                "text_deep": "#222222", "active_tab": "#F1F2F2", "sidebar_bg": "#E6E7E8",
                "accent_orange": "#A7A9AC", "hover_yellow": "#E6E7E8", "slide_bg": "#FFFFFF"
            },
            "Obsidian Black": {
                "outer_bg": "#111111", "apron_bg": "#3A3A3A", "ribbon_body": "#555555",
                "text_deep": "#FFFFFF", "active_tab": "#222222", "sidebar_bg": "#444444",
                "accent_orange": "#FFCC00", "hover_yellow": "#666666", "slide_bg": "#1A1A1A"
            },
            "Forest Green": {
                "outer_bg": "#1E2F23", "apron_bg": "#A3C1AD", "ribbon_body": "#C1D7C4",
                "text_deep": "#0E1A12", "active_tab": "#E8F0EA", "sidebar_bg": "#CDE0D2",
                "accent_orange": "#708D75", "hover_yellow": "#F0F7F2", "slide_bg": "#FFFFFF"
            }
        }
        self.current_theme_token = "Luna Blue"
        self.apply_theme_vectors()

        # =========================================================================
        # 2. PRESENTATION STORAGE ARCHITECTURE DATABASE
        # =========================================================================
        self.presentation_database = [
            {
                "title": "Welcome to PowerPoint 2007",
                "subtitle": "The ultimate standalone single-file production suite build.",
                "shapes": [
                    {"type": "rectangle", "coords": [50, 400, 250, 480], "color": "#1E395B", "fill_color": "#D4E2F3"},
                    {"type": "oval", "coords": [400, 380, 500, 480], "color": "#E2A936", "fill_color": "#FFF4CE"}
                ],
                "notes": "Speaker Notes: Welcome the audience to the presentation.",
                "transition": "Fade"
            },
            {
                "title": "Interactive Shape Vectors Engine",
                "subtitle": "Inject geometric wireframes and reposition them directly using mouse grids.",
                "shapes": [
                    {"type": "line", "coords": [100, 420, 600, 420], "color": "#A80F1A", "fill_color": ""}
                ],
                "notes": "Speaker Notes: Demonstrate live shape interactions on canvas.",
                "transition": "Slide Left"
            }
        ]
        self.active_slide_ptr = 0
        self.active_tab_pointer = "Home"
        self.selected_shape_index = None
        self.drag_start_coordinates = (0, 0)
        self.is_slideshow_active = False

        # Build Core Subsystems
        self.build_hardware_window_shell()
        self.refresh_ribbon_tab_view("Home")
        self.rebuild_slide_sorter_rail()
        self.load_active_slide_to_viewport()

    def apply_theme_vectors(self):
        t = self.themes_registry[self.current_theme_token]
        self.CLR_OUTER_BG = t["outer_bg"]
        self.CLR_LUNA_BLUE = t["apron_bg"]
        self.CLR_RIBBON_BODY = t["ribbon_body"]
        self.CLR_TEXT_DEEP = t["text_deep"]
        self.CLR_ACTIVE_TAB = t["active_tab"]
        self.CLR_SIDEBAR_BG = t["sidebar_bg"]
        self.CLR_ACCENT_ORANGE = t["accent_orange"]
        self.CLR_HOVER_YELLOW = t["hover_yellow"]
        self.CLR_SLIDE_BG = t["slide_bg"]

    # =========================================================================
    # 3. WINDOW INFRASTRUCTURE & LAYOUT DESKTOP FRAMEWORK
    # =========================================================================
    def build_hardware_window_shell(self):
        self.root.configure(bg=self.CLR_OUTER_BG)
        
        # Ribbon Top Banner Matrix
        self.ribbon_apron = tk.Frame(self.root, bg=self.CLR_LUNA_BLUE, height=135)
        self.ribbon_apron.pack(fill=tk.X, side=tk.TOP, padx=4, pady=(4, 0))
        self.ribbon_apron.pack_propagate(False)
        
        # The Premium Metallic Office Orb Menu Button Vector Graphic
        self.orb_canvas = tk.Canvas(self.ribbon_apron, width=48, height=48, bg=self.CLR_LUNA_BLUE, bd=0, highlightthickness=0, cursor="hand2")
        self.orb_canvas.place(x=6, y=6)
        self.orb_canvas.create_oval(2, 2, 46, 46, fill="#D44E00", outline="#FFFFFF", width=2)
        self.orb_canvas.create_oval(6, 6, 42, 42, fill="#FFB03A", outline="")
        self.orb_canvas.create_rectangle(16, 16, 23, 22, fill="#FFFFFF", outline="")
        self.orb_canvas.create_rectangle(25, 16, 32, 22, fill="#FFFFFF", outline="")
        self.orb_canvas.create_rectangle(16, 24, 23, 30, fill="#FFFFFF", outline="")
        self.orb_canvas.create_rectangle(25, 24, 32, 30, fill="#FFFFFF", outline="")
        self.orb_canvas.bind("<Button-1>", lambda e: self.deploy_office_orb_dropdown_menu())

        # Ribbon Menu Tab Switches Layout Setup Array
        self.tabs_uicontrols_dictionary = {}
        tab_headers_pool = ["Home", "Insert", "Design", "Animations", "Slide Show"]
        for idx, tab_title in enumerate(tab_headers_pool):
            frame_node = tk.Frame(self.ribbon_apron, bg=self.CLR_LUNA_BLUE, bd=0)
            frame_node.place(x=70 + (idx * 80), y=28, width=75, height=24)
            
            lbl_node = tk.Label(frame_node, text=tab_title, font=("Segoe UI", 9, "bold" if tab_title == "Home" else "normal"), bg=self.CLR_LUNA_BLUE, fg=self.CLR_TEXT_DEEP)
            lbl_node.pack(fill=tk.BOTH, expand=True)
            
            self.tabs_uicontrols_dictionary[tab_title] = (frame_node, lbl_node)
            lbl_node.bind("<Button-1>", lambda e, name=tab_title: self.refresh_ribbon_tab_view(name))

        # Core Tools Command Matrix Container Area
        self.ribbon_tool_belt = tk.Frame(self.ribbon_apron, bg=self.CLR_RIBBON_BODY, bd=1, relief=tk.SOLID)
        self.ribbon_tool_belt.place(x=0, y=54, relwidth=1.0, height=80)

        # Primary Splitted Midsection Workspace Split Box Configuration
        self.middle_split_box = tk.Frame(self.root, bg=self.CLR_OUTER_BG)
        self.middle_split_box.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        
        # Left Panel Track Sorter System Rail Frame Setup
        self.sidebar_rail_track = tk.Frame(self.middle_split_box, bg=self.CLR_SIDEBAR_BG, width=210, bd=1, relief=tk.SUNKEN)
        self.sidebar_rail_track.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar_rail_track.pack_propagate(False)
        
        lbl_sidebar_header = tk.Label(self.sidebar_rail_track, text="Slides Deck Index", font=("Segoe UI", 9, "bold"), bg="#CBD9EC", fg=self.CLR_TEXT_DEEP, bd=1, relief=tk.RAISED, pady=4)
        lbl_sidebar_header.pack(fill=tk.X)
        
        self.sidebar_scroll_canvas = tk.Canvas(self.sidebar_rail_track, bg=self.CLR_SIDEBAR_BG, bd=0, highlightthickness=0)
        self.sidebar_scroll_canvas.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Central Master Display Presenter Engine Architecture Stage
        self.workspace_viewport_center_stage = tk.Frame(self.middle_split_box, bg="#566E8D")
        self.workspace_viewport_center_stage.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Fixed 4:3 Aspect Ratio Slide Sheet Boundary Container Box System
        self.slide_sheet_board = tk.Frame(self.workspace_viewport_center_stage, bg=self.CLR_SLIDE_BG, width=720, height=540, bd=1, relief=tk.SOLID)
        self.slide_sheet_board.place(relx=0.5, rely=0.46, anchor=tk.CENTER)
        self.slide_sheet_board.pack_propagate(False)
        
        # Primary Interactive Vector Shape Graphics Render Engine Canvas
        self.slide_graphics_canvas = tk.Canvas(self.slide_sheet_board, bg=self.CLR_SLIDE_BG, highlightthickness=0, bd=0)
        self.slide_graphics_canvas.place(x=0, y=0, relwidth=1.0, relheight=1.0)
        
        # Attach Core Hardware Graphics Mouse Input Triggers to Canvas Grid Vector Objects
        self.slide_graphics_canvas.bind("<Button-1>", self.execute_mouse_click_shape_detection)
        self.slide_graphics_canvas.bind("<B1-Motion>", self.execute_mouse_drag_shape_repositioning)
        self.slide_graphics_canvas.bind("<ButtonRelease-1>", lambda e: self.finalize_shape_transformation_pipeline())

        # Document Text Input Field Box Layers (Positioned precisely within 4:3 grid matrix sheets)
        self.input_text_headline = tk.Text(self.slide_sheet_board, font=("Segoe UI", 28, "bold"), fg=self.CLR_TEXT_DEEP, bg=self.CLR_SLIDE_BG, bd=1, relief=tk.DASHED, height=2, wrap=tk.WORD, insertbackground=self.CLR_TEXT_DEEP)
        self.input_text_headline.place(x=45, y=55, width=630, height=85)
        self.input_text_headline.bind("<KeyRelease>", self.synchronize_workspace_inputs_to_memory)
        
        self.input_text_subline = tk.Text(self.slide_sheet_board, font=("Segoe UI", 14), fg="#555555" if self.current_theme_token != "Obsidian Black" else "#CCCCCC", bg=self.CLR_SLIDE_BG, bd=1, relief=tk.DASHED, height=6, wrap=tk.WORD, insertbackground=self.CLR_TEXT_DEEP)
        self.input_text_subline.place(x=65, y=175, width=590, height=200)
        self.input_text_subline.bind("<KeyRelease>", self.synchronize_workspace_inputs_to_memory)

        # Bottom Executive Real-Time Presenter Speaker Notes Tray Section Frame
        self.speaker_notes_panel = tk.Frame(self.workspace_viewport_center_stage, bg=self.CLR_SIDEBAR_BG, height=80, bd=1, relief=tk.SUNKEN)
        self.speaker_notes_panel.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0, 10))
        
        lbl_notes_flag = tk.Label(self.speaker_notes_panel, text="Click to add speaker notes:", font=("Segoe UI", 8, "italic"), bg=self.CLR_SIDEBAR_BG, fg=self.CLR_TEXT_DEEP, anchor=tk.W)
        lbl_notes_flag.pack(fill=tk.X, padx=5, pady=2)
        
        self.input_speaker_notes_field = tk.Text(self.speaker_notes_panel, font=("Segoe UI", 10), bg="#FFFFFF", fg="#222222", bd=1, relief=tk.SOLID, height=2)
        self.input_speaker_notes_field.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))
        self.input_speaker_notes_field.bind("<KeyRelease>", self.synchronize_workspace_inputs_to_memory)

        # Bottom System Appearance Status Strip Notification Line
        self.status_footer_line = tk.Frame(self.root, bg="#CBD9EC", height=24, bd=1, relief=tk.SUNKEN)
        self.status_footer_line.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.lbl_status_indexer = tk.Label(self.status_footer_line, text="Slide 1 of 2", font=("Segoe UI", 9), bg="#CBD9EC", fg=self.CLR_TEXT_DEEP)
        self.lbl_status_indexer.pack(side=tk.LEFT, padx=15)
        
        self.lbl_status_theme_flag = tk.Label(self.status_footer_line, text=f"Active Theme: {self.current_theme_token}", font=("Segoe UI", 9), bg="#CBD9EC", fg=self.CLR_TEXT_DEEP)
        self.lbl_status_theme_flag.pack(side=tk.LEFT, padx=30)
        
        lbl_system_signature = tk.Label(self.status_footer_line, text="Microsoft Office PowerPoint 2007 Engine Framework Standard Core Build V2.06 [OK]", font=("Consolas", 9, "bold"), bg="#CBD9EC", fg="#4C607A")
        lbl_system_signature.pack(side=tk.RIGHT, padx=20)

    # =========================================================================
    # 4. RIBBON ARCHITECTURE ALTERNATION MULTI-TAB CONTROLLER SUITE
    # =========================================================================
    def refresh_ribbon_tab_view(self, target_tab_token):
        self.active_tab_pointer = target_tab_token
        
        # Destruct and Purge Current Active Control Layout Components Inside Toolbelt Frame Container
        for active_child_node in self.ribbon_tool_belt.winfo_children():
            active_child_node.destroy()
            
        # Graphical State Mapping Rebalancer Graph Engine Loop
        for name, UI_tuple in self.tabs_uicontrols_dictionary.items():
            if name == target_tab_token:
                UI_tuple[0].configure(bg=self.CLR_RIBBON_BODY, bd=1, relief=tk.RAISED)
                UI_tuple[1].configure(bg=self.CLR_RIBBON_BODY, font=("Segoe UI", 9, "bold"))
            else:
                UI_tuple[0].configure(bg=self.CLR_LUNA_BLUE, bd=0, relief=tk.FLAT)
                UI_tuple[1].configure(bg=self.CLR_LUNA_BLUE, font=("Segoe UI", 9, "normal"))

        # Assembly Pipeline Map Routers
        if target_tab_token == "Home":
            self.assemble_home_tab_commands()
        elif target_tab_token == "Insert":
            self.assemble_insert_tab_commands()
        elif target_tab_token == "Design":
            self.assemble_design_tab_commands()
        elif target_tab_token == "Animations":
            self.assemble_animations_tab_commands()
        elif target_tab_token == "Slide Show":
            self.assemble_slideshow_tab_commands()

    # --- HOME TAB CONTROL GROUP ---
    def assemble_home_tab_commands(self):
        self.generate_ribbon_divider_accent(5, "Clipboard", 100)
        self.inject_large_action_node(15, "Paste\nData", self.execute_clipboard_paste_simulation)
        self.inject_small_stacked_action(75, 6, "Cut Text Block", lambda: self.execute_text_formatting_processor("cut"))
        self.inject_small_stacked_action(75, 32, "Copy Buffer", lambda: self.execute_text_formatting_processor("copy"))
        
        self.generate_ribbon_divider_accent(115, "Slides Deck", 145)
        self.inject_large_action_node(125, "New\nSlide", self.execute_append_new_slide_instance)
        self.inject_small_stacked_action(190, 6, "Reset Elements", self.execute_reset_canvas_text_fields)
        self.inject_small_stacked_action(190, 32, "Delete Slide Node", self.execute_purge_selected_slide_instance, variant_fg="#A80F1A")
        
        self.generate_ribbon_divider_accent(270, "Font Alignment Suite", 260)
        self.inject_small_stacked_action(280, 6, "Bold Weight", lambda: self.execute_text_formatting_processor("bold"))
        self.inject_small_stacked_action(280, 32, "Italic Slant", lambda: self.execute_text_formatting_processor("italic"))
        self.inject_small_stacked_action(385, 6, "Underline Line", lambda: self.execute_text_formatting_processor("underline"))
        self.inject_small_stacked_action(385, 32, "Color Crimson", lambda: self.execute_text_formatting_processor("color", "#A80F1A"))
        self.inject_small_stacked_action(480, 6, "Align Left", lambda: self.execute_text_formatting_processor("align_left"))
        self.inject_small_stacked_action(480, 32, "Align Center", lambda: self.execute_text_formatting_processor("align_center"))

        self.generate_ribbon_divider_accent(540, "Search Suite", 110)
        self.inject_large_action_node(550, "Find &\nReplace", self.deploy_precision_search_modal)

    # --- INSERT TAB CONTROL GROUP ---
    def assemble_insert_tab_commands(self):
        self.generate_ribbon_divider_accent(5, "Tables Generation Matrix", 110)
        self.inject_large_action_node(15, "Insert\nTable", lambda: messagebox.showinfo("Table Engine", "Office Table Grid Matrix Generator Node initialized."))
        
        self.generate_ribbon_divider_accent(125, "Illustrations Vector Shapes System", 410)
        self.inject_small_stacked_action(135, 6, "+ Add Rectangle Box", lambda: self.inject_vector_shape_to_slide_db("rectangle"))
        self.inject_small_stacked_action(135, 32, "+ Add Oval Circular", lambda: self.inject_vector_shape_to_slide_db("oval"))
        self.inject_small_stacked_action(285, 6, "+ Add Linear Vector", lambda: self.inject_vector_shape_to_slide_db("line"))
        self.inject_small_stacked_action(285, 32, "Purge All Drawn Shapes", self.execute_clear_all_shapes_context, variant_fg="#A80F1A")
        
        self.inject_small_stacked_action(435, 6, "Color Fill: Blue", lambda: self.modify_selected_shape_color_attributes("fill", "#D4E2F3"))
        self.inject_small_stacked_action(435, 32, "Color Fill: Orange", lambda: self.modify_selected_shape_color_attributes("fill", "#FFF4CE"))

        self.generate_ribbon_divider_accent(545, "Media Framework", 110)
        self.inject_large_action_node(555, "Movie\nClip", lambda: messagebox.showinfo("Media Link Pipeline", "Direct Show Filters Interface Media Asset Audio/Video Loader attached."))

    # --- DESIGN TAB CONTROL GROUP ---
    def assemble_design_tab_commands(self):
        self.generate_ribbon_divider_accent(5, "Page Setup Matrix", 130)
        self.inject_small_stacked_action(15, 6, "Aspect Ratio 4:3", lambda: messagebox.showinfo("Display System", "Slide boundary constrained internally to standard 4:3 canvas metrics."))
        self.inject_small_stacked_action(15, 32, "Orientation: Landscape", lambda: True)
        
        self.generate_ribbon_divider_accent(145, "Themes Library Registry (Select Theme)", 450)
        # Dynamic Office 2007 Theme Swappers Configuration Map
        self.inject_small_stacked_action(155, 6, "Luna Blue Palette Theme", lambda: self.execute_theme_hot_swap_pipeline("Luna Blue"))
        self.inject_small_stacked_action(155, 32, "Classic Silver Theme Metallic", lambda: self.execute_theme_hot_swap_pipeline("Classic Silver"))
        self.inject_small_stacked_action(355, 6, "Obsidian Black Theme Dark", lambda: self.execute_theme_hot_swap_pipeline("Obsidian Black"))
        self.inject_small_stacked_action(355, 32, "Forest Green Theme Bio", lambda: self.execute_theme_hot_swap_pipeline("Forest Green"))

    # --- ANIMATIONS TAB CONTROL GROUP ---
    def assemble_animations_tab_commands(self):
        self.generate_ribbon_divider_accent(5, "Transitions Simulator Configuration Pipeline", 460)
        self.inject_small_stacked_action(15, 6, "Transition Mode: None", lambda: self.modify_slide_transition_meta_token("None"))
        self.inject_small_stacked_action(15, 32, "Transition Mode: Fade In", lambda: self.modify_slide_transition_meta_token("Fade"))
        self.inject_small_stacked_action(225, 6, "Transition Mode: Slide Left", lambda: self.modify_slide_transition_meta_token("Slide Left"))
        self.inject_small_stacked_action(225, 32, "Transition Mode: Zoom Pop", lambda: self.modify_slide_transition_meta_token("Zoom"))
        
        # Display Status Flag Indicators inside the Ribbon Panel Core itself
        active_transition = self.presentation_database[self.active_slide_ptr].get("transition", "None")
        lbl_info = tk.Label(self.ribbon_tool_belt, text=f"Active Trans: {active_transition}", font=("Segoe UI", 9, "bold"), bg=self.CLR_RIBBON_BODY, fg="#A80F1A")
        lbl_info.place(x=485, y=18)

    # --- SLIDE SHOW TAB CONTROL GROUP ---
    def assemble_slideshow_tab_commands(self):
        self.generate_ribbon_divider_accent(5, "Initialize Production Presentation Stream", 360)
        self.inject_large_action_node(15, "From\nBeginning", lambda: self.deploy_fullscreen_slideshow_stream(0))
        self.inject_large_action_node(95, "From\nCurrent", lambda: self.deploy_fullscreen_slideshow_stream(self.active_slide_ptr))
        
        self.inject_small_stacked_action(175, 6, "Shortcut F5 Trigger", lambda: self.deploy_fullscreen_slideshow_stream(0))
        self.inject_small_stacked_action(175, 32, "Resolution Status: Full Monitor", lambda: True)

    # =========================================================================
    # 5. INDUSTRIAL GRAPHICS VECTOR RENDERERS UTILITY ROUTINES
    # =========================================================================
    def generate_ribbon_divider_accent(self, anchor_x, caption_string, layout_width):
        lbl_group = tk.Label(self.ribbon_tool_belt, text=caption_string, font=("Segoe UI", 8), bg=self.CLR_RIBBON_BODY, fg="#4C607A")
        lbl_group.place(x=anchor_x + (layout_width // 2) - 45, y=58)
        
        canvas_bar = tk.Canvas(self.ribbon_tool_belt, width=2, height=64, bg=self.CLR_RIBBON_BODY, bd=0, highlightthickness=0)
        canvas_bar.place(x=anchor_x + layout_width, y=2)
        canvas_bar.create_line(0, 0, 0, 64, fill="#A5BCDB")

    def inject_large_action_node(self, anchor_x, caption_string, click_routing_cmd):
        btn_frame = tk.Frame(self.ribbon_tool_belt, bg=self.CLR_RIBBON_BODY, bd=1, relief=tk.FLAT)
        btn_frame.place(x=anchor_x, y=4, width=60, height=52)
        
        ico_canvas = tk.Canvas(btn_frame, width=24, height=22, bg=self.CLR_RIBBON_BODY, bd=0, highlightthickness=0)
        ico_canvas.pack(side=tk.TOP, pady=2)
        ico_canvas.create_rectangle(4, 3, 20, 19, fill="#FFFFFF", outline=self.CLR_TEXT_DEEP, width=1.5)
        ico_canvas.create_line(7, 7, 17, 7, fill=self.CLR_TEXT_DEEP)
        ico_canvas.create_line(7, 11, 14, 11, fill=self.CLR_TEXT_DEEP)
        
        lbl_txt = tk.Label(btn_frame, text=caption_string, font=("Segoe UI", 7), bg=self.CLR_RIBBON_BODY, fg=self.CLR_TEXT_DEEP, justify=tk.CENTER)
        lbl_txt.pack(side=tk.TOP)
        
        for visual_node in [btn_frame, ico_canvas, lbl_txt]:
            visual_node.bind("<Enter>", lambda e: btn_frame.configure(bg=self.CLR_HOVER_YELLOW, bd=1, relief=tk.SOLID, highlightbackground=self.CLR_ACCENT_ORANGE))
            visual_node.bind("<Leave>", lambda e: btn_frame.configure(bg=self.CLR_RIBBON_BODY, bd=1, relief=tk.FLAT))
            visual_node.bind("<Button-1>", lambda e: click_routing_cmd())

    def inject_small_stacked_action(self, anchor_x, anchor_y, caption_string, click_routing_cmd, variant_fg=None):
        fg_color = variant_fg if variant_fg else self.CLR_TEXT_DEEP
        lbl_action_node = tk.Label(self.ribbon_tool_belt, text=f"• {caption_string}", font=("Segoe UI", 8), bg=self.CLR_RIBBON_BODY, fg=fg_color, anchor=tk.W, cursor="hand2")
        lbl_action_node.place(x=anchor_x, y=anchor_y, width=140, height=22)
        
        lbl_action_node.bind("<Enter>", lambda e: lbl_action_node.configure(bg=self.CLR_HOVER_YELLOW))
        lbl_action_node.bind("<Leave>", lambda e: lbl_action_node.configure(bg=self.CLR_RIBBON_BODY))
        lbl_action_node.bind("<Button-1>", lambda e: click_routing_cmd())

    # =========================================================================
    # 6. DIALOG MANAGEMENT SYSTEM (THE CORE FIND / REPLACE MODAL SUBWINDOW)
    # =========================================================================
    def deploy_precision_search_modal(self):
        modal_root = tk.Toplevel(self.root)
        modal_root.title("Find and Replace Engine")
        modal_root.geometry("440x190")
        modal_root.resizable(False, False)
        modal_root.configure(bg=self.CLR_SIDEBAR_BG)
        modal_root.transient(self.root)
        modal_root.grab_set()
        
        tk.Label(modal_root, text="Find String Sequence:", font=("Segoe UI", 9, "bold"), bg=self.CLR_SIDEBAR_BG, fg=self.CLR_TEXT_DEEP).place(x=20, y=30)
        entry_find = tk.Entry(modal_root, font=("Segoe UI", 10), bd=1, relief=tk.SOLID)
        entry_find.place(x=150, y=28, width=170, height=22)
        
        tk.Label(modal_root, text="Replace Mapping Target:", font=("Segoe UI", 9), bg=self.CLR_SIDEBAR_BG, fg=self.CLR_TEXT_DEEP).place(x=20, y=70)
        entry_replace = tk.Entry(modal_root, font=("Segoe UI", 10), bd=1, relief=tk.SOLID)
        entry_replace.place(x=150, y=68, width=170, height=22)

        def search_operation_pipeline():
            find_token = entry_find.get()
            if not find_token: return
            
            flag_found = False
            for target_idx, slide_node in enumerate(self.presentation_database):
                if find_token.lower() in slide_node["title"].lower() or find_token.lower() in slide_node["subtitle"].lower():
                    self.route_active_slide_viewport(target_idx)
                    flag_found = True
                    messagebox.showinfo("Search Routine Indexer", f"Parameter match successful inside Presentation Slide Element Index {target_idx + 1}.")
                    break
            if not flag_found:
                messagebox.showwarning("Search Boundary Alert", "Specified query token pattern string parameter could not be identified inside presentation registry.")

        def transform_operation_pipeline():
            find_token = entry_find.get()
            replace_token = entry_replace.get()
            if not find_token: return
            
            current_slide = self.presentation_database[self.active_slide_ptr]
            if find_token.lower() in current_slide["title"].lower():
                current_slide["title"] = current_slide["title"].replace(find_token, replace_token)
            if find_token.lower() in current_slide["subtitle"].lower():
                current_slide["subtitle"] = current_slide["subtitle"].replace(find_token, replace_token)
                
            self.load_active_slide_to_viewport()
            self.rebuild_slide_sorter_rail()
            messagebox.showinfo("Matrix Pipeline Processed", "String mutation completed internally within the currently active slide scope domain layout.")
            modal_root.destroy()

        btn_find = tk.Button(modal_root, text="Find Next", font=("Segoe UI", 9), bg=self.CLR_RIBBON_BODY, fg=self.CLR_TEXT_DEEP, command=search_operation_pipeline)
        btn_find.place(x=340, y=26, width=80, height=24)
        
        btn_rep = tk.Button(modal_root, text="Replace", font=("Segoe UI", 9), bg=self.CLR_RIBBON_BODY, fg=self.CLR_TEXT_DEEP, command=transform_operation_pipeline)
        btn_rep.place(x=340, y=66, width=80, height=24)
        
        btn_exit = tk.Button(modal_root, text="Close", font=("Segoe UI", 9), bg=self.CLR_RIBBON_BODY, fg=self.CLR_TEXT_DEEP, command=modal_root.destroy)
        btn_exit.place(x=340, y=130, width=80, height=24)

    def deploy_office_orb_dropdown_menu(self):
        orb_menu = tk.Toplevel(self.root)
        orb_menu.title("Office Menu")
        orb_menu.geometry("180x220")
        orb_menu.overrideredirect(True) # Remove traditional window controls to simulate native popup drop overlays
        
        # Positions popup context coordinates precisely aligned underneath the circular Orb Canvas Anchor node
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        orb_menu.geometry(f"180x220+{root_x + 10}+{root_y + 90}")
        orb_menu.configure(bg=self.CLR_RIBBON_BODY, bd=2, relief=tk.RAISED)
        
        def close_orb_and_alert(action_name):
            orb_menu.destroy()
            messagebox.showinfo("Office Operations Registry", f"File System Stream Handling Task [{action_name}] invoked successfully.")

        operations_manifest = ["New Presentation", "Open Presentation", "Save Document", "Print Slide Deck", "Publish Packages", "Close Program System"]
        for option_name in operations_manifest:
            lbl_row = tk.Label(orb_menu, text=option_name, font=("Segoe UI", 10), bg=self.CLR_RIBBON_BODY, fg=self.CLR_TEXT_DEEP, anchor=tk.W, padx=15, pady=4)
            lbl_row.pack(fill=tk.X)
            lbl_row.bind("<Enter>", lambda e, r=lbl_row: r.configure(bg=self.CLR_HOVER_YELLOW))
            lbl_row.bind("<Leave>", lambda e, r=lbl_row: r.configure(bg=self.CLR_RIBBON_BODY))
            
            if "Close Program" in option_name:
                lbl_row.bind("<Button-1>", lambda e: self.root.quit())
            else:
                lbl_row.bind("<Button-1>", lambda e, name=option_name: close_orb_and_alert(name))
                
        # Closes target dropdown modal window automatically if focus drifts elsewhere on screen bounds matrix
        orb_menu.bind("<FocusOut>", lambda e: orb_menu.destroy())

    # =========================================================================
    # 7. PRESENTATION RUNTIME CORE: ENGINE PIPELINES & VECTOR LOGICS
    # =========================================================================
    def rebuild_slide_sorter_rail(self):
        # Clear out previous widgets
        self.sidebar_scroll_canvas.delete("all")
        for active_child in self.sidebar_scroll_canvas.winfo_children():
            active_child.destroy()
            
        container_frame = tk.Frame(self.sidebar_scroll_canvas, bg=self.CLR_SIDEBAR_BG)
        container_frame.pack(fill=tk.BOTH, expand=True)
        
        for data_index, slide_packet in enumerate(self.presentation_database):
            is_active_target = (data_index == self.active_slide_ptr)
            border_color_highlight = self.CLR_ACCENT_ORANGE if is_active_target else "#7F9DB9"
            
            card_shell_box = tk.Frame(container_frame, bg=self.CLR_SIDEBAR_BG, height=75)
            card_shell_box.pack(fill=tk.X, padx=10, pady=6)
            card_shell_box.pack_propagate(False)
            
            lbl_idx = tk.Label(card_shell_box, text=str(data_index + 1), font=("Segoe UI", 10, "bold"), bg=self.CLR_SIDEBAR_BG, fg=self.CLR_TEXT_DEEP)
            lbl_idx.pack(side=tk.LEFT, padx=(0, 4))
            
            miniature_wireframe_sheet = tk.Frame(card_shell_box, bg="#FFFFFF", bd=2 if is_active_target else 1, relief=tk.SOLID, highlightbackground=border_color_highlight)
            miniature_wireframe_sheet.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            mini_preview_text = slide_packet["title"].replace("\n", " ")
            if len(mini_preview_text) > 18:
                mini_preview_text = mini_preview_text[:16] + "..."
                
            lbl_preview = tk.Label(miniature_wireframe_sheet, text=mini_preview_text if mini_preview_text.strip() else f"[Slide {data_index+1}]", font=("Segoe UI", 8), bg="#FFFFFF", fg="#555555", anchor=tk.NW, justify=tk.LEFT, padx=3, pady=3)
            lbl_preview.pack(fill=tk.BOTH, expand=True)
            
            for operational_trigger_node in [miniature_wireframe_sheet, lbl_preview]:
                operational_trigger_node.bind("<Button-1>", lambda e, key_index=data_index: self.route_active_slide_viewport(key_index))

        self.lbl_status_indexer.configure(text=f"Slide {self.active_slide_ptr + 1} of {len(self.presentation_database)}")

    def route_active_slide_viewport(self, target_data_index):
        self.active_slide_ptr = target_data_index
        self.load_active_slide_to_viewport()
        self.rebuild_slide_sorter_rail()
        if self.active_tab_pointer == "Animations":
            self.refresh_ribbon_tab_view("Animations") # Dynamic live updates for status checks

    def load_active_slide_to_viewport(self):
        active_slide_data_packet = self.presentation_database[self.active_slide_ptr]
        
        # Block Text Synchronization loops during population phase
        self.input_text_headline.delete("1.0", tk.END)
        self.input_text_headline.insert("1.0", active_slide_data_packet["title"])
        
        self.input_text_subline.delete("1.0", tk.END)
        self.input_text_subline.insert("1.0", active_slide_data_packet["subtitle"])
        
        self.input_speaker_notes_field.delete("1.0", tk.END)
        self.input_speaker_notes_field.insert("1.0", active_slide_data_packet.get("notes", ""))
        
        # Complete clear cycle and structural refresh of Vector Layout Canvas elements layer
        self.slide_graphics_canvas.delete("all")
        for shape_blueprint in active_slide_data_packet["shapes"]:
            geometry_type = shape_blueprint["type"]
            c = shape_blueprint["coords"]
            stroke_clr = shape_blueprint["color"]
            fill_clr = shape_blueprint["fill_color"]
            
            if geometry_type == "rectangle":
                self.slide_graphics_canvas.create_rectangle(c[0], c[1], c[2], c[3], outline=stroke_clr, fill=fill_clr, width=2)
            elif geometry_type == "oval":
                self.slide_graphics_canvas.create_oval(c[0], c[1], c[2], c[3], outline=stroke_clr, fill=fill_clr, width=2)
            elif geometry_type == "line":
                self.slide_graphics_canvas.create_line(c[0], c[1], c[2], c[3], fill=stroke_clr, width=3)

    def synchronize_workspace_inputs_to_memory(self, event_arg=None):
        extracted_headline_string = self.input_text_headline.get("1.0", tk.END).strip("\n")
        extracted_subline_string = self.input_text_subline.get("1.0", tk.END).strip("\n")
        extracted_notes_string = self.input_speaker_notes_field.get("1.0", tk.END).strip("\n")
        
        self.presentation_database[self.active_slide_ptr]["title"] = extracted_headline_string
        self.presentation_database[self.active_slide_ptr]["subtitle"] = extracted_subline_string
        self.presentation_database[self.active_slide_ptr]["notes"] = extracted_notes_string
        
        # Dynamic continuous thumbnail synchronization callback invocation loop
        self.rebuild_slide_sorter_rail()

    # =========================================================================
    # 8. PRESENTATION RUNTIME ENGINE OPERATIONS COMMAND FUNCTIONS IMPLEMENTATION
    # =========================================================================
    def execute_append_new_slide_instance(self):
        new_clean_blueprint = {"title": "Click to add title", "subtitle": "Click to add subtitle", "shapes": [], "notes": "", "transition": "None"}
        self.presentation_database.insert(self.active_slide_ptr + 1, new_clean_blueprint)
        self.active_slide_ptr += 1
        self.load_active_slide_to_viewport()
        self.rebuild_slide_sorter_rail()

    def execute_purge_selected_slide_instance(self):
        if len(self.presentation_database) <= 1:
            messagebox.showerror("Structural Array Exception", "PowerPoint Application Workspace requires keeping at least 1 Presentation Active Slide node.")
            return
        self.presentation_database.pop(self.active_slide_ptr)
        if self.active_slide_ptr >= len(self.presentation_database):
            self.active_slide_ptr = len(self.presentation_database) - 1
        self.load_active_slide_to_viewport()
        self.rebuild_slide_sorter_rail()

    def execute_reset_canvas_text_fields(self):
        self.input_text_headline.delete("1.0", tk.END)
        self.input_text_headline.insert("1.0", "Click to add title")
        self.input_text_subline.delete("1.0", tk.END)
        self.input_text_subline.insert("1.0", "Click to add subtitle")
        self.synchronize_workspace_inputs_to_memory()

    def execute_text_formatting_processor(self, dynamic_action_token, configuration_arg=None):
        try:
            target_field = self.input_text_headline # Defaults operations stream mapping context directly to Title box scope
            headline_font = tkfont.Font(font=target_field['font'])
            
            if dynamic_action_token == "bold":
                target_field.tag_add("bold_matrix", "sel.first", "sel.last")
                target_field.tag_config("bold_matrix", font=("Segoe UI", 28, "bold"))
            elif dynamic_action_token == "italic":
                target_field.tag_add("italic_matrix", "sel.first", "sel.last")
                target_field.tag_config("italic_matrix", font=("Segoe UI", 28, "italic"))
            elif dynamic_action_token == "underline":
                target_field.tag_add("underline_matrix", "sel.first", "sel.last")
                target_field.tag_config("underline_matrix", underline=True)
            elif dynamic_action_token == "color":
                tag_id = f"clr_{configuration_arg.replace('#','')}"
                target_field.tag_add(tag_id, "sel.first", "sel.last")
                target_field.tag_config(tag_id, foreground=configuration_arg)
            elif "align_" in dynamic_action_token:
                alignment_rule = dynamic_action_token.split("_")[1]
                target_field.tag_add("alignment_rule", "1.0", tk.END)
                target_field.tag_config("alignment_rule", justify=alignment_rule)
        except tk.TclError:
            pass # Suppresses operations exceptions cleanly during zero cursor string selection anomalies

    def execute_clipboard_paste_simulation(self):
        self.input_text_headline.delete("1.0", tk.END)
        self.input_text_headline.insert("1.0", "Raju Production System Buffer Core Sync Live")
        self.synchronize_workspace_inputs_to_memory()

    def execute_theme_hot_swap_pipeline(self, target_theme_token):
        self.current_theme_token = target_theme_token
        self.apply_theme_vectors()
        
        # Redeply Application Frame Layout Skin elements dynamically without destroying state registry memory
        self.root.configure(bg=self.CLR_OUTER_BG)
        self.ribbon_apron.configure(bg=self.CLR_LUNA_BLUE)
        self.orb_canvas.configure(bg=self.CLR_LUNA_BLUE)
        self.sidebar_rail_track.configure(bg=self.CLR_SIDEBAR_BG)
        self.sidebar_scroll_canvas.configure(bg=self.CLR_SIDEBAR_BG)
        self.slide_sheet_board.configure(bg=self.CLR_SLIDE_BG)
        self.slide_graphics_canvas.configure(bg=self.CLR_SLIDE_BG)
        self.speaker_notes_panel.configure(bg=self.CLR_SIDEBAR_BG)
        self.status_footer_line.configure(bg="#CBD9EC")
        
        self.lbl_status_theme_flag.configure(text=f"Active Theme: {self.current_theme_token}")
        self.refresh_ribbon_tab_view(self.active_tab_pointer)
        self.load_active_slide_to_viewport()
        self.rebuild_slide_sorter_rail()

    def modify_slide_transition_meta_token(self, target_transition_mode):
        self.presentation_database[self.active_slide_ptr]["transition"] = target_transition_mode
        self.refresh_ribbon_tab_view("Animations")

    # =========================================================================
    # 9. INTERACTIVE ILLUSTRATIONS VECTOR ENGINE - SHAPES HANDLERS
    # =========================================================================
    def inject_vector_shape_to_slide_db(self, shape_type_selector):
        shapes_list = self.presentation_database[self.active_slide_ptr]["shapes"]
        loop_offset = (len(shapes_list) * 20) % 160
        
        calculated_coords = [60 + loop_offset, 380 + (loop_offset // 2), 240 + loop_offset, 470 + (loop_offset // 2)]
        new_shape_node = {
            "type": shape_type_selector,
            "coords": calculated_coords,
            "color": self.CLR_TEXT_DEEP,
            "fill_color": self.CLR_RIBBON_BODY if shape_type_selector != "line" else ""
        }
        shapes_list.append(new_shape_node)
        self.load_active_slide_to_viewport()

    def execute_clear_all_shapes_context(self):
        self.presentation_database[self.active_slide_ptr]["shapes"] = []
        self.load_active_slide_to_viewport()

    def modify_selected_shape_color_attributes(self, target_attribute, input_color_hex):
        if self.selected_shape_index is None:
            messagebox.showwarning("Selection Anchor Missing", "Please select a vector shape element object on the whiteboard slide sheet canvas area grid first.")
            return
        shapes_array = self.presentation_database[self.active_slide_ptr]["shapes"]
        if target_attribute == "fill" and shapes_array[self.selected_shape_index]["type"] != "line":
            shapes_array[self.selected_shape_index]["fill_color"] = input_color_hex
        self.load_active_slide_to_viewport()

    # Precise Hit-Box Intersection Detection algorithms for Canvas Objects Vector Arrays
    def execute_mouse_click_shape_detection(self, event_arg):
        mouse_x = event_arg.x
        mouse_y = event_arg.y
        shapes_array = self.presentation_database[self.active_slide_ptr]["shapes"]
        
        self.selected_shape_index = None
        # Traverses loop inverted backwards to anchor target interaction onto the highest visual layer stack
        for idx in reversed(range(len(shapes_array))):
            coords = shapes_array[idx]["coords"]
            if coords[0] <= mouse_x <= coords[2] and coords[1] <= mouse_y <= coords[4] if len(coords)>3 else coords[3]:
                # Special boundary bounding box checker
                if mouse_x >= coords[0] and mouse_x <= coords[2] and mouse_y >= coords[1] and mouse_y <= coords[3]:
                    self.selected_shape_index = idx
                    self.drag_start_coordinates = (mouse_x, mouse_y)
                    # Provides instantaneous dynamic selection border visualization accent line indicators
                    self.slide_graphics_canvas.create_rectangle(coords[0]-3, coords[1]-3, coords[2]+3, coords[3]+3, outline=self.CLR_ACCENT_ORANGE, dash=(3,3), width=1, tag="selection_bounding_box")
                    break

    def execute_mouse_drag_shape_repositioning(self, event_arg):
        if self.selected_shape_index is None: return
        
        current_mouse_x = event_arg.x
        current_mouse_y = event_arg.y
        
        delta_x = current_mouse_x - self.drag_start_coordinates[0]
        delta_y = current_mouse_y - self.drag_start_coordinates[1]
        
        shapes_array = self.presentation_database[self.active_slide_ptr]["shapes"]
        coords = shapes_array[self.selected_shape_index]["coords"]
        
        # Real-time modification updates to structure coordinate limits parameters
        shapes_array[self.selected_shape_index]["coords"] = [
            coords[0] + delta_x, coords[1] + delta_y,
            coords[2] + delta_x, coords[3] + delta_y
        ]
        
        self.drag_start_coordinates = (current_mouse_x, current_mouse_y)
        self.load_active_slide_to_viewport()

    def finalize_shape_transformation_pipeline(self):
        self.slide_graphics_canvas.delete("selection_bounding_box")

    # =========================================================================
    # 10. ADVANCED PRODUCTION UNIT: FULLSCREEN ENGINE SLIDE SHOW MODULE
    # =========================================================================
    def deploy_fullscreen_slideshow_stream(self, initialization_index_ptr):
        self.is_slideshow_active = True
        self.slideshow_ptr = initialization_index_ptr
        
        # Builds autonomous separate graphic sub-window layer covering desktop bounds completely
        self.show_root = tk.Toplevel(self.root)
        self.show_root.attributes("-fullscreen", True)
        self.show_root.configure(bg="#000000")
        
        self.show_canvas = tk.Canvas(self.show_root, bg="#000000", highlightthickness=0, bd=0)
        self.show_canvas.pack(fill=tk.BOTH, expand=True)
        
        # Input Key Event Listeners mapping for Presentation Navigation operations
        self.show_root.bind("<Escape>", lambda e: self.terminate_slideshow_stream())
        self.show_root.bind("<Button-1>", lambda e: self.navigate_slideshow_direction_step(1))
        self.show_root.bind("<Right>", lambda e: self.navigate_slideshow_direction_step(1))
        self.show_root.bind("<Left>", lambda e: self.navigate_slideshow_direction_step(-1))
        
        self.render_slideshow_presentation_frame()

    def render_slideshow_presentation_frame(self):
        self.show_canvas.delete("all")
        scr_width = self.show_canvas.winfo_screenwidth()
        scr_height = self.show_canvas.winfo_screenheight()
        
        slide_node = self.presentation_database[self.slideshow_ptr]
        transition_type = slide_node.get("transition", "None")
        
        # Simulates transition animations algorithm logic inside tk canvas
        if transition_type == "Zoom":
            for step in range(1, 6):
                self.show_root.after(step * 30, lambda s=step: self.draw_slideshow_elements_matrix(slide_node, scr_width, scr_height, scale_factor=s/5.0))
        elif transition_type == "Fade":
            self.show_canvas.configure(bg="#333333")
            self.show_root.after(80, lambda: self.show_canvas.configure(bg="#000000"))
            self.show_root.after(100, lambda: self.draw_slideshow_elements_matrix(slide_node, scr_width, scr_height))
        else:
            self.draw_slideshow_elements_matrix(slide_node, scr_width, scr_height)

    def draw_slideshow_elements_matrix(self, slide_node, scr_w, scr_h, scale_factor=1.0):
        self.show_canvas.delete("elements")
        
        # Draws clean uniform aspect ratio adjusted backdrop presentation plate
        plate_w = 1024 * scale_factor
        plate_h = 768 * scale_factor
        offset_x = (scr_w - plate_w) // 2
        offset_y = (scr_h - plate_h) // 2
        
        # Back sheet presentation layer canvas drawing
        self.show_canvas.create_rectangle(offset_x, offset_y, offset_x + plate_w, offset_y + plate_h, fill=self.CLR_SLIDE_BG, outline="#333333", tag="elements")
        
        title_font_size = int(44 * scale_factor) if int(44 * scale_factor) > 8 else 9
        subtitle_font_size = int(22 * scale_factor) if int(22 * scale_factor) > 6 else 7
        
        # Text string layouts projection pipelines
        self.show_canvas.create_text(offset_x + (plate_w // 2), offset_y + (150 * scale_factor), text=slide_node["title"], font=("Segoe UI", title_font_size, "bold"), fill=self.CLR_TEXT_DEEP, justify=tk.CENTER, tag="elements", width=plate_w - 100)
        self.show_canvas.create_text(offset_x + (plate_w // 2), offset_y + (340 * scale_factor), text=slide_node["subtitle"], font=("Segoe UI", subtitle_font_size), fill="#555555" if self.current_theme_token != "Obsidian Black" else "#CCCCCC", justify=tk.CENTER, tag="elements", width=plate_w - 150)
        
        # Render Associated Vectors onto the Presentation Screen layer matrix dynamically
        for shape in slide_node["shapes"]:
            c = shape["coords"]
            t = shape["type"]
            s_clr = shape["color"]
            f_clr = shape["fill_color"]
            
            # Normalizes coordinates to fit presentation resolution matrix scales smoothly
            x1 = offset_x + (c[0] * (plate_w / 720.0))
            y1 = offset_y + (c[1] * (plate_h / 540.0))
            x2 = offset_x + (c[2] * (plate_w / 720.0))
            y2 = offset_y + (c[3] * (plate_h / 540.0))
            
            if t == "rectangle":
                self.show_canvas.create_rectangle(x1, y1, x2, y2, outline=s_clr, fill=f_clr, width=3, tag="elements")
            elif t == "oval":
                self.show_canvas.create_oval(x1, y1, x2, y2, outline=s_clr, fill=f_clr, width=3, tag="elements")
            elif t == "line":
                self.show_canvas.create_line(x1, y1, x2, y2, fill=s_clr, width=4, tag="elements")

    def navigate_slideshow_direction_step(self, step_delta_value):
        self.slideshow_ptr += step_delta_value
        if self.slideshow_ptr < 0:
            self.slideshow_ptr = 0
        elif self.slideshow_ptr >= len(self.presentation_database):
            self.terminate_slideshow_stream()
            return
        self.render_slideshow_presentation_frame()

    def terminate_slideshow_stream(self):
        self.is_slideshow_active = False
        if hasattr(self, "show_root"):
            self.show_root.destroy()

if __name__ == "__main__":
    app_engine_kernel_root = tk.Tk()
    
    # Anti-crash styling window initialization block
    app_engine_kernel_root.configure(bg="#2b3e5a")
    
    production_app_instance = MicrosoftPowerPoint2007Ultimate(app_engine_kernel_root)
    app_engine_kernel_root.mainloop()
