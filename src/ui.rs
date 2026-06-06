use crate::chm_doc::ChmDocumentationProvider;
use crate::idlgen::{EnumItemInfo, MethodInfo, TypeLibInfo};
use crossterm::{
    event::{
        self, DisableMouseCapture, EnableMouseCapture, Event, KeyCode, KeyEventKind, KeyModifiers,
        MouseEvent, MouseEventKind,
    },
    execute,
    terminal::{EnterAlternateScreen, LeaveAlternateScreen, disable_raw_mode, enable_raw_mode},
};
use ratatui::{
    Terminal,
    backend::CrosstermBackend,
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span, Text},
    widgets::{
        Block, BorderType, Borders, Cell, Clear, List, ListItem, ListState, Paragraph, Row,
        Scrollbar, ScrollbarOrientation, ScrollbarState, Table, TableState, Wrap,
    },
};
use std::{error::Error, io, path::PathBuf};

// ─── Color palette ───────────────────────────────────────────────────────────
const BG_PRIMARY: Color = Color::Indexed(235);

const FG_PRIMARY: Color = Color::Indexed(252);
const FG_SECONDARY: Color = Color::Indexed(248);
const FG_DIM: Color = Color::Indexed(242);
const ACCENT_CYAN: Color = Color::Indexed(44);
const ACCENT_YELLOW: Color = Color::Indexed(186);
const ACCENT_MAGENTA: Color = Color::Indexed(141);
const ACCENT_GREEN: Color = Color::Indexed(114);
const ACCENT_RED: Color = Color::Indexed(203);
const ACCENT_BLUE: Color = Color::Indexed(75);
const BORDER_FOCUSED: Color = Color::Indexed(75);
const BORDER_UNFOCUSED: Color = Color::Indexed(240);
const SELECTION_BG: Color = Color::Indexed(24);
const TITLE_BAR_BG: Color = Color::Indexed(236);
const FOOTER_BG: Color = Color::Indexed(236);

#[derive(PartialEq)]
enum ViewMode {
    Idl,
    Structured,
}

#[derive(PartialEq)]
enum SearchTarget {
    Types,
    Members,
}

#[derive(PartialEq)]
enum Focus {
    TypeList,
    MethodList,
    Details,
    IdlView,
}

struct SearchItem {
    type_index: usize,
    type_name: String,
    member_name: String,
    kind: String, // "Method" or "Enum"
}

struct App {
    type_lib_info: TypeLibInfo,
    doc_provider: Option<ChmDocumentationProvider>,
    types: Vec<(String, String)>,                 // Name, Kind
    filtered_types: Vec<(usize, String, String)>, // Original Index, Name, Kind
    list_state: ListState,
    list_scroll_state: ScrollbarState, // Scrollbar for Type List
    current_idl: String,
    idl_scroll_offset: u16,
    idl_scroll_state: ScrollbarState,
    current_methods: Vec<MethodInfo>,
    current_enums: Vec<EnumItemInfo>,
    search_query: String,
    member_search_query: String,
    view_mode: ViewMode,
    search_target: SearchTarget,
    focus: Focus,
    method_list_state: ListState, // State for Method List (Middle Panel)
    method_list_scroll_state: ScrollbarState, // Scrollbar for Method List
    details_scroll_offset: u16,   // Scroll offset for Details Panel
    details_scroll_state: ScrollbarState, // Scrollbar for Details Panel
    content_table_state: TableState, // Kept for Enums
    content_scroll_state: ScrollbarState, // Kept for Enums
    // Global Search
    all_search_items: Vec<SearchItem>,
    show_global_search: bool,
    global_search_query: String,
    global_search_results: Vec<usize>, // Indices into all_search_items
    global_search_state: ListState,
    global_search_scroll_state: ScrollbarState, // Scrollbar for Global Search
    show_exit_confirmation: bool,
    show_help: bool,
}

impl App {
    fn new(tlb_path: PathBuf, chm_path: Option<String>) -> Result<Self, Box<dyn Error>> {
        let mut type_lib_info = TypeLibInfo::new();
        type_lib_info.load_type_lib(&tlb_path)?;

        let count = type_lib_info.get_type_info_count();
        let mut types = Vec::new();
        let mut all_search_items = Vec::new();

        for i in 0..count {
            if let Ok((name, kind)) = type_lib_info.get_type_name_and_kind(i) {
                types.push((name.clone(), kind.clone()));

                // Index the type itself
                all_search_items.push(SearchItem {
                    type_index: i as usize,
                    type_name: name.clone(),
                    member_name: name.clone(), // For types, member_name is the same as type_name
                    kind: kind.clone(),
                });

                // Pre-index methods
                if let Ok(methods) = type_lib_info.get_type_methods(i) {
                    for method in methods {
                        all_search_items.push(SearchItem {
                            type_index: i as usize,
                            type_name: name.clone(),
                            member_name: method.name,
                            kind: "Method".to_string(),
                        });
                    }
                }

                // Pre-index enums
                if let Ok(enums) = type_lib_info.get_type_enums(i) {
                    for item in enums {
                        all_search_items.push(SearchItem {
                            type_index: i as usize,
                            type_name: name.clone(),
                            member_name: item.name,
                            kind: "EnumValue".to_string(),
                        });
                    }
                }
            }
        }

        let doc_provider = if let Some(path) = chm_path {
            ChmDocumentationProvider::new(&path).ok()
        } else {
            None
        };

        let mut app = App {
            type_lib_info,
            doc_provider,
            types,
            filtered_types: Vec::new(),
            list_state: ListState::default(),
            list_scroll_state: ScrollbarState::default(),
            current_idl: String::new(),
            idl_scroll_offset: 0,
            idl_scroll_state: ScrollbarState::default(),
            current_methods: Vec::new(),
            current_enums: Vec::new(),

            search_query: String::new(),
            member_search_query: String::new(),
            view_mode: ViewMode::Structured,
            search_target: SearchTarget::Types,
            focus: Focus::TypeList,
            method_list_state: ListState::default(),
            method_list_scroll_state: ScrollbarState::default(),
            details_scroll_offset: 0,
            details_scroll_state: ScrollbarState::default(),
            content_table_state: TableState::default(),
            content_scroll_state: ScrollbarState::default(),
            all_search_items,
            show_global_search: false,
            global_search_query: String::new(),
            global_search_results: Vec::new(),
            global_search_state: ListState::default(),
            global_search_scroll_state: ScrollbarState::default(),
            show_exit_confirmation: false,
            show_help: false,
        };
        app.update_filter();
        Ok(app)
    }

    fn update_filter(&mut self) {
        let query = self.search_query.to_lowercase();
        self.filtered_types = self
            .types
            .iter()
            .enumerate()
            .filter(|(_, (name, _))| name.to_lowercase().contains(&query))
            .map(|(i, (name, kind))| (i, name.clone(), kind.clone()))
            .collect();

        self.list_state.select(None);
        self.list_scroll_state = self
            .list_scroll_state
            .content_length(self.filtered_types.len());
        if !self.filtered_types.is_empty() {
            self.list_state.select(Some(0));
            self.update_selection();
        } else {
            self.current_idl.clear();
            self.current_methods.clear();
            self.current_enums.clear();
            self.content_table_state.select(None);
        }
    }

    fn update_selection(&mut self) {
        if let Some(selected_idx) = self.list_state.selected() {
            if let Some((original_idx, _, _)) = self.filtered_types.get(selected_idx) {
                if let Ok(idl) = self.type_lib_info.get_type_idl(*original_idx as u32) {
                    self.current_idl = idl;
                }
                if let Ok(methods) = self.type_lib_info.get_type_methods(*original_idx as u32) {
                    self.current_methods = methods;
                } else {
                    self.current_methods.clear();
                }
                if let Ok(enums) = self.type_lib_info.get_type_enums(*original_idx as u32) {
                    self.current_enums = enums;
                } else {
                    self.current_enums.clear();
                }

                // Reset content selection and scroll
                self.method_list_state.select(None);
                self.method_list_scroll_state = ScrollbarState::default();
                self.details_scroll_offset = 0;
                self.details_scroll_state = ScrollbarState::default();
                self.idl_scroll_offset = 0;
                self.idl_scroll_state = ScrollbarState::default();

                self.content_table_state.select(None);
                self.content_scroll_state = ScrollbarState::default();

                if !self.current_methods.is_empty() {
                    self.method_list_state.select(Some(0));
                } else if !self.current_enums.is_empty() {
                    self.content_table_state.select(Some(0));
                }
            }
        }
    }

    fn next(&mut self) {
        match self.focus {
            Focus::TypeList => {
                let i = match self.list_state.selected() {
                    Some(i) => {
                        if i >= self.filtered_types.len() - 1 {
                            0
                        } else {
                            i + 1
                        }
                    }
                    None => 0,
                };
                self.list_state.select(Some(i));
                self.list_scroll_state = self.list_scroll_state.position(i);
                self.update_selection();
            }
            Focus::MethodList => {
                if !self.current_methods.is_empty() {
                    let i = match self.method_list_state.selected() {
                        Some(i) => {
                            if i >= self.current_methods.len() - 1 {
                                0
                            } else {
                                i + 1
                            }
                        }
                        None => 0,
                    };
                    self.method_list_state.select(Some(i));
                    self.method_list_scroll_state = self.method_list_scroll_state.position(i);
                    // Reset details scroll when changing method
                    self.details_scroll_offset = 0;
                    self.details_scroll_state = ScrollbarState::default();
                } else if !self.current_enums.is_empty() {
                    // Enums use content_table_state (2 panel layout)
                    let i = match self.content_table_state.selected() {
                        Some(i) => {
                            if i >= self.current_enums.len() - 1 {
                                0
                            } else {
                                i + 1
                            }
                        }
                        None => 0,
                    };
                    self.content_table_state.select(Some(i));
                    self.content_scroll_state = self.content_scroll_state.position(i);
                }
            }
            Focus::Details => {
                // Scroll details
                self.details_scroll_offset = self.details_scroll_offset.saturating_add(1);
                self.details_scroll_state = self
                    .details_scroll_state
                    .position(self.details_scroll_offset as usize);
            }
            Focus::IdlView => {
                self.idl_scroll_offset = self.idl_scroll_offset.saturating_add(1);
                self.idl_scroll_state = self
                    .idl_scroll_state
                    .position(self.idl_scroll_offset as usize);
            }
        }
    }

    fn previous(&mut self) {
        match self.focus {
            Focus::TypeList => {
                let i = match self.list_state.selected() {
                    Some(i) => {
                        if i == 0 {
                            self.filtered_types.len() - 1
                        } else {
                            i - 1
                        }
                    }
                    None => 0,
                };
                self.list_state.select(Some(i));
                self.list_scroll_state = self.list_scroll_state.position(i);
                self.update_selection();
            }
            Focus::MethodList => {
                if !self.current_methods.is_empty() {
                    let i = match self.method_list_state.selected() {
                        Some(i) => {
                            if i == 0 {
                                self.current_methods.len() - 1
                            } else {
                                i - 1
                            }
                        }
                        None => 0,
                    };
                    self.method_list_state.select(Some(i));
                    self.method_list_scroll_state = self.method_list_scroll_state.position(i);
                    self.details_scroll_offset = 0;
                    self.details_scroll_state = ScrollbarState::default();
                } else if !self.current_enums.is_empty() {
                    let i = match self.content_table_state.selected() {
                        Some(i) => {
                            if i == 0 {
                                self.current_enums.len() - 1
                            } else {
                                i - 1
                            }
                        }
                        None => 0,
                    };
                    self.content_table_state.select(Some(i));
                    self.content_scroll_state = self.content_scroll_state.position(i);
                }
            }
            Focus::Details => {
                self.details_scroll_offset = self.details_scroll_offset.saturating_sub(1);
                self.details_scroll_state = self
                    .details_scroll_state
                    .position(self.details_scroll_offset as usize);
            }
            Focus::IdlView => {
                self.idl_scroll_offset = self.idl_scroll_offset.saturating_sub(1);
                self.idl_scroll_state = self
                    .idl_scroll_state
                    .position(self.idl_scroll_offset as usize);
            }
        }
    }

    fn next_page(&mut self) {
        let page_size: usize = 10;
        match self.focus {
            Focus::TypeList => {
                let i = match self.list_state.selected() {
                    Some(i) => {
                        let next = i.saturating_add(page_size);
                        if next >= self.filtered_types.len() {
                            self.filtered_types.len() - 1
                        } else {
                            next
                        }
                    }
                    None => 0,
                };
                self.list_state.select(Some(i));
                self.list_scroll_state = self.list_scroll_state.position(i);
                self.update_selection();
            }
            Focus::MethodList => {
                if !self.current_methods.is_empty() {
                    let i = match self.method_list_state.selected() {
                        Some(i) => {
                            let next = i.saturating_add(page_size);
                            if next >= self.current_methods.len() {
                                self.current_methods.len() - 1
                            } else {
                                next
                            }
                        }
                        None => 0,
                    };
                    self.method_list_state.select(Some(i));
                    self.method_list_scroll_state = self.method_list_scroll_state.position(i);
                    self.details_scroll_offset = 0;
                    self.details_scroll_state = ScrollbarState::default();
                } else if !self.current_enums.is_empty() {
                    let i = match self.content_table_state.selected() {
                        Some(i) => {
                            let next = i.saturating_add(page_size);
                            if next >= self.current_enums.len() {
                                self.current_enums.len() - 1
                            } else {
                                next
                            }
                        }
                        None => 0,
                    };
                    self.content_table_state.select(Some(i));
                    self.content_scroll_state = self.content_scroll_state.position(i);
                }
            }
            Focus::Details => {
                self.details_scroll_offset =
                    self.details_scroll_offset.saturating_add(page_size as u16);
                self.details_scroll_state = self
                    .details_scroll_state
                    .position(self.details_scroll_offset as usize);
            }
            Focus::IdlView => {
                self.idl_scroll_offset =
                    self.idl_scroll_offset.saturating_add(page_size as u16);
                self.idl_scroll_state = self
                    .idl_scroll_state
                    .position(self.idl_scroll_offset as usize);
            }
        }
    }

    fn previous_page(&mut self) {
        let page_size: usize = 10;
        match self.focus {
            Focus::TypeList => {
                let i = match self.list_state.selected() {
                    Some(i) => i.saturating_sub(page_size),
                    None => 0,
                };
                self.list_state.select(Some(i));
                self.list_scroll_state = self.list_scroll_state.position(i);
                self.update_selection();
            }
            Focus::MethodList => {
                if !self.current_methods.is_empty() {
                    let i = match self.method_list_state.selected() {
                        Some(i) => i.saturating_sub(page_size),
                        None => 0,
                    };
                    self.method_list_state.select(Some(i));
                    self.method_list_scroll_state = self.method_list_scroll_state.position(i);
                    self.details_scroll_offset = 0;
                    self.details_scroll_state = ScrollbarState::default();
                } else if !self.current_enums.is_empty() {
                    let i = match self.content_table_state.selected() {
                        Some(i) => i.saturating_sub(page_size),
                        None => 0,
                    };
                    self.content_table_state.select(Some(i));
                    self.content_scroll_state = self.content_scroll_state.position(i);
                }
            }
            Focus::Details => {
                self.details_scroll_offset =
                    self.details_scroll_offset.saturating_sub(page_size as u16);
                self.details_scroll_state = self
                    .details_scroll_state
                    .position(self.details_scroll_offset as usize);
            }
            Focus::IdlView => {
                self.idl_scroll_offset =
                    self.idl_scroll_offset.saturating_sub(page_size as u16);
                self.idl_scroll_state = self
                    .idl_scroll_state
                    .position(self.idl_scroll_offset as usize);
            }
        }
    }

    fn toggle_view(&mut self) {
        self.view_mode = match self.view_mode {
            ViewMode::Idl => ViewMode::Structured,
            ViewMode::Structured => ViewMode::Idl,
        };
    }

    fn toggle_search_target(&mut self) {
        self.search_target = match self.search_target {
            SearchTarget::Types => SearchTarget::Members,
            SearchTarget::Members => SearchTarget::Types,
        };
    }

    fn update_global_search(&mut self) {
        let query = self.global_search_query.to_lowercase();
        if query.is_empty() {
            self.global_search_results.clear();
            self.global_search_state.select(None);
            return;
        }

        self.global_search_results = self
            .all_search_items
            .iter()
            .enumerate()
            .filter(|(_, item)| item.member_name.to_lowercase().contains(&query))
            .map(|(i, _)| i)
            .collect();

        if !self.global_search_results.is_empty() {
            self.global_search_state.select(Some(0));
        } else {
            self.global_search_state.select(None);
        }
    }

    fn next_global_result(&mut self) {
        if self.global_search_results.is_empty() {
            return;
        }
        let i = match self.global_search_state.selected() {
            Some(i) => {
                if i >= self.global_search_results.len() - 1 {
                    0
                } else {
                    i + 1
                }
            }
            None => 0,
        };
        self.global_search_state.select(Some(i));
    }

    fn previous_global_result(&mut self) {
        if self.global_search_results.is_empty() {
            return;
        }
        let i = match self.global_search_state.selected() {
            Some(i) => {
                if i == 0 {
                    self.global_search_results.len() - 1
                } else {
                    i - 1
                }
            }
            None => 0,
        };
        self.global_search_state.select(Some(i));
    }

    fn select_global_result(&mut self) {
        if let Some(selected_idx) = self.global_search_state.selected() {
            if let Some(&item_idx) = self.global_search_results.get(selected_idx) {
                let (type_index, member_name, kind) =
                    if let Some(item) = self.all_search_items.get(item_idx) {
                        (item.type_index, item.member_name.clone(), item.kind.clone())
                    } else {
                        return;
                    };

                self.show_global_search = false;
                self.search_query.clear();
                self.update_filter();

                if let Some(pos) = self
                    .filtered_types
                    .iter()
                    .position(|(idx, _, _)| *idx == type_index)
                {
                    self.list_state.select(Some(pos));
                    self.update_selection();
                }

                // If it's a method or enum value, select it in the content table
                if kind == "Method" || kind == "EnumValue" {
                    self.member_search_query = member_name.clone();
                    self.search_target = SearchTarget::Members; // Switch focus to member search so user can see/clear it

                    // Need to find the index of the member in the current list
                    let member_query = member_name.to_lowercase();
                    if !self.current_methods.is_empty() {
                        if let Some(pos) = self
                            .current_methods
                            .iter()
                            .position(|m| m.name.to_lowercase() == member_query)
                        {
                            self.method_list_state.select(Some(pos));
                        }
                    } else if !self.current_enums.is_empty() {
                        if let Some(pos) = self
                            .current_enums
                            .iter()
                            .position(|e| e.name.to_lowercase() == member_query)
                        {
                            self.content_table_state.select(Some(pos));
                        }
                    }
                } else {
                    // It's a type (Interface, Enum, Dispatch, etc.)
                    // We already selected the type in the left panel.
                    // Just ensure we are focusing on the type list and clear member search
                    self.member_search_query.clear();
                    self.search_target = SearchTarget::Types;
                    self.focus = Focus::TypeList;
                }
            }
        }
    }
    fn handle_key_event(&mut self, key: event::KeyEvent) -> bool {
        if key.kind != KeyEventKind::Press {
            return false;
        }

        if self.show_exit_confirmation {
            match key.code {
                KeyCode::Char('y') | KeyCode::Enter => return true,
                KeyCode::Char('n') | KeyCode::Esc => {
                    self.show_exit_confirmation = false;
                }
                _ => {}
            }
        } else if self.show_global_search {
            match key.code {
                KeyCode::Esc => self.show_global_search = false,
                KeyCode::Down => self.next_global_result(),
                KeyCode::Up => self.previous_global_result(),
                KeyCode::Enter => self.select_global_result(),
                KeyCode::Char(c) => {
                    self.global_search_query.push(c);
                    self.update_global_search();
                }
                KeyCode::Backspace => {
                    self.global_search_query.pop();
                    self.update_global_search();
                }
                _ => {}
            }
        } else {
            match key.code {
                KeyCode::Char('c') if key.modifiers.contains(KeyModifiers::CONTROL) => {
                    return true;
                }
                KeyCode::Char('?') | KeyCode::F(1) => {
                    self.show_help = !self.show_help;
                }
                KeyCode::Char('q')
                    if self.search_query.is_empty() && self.member_search_query.is_empty() =>
                {
                    self.show_exit_confirmation = true;
                }
                KeyCode::Esc => {
                    if self.show_help {
                        self.show_help = false;
                    } else {
                        self.show_exit_confirmation = true;
                    }
                }
                KeyCode::Down => self.next(),
                KeyCode::Up => self.previous(),
                KeyCode::PageDown => self.next_page(),
                KeyCode::PageUp => self.previous_page(),
                KeyCode::Right => match self.focus {
                    Focus::TypeList => {
                        if self.view_mode == ViewMode::Idl {
                            self.focus = Focus::IdlView;
                        } else {
                            self.focus = Focus::MethodList;
                        }
                    }
                    Focus::MethodList => {
                        if !self.current_methods.is_empty() {
                            self.focus = Focus::Details;
                        }
                    }
                    _ => {}
                },
                KeyCode::Left => match self.focus {
                    Focus::Details => self.focus = Focus::MethodList,
                    Focus::MethodList => self.focus = Focus::TypeList,
                    Focus::IdlView => self.focus = Focus::TypeList,
                    _ => {}
                },
                KeyCode::Tab | KeyCode::Char('v') => self.toggle_view(),
                KeyCode::Char('p') if key.modifiers.contains(KeyModifiers::CONTROL) => {
                    self.show_global_search = true;
                    self.global_search_query.clear();
                    self.update_global_search();
                }
                KeyCode::Char('f') if key.modifiers.contains(KeyModifiers::CONTROL) => {
                    self.toggle_search_target();
                }
                KeyCode::Enter => {
                    // Enter key logic if needed
                }
                KeyCode::Char(c) => match self.search_target {
                    SearchTarget::Types => {
                        self.search_query.push(c);
                        self.update_filter();
                    }
                    SearchTarget::Members => {
                        self.member_search_query.push(c);
                    }
                },
                KeyCode::Backspace => match self.search_target {
                    SearchTarget::Types => {
                        self.search_query.pop();
                        self.update_filter();
                    }
                    SearchTarget::Members => {
                        self.member_search_query.pop();
                    }
                },
                _ => {}
            }
        }
        false
    }

    fn handle_mouse_event(&mut self, mouse: MouseEvent) {
        match mouse.kind {
            MouseEventKind::ScrollDown => match self.focus {
                Focus::TypeList => {
                    self.next();
                }
                Focus::MethodList => {
                    if !self.current_methods.is_empty() {
                        self.next();
                    }
                }
                Focus::Details => {
                    self.next();
                }
                Focus::IdlView => {
                    self.next();
                }
            },
            MouseEventKind::ScrollUp => match self.focus {
                Focus::TypeList => {
                    self.previous();
                }
                Focus::MethodList => {
                    if !self.current_methods.is_empty() {
                        self.previous();
                    }
                }
                Focus::Details => {
                    self.previous();
                }
                Focus::IdlView => {
                    self.previous();
                }
            },
            _ => {}
        }
    }
}

pub fn run(tlb_path: PathBuf, chm_path: Option<String>) -> Result<(), Box<dyn Error>> {
    // Setup terminal
    enable_raw_mode()?;
    let mut stdout = io::stdout();
    execute!(stdout, EnterAlternateScreen, EnableMouseCapture)?;
    let backend = CrosstermBackend::new(stdout);
    let mut terminal = Terminal::new(backend)?;

    // Create app
    let app = App::new(tlb_path, chm_path)?;
    let res = run_app(&mut terminal, app);

    // Restore terminal
    disable_raw_mode()?;
    execute!(
        terminal.backend_mut(),
        LeaveAlternateScreen,
        DisableMouseCapture
    )?;
    terminal.show_cursor()?;

    if let Err(err) = res {
        println!("{:?}", err)
    }

    Ok(())
}

fn run_app(
    terminal: &mut Terminal<CrosstermBackend<std::io::Stdout>>,
    mut app: App,
) -> io::Result<()> {
    loop {
        terminal.draw(|f| ui(f, &mut app))?;

        match event::read()? {
            Event::Key(key) => {
                if app.handle_key_event(key) {
                    return Ok(());
                }
            }
            Event::Mouse(mouse) => {
                app.handle_mouse_event(mouse);
            }
            _ => {}
        }
    }
}

fn ui(f: &mut ratatui::Frame, app: &mut App) {
    let main_chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints(
            [
                Constraint::Length(1),
                Constraint::Length(3),
                Constraint::Min(0),
                Constraint::Length(2),
            ]
            .as_ref(),
        )
        .split(f.area());

    render_title_bar(f, app, main_chunks[0]);
    render_search_bar(f, app, main_chunks[1]);
    render_footer(f, main_chunks[3]);

    let content_chunks = Layout::default()
        .direction(Direction::Horizontal)
        .constraints([Constraint::Percentage(30), Constraint::Percentage(70)].as_ref())
        .split(main_chunks[2]);

    render_type_list(f, app, content_chunks[0]);

    let right_area = content_chunks[1];
    match app.view_mode {
        ViewMode::Idl => render_idl_view(f, app, right_area),
        ViewMode::Structured => render_structured_view(f, app, right_area),
    }

    if app.show_global_search {
        render_global_search_popup(f, app);
    }

    if app.show_exit_confirmation {
        render_exit_confirmation(f);
    }

    if app.show_help {
        render_help_popup(f);
    }
}

fn centered_rect(percent_x: u16, percent_y: u16, r: Rect) -> Rect {
    let popup_layout = Layout::default()
        .direction(Direction::Vertical)
        .constraints(
            [
                Constraint::Percentage((100 - percent_y) / 2),
                Constraint::Percentage(percent_y),
                Constraint::Percentage((100 - percent_y) / 2),
            ]
            .as_ref(),
        )
        .split(r);

    Layout::default()
        .direction(Direction::Horizontal)
        .constraints(
            [
                Constraint::Percentage((100 - percent_x) / 2),
                Constraint::Percentage(percent_x),
                Constraint::Percentage((100 - percent_x) / 2),
            ]
            .as_ref(),
        )
        .split(popup_layout[1])[1]
}

fn render_title_bar(f: &mut ratatui::Frame, app: &App, area: Rect) {
    let focus_label = match app.focus {
        Focus::TypeList => "Type List",
        Focus::MethodList => "Methods",
        Focus::Details => "Details",
        Focus::IdlView => "IDL View",
    };
    let view_label = match app.view_mode {
        ViewMode::Idl => "IDL",
        ViewMode::Structured => "Structured",
    };
    let title = Line::from(vec![
        Span::styled(
            " TLB→WinMD ",
            Style::default()
                .fg(ACCENT_CYAN)
                .bg(TITLE_BAR_BG)
                .add_modifier(Modifier::BOLD),
        ),
        Span::styled("│", Style::default().fg(FG_DIM).bg(TITLE_BAR_BG)),
        Span::styled(
            format!(" {} types ", app.types.len()),
            Style::default().fg(FG_SECONDARY).bg(TITLE_BAR_BG),
        ),
        Span::styled("│", Style::default().fg(FG_DIM).bg(TITLE_BAR_BG)),
        Span::styled(
            format!(" {}", view_label),
            Style::default().fg(ACCENT_YELLOW).bg(TITLE_BAR_BG),
        ),
        Span::styled(" | ", Style::default().fg(FG_DIM).bg(TITLE_BAR_BG)),
        Span::styled(
            focus_label,
            Style::default().fg(FG_PRIMARY).bg(TITLE_BAR_BG),
        ),
    ]);
    let bar = Paragraph::new(title).style(Style::default().bg(TITLE_BAR_BG));
    f.render_widget(bar, area);
}

fn render_search_bar(f: &mut ratatui::Frame, app: &App, area: Rect) {
    let search_title = match app.search_target {
        SearchTarget::Types => " Search Types (Ctrl+F to switch, Ctrl+P for Global Search) ",
        SearchTarget::Members => " Search Members (Ctrl+F to switch, Ctrl+P for Global Search) ",
    };
    let search_text = match app.search_target {
        SearchTarget::Types => &app.search_query,
        SearchTarget::Members => &app.member_search_query,
    };
    let search_color = match app.search_target {
        SearchTarget::Types => FG_PRIMARY,
        SearchTarget::Members => ACCENT_CYAN,
    };

    let search_paragraph = Paragraph::new(search_text.as_str())
        .block(
            Block::default()
                .borders(Borders::ALL)
                .border_type(BorderType::Rounded)
                .border_style(Style::default().fg(ACCENT_YELLOW))
                .title(search_title)
                .title_style(Style::default().fg(FG_SECONDARY)),
        )
        .style(Style::default().fg(search_color));
    f.render_widget(search_paragraph, area);
}

fn render_footer(f: &mut ratatui::Frame, area: Rect) {
    let legend_line = Line::from(vec![
        Span::styled(" Legend ", Style::default().fg(FG_PRIMARY).bg(FG_DIM)),
        Span::styled(" ↓ ", Style::default().fg(ACCENT_GREEN)),
        Span::styled("In", Style::default().fg(FG_SECONDARY)),
        Span::styled(" ↑ ", Style::default().fg(ACCENT_RED)),
        Span::styled("Out", Style::default().fg(FG_SECONDARY)),
        Span::styled(" ? ", Style::default().fg(ACCENT_YELLOW)),
        Span::styled("Optional", Style::default().fg(FG_SECONDARY)),
        Span::styled(" = ", Style::default().fg(ACCENT_BLUE)),
        Span::styled("Default", Style::default().fg(FG_SECONDARY)),
    ]);
    let keys_line = Line::from(vec![
        Span::styled(" Keys ", Style::default().fg(FG_PRIMARY).bg(FG_DIM)),
        Span::styled(" Tab/V ", Style::default().fg(ACCENT_CYAN)),
        Span::styled("View", Style::default().fg(FG_SECONDARY)),
        Span::styled(" Ctrl+F ", Style::default().fg(ACCENT_CYAN)),
        Span::styled("Search", Style::default().fg(FG_SECONDARY)),
        Span::styled(" Ctrl+P ", Style::default().fg(ACCENT_CYAN)),
        Span::styled("Global", Style::default().fg(FG_SECONDARY)),
        Span::styled(" Esc/q ", Style::default().fg(ACCENT_CYAN)),
        Span::styled("Exit", Style::default().fg(FG_SECONDARY)),
        Span::styled(" ? ", Style::default().fg(ACCENT_CYAN)),
        Span::styled("Help", Style::default().fg(FG_SECONDARY)),
        Span::styled(" PgUp/PgDn ", Style::default().fg(ACCENT_CYAN)),
        Span::styled("Page", Style::default().fg(FG_SECONDARY)),
    ]);
    let footer = Paragraph::new(Text::from(vec![legend_line, keys_line]))
        .style(Style::default().bg(FOOTER_BG).fg(FG_DIM));
    f.render_widget(footer, area);
}

fn render_type_list(f: &mut ratatui::Frame, app: &mut App, area: Rect) {
    let is_focused = app.focus == Focus::TypeList;
    let border_style = if is_focused {
        Style::default().fg(BORDER_FOCUSED)
    } else {
        Style::default().fg(BORDER_UNFOCUSED)
    };
    let fg_text = if is_focused { FG_PRIMARY } else { FG_DIM };

    let items: Vec<ListItem> = app
        .filtered_types
        .iter()
        .map(|(_, name, kind)| {
            let content = Line::from(vec![
                Span::styled(
                    format!("{:<10}", kind),
                    Style::default().fg(if is_focused { ACCENT_YELLOW } else { FG_DIM }),
                ),
                Span::styled(name, Style::default().fg(fg_text)),
            ]);
            ListItem::new(content)
        })
        .collect();

    let list = List::new(items)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .border_type(BorderType::Rounded)
                .border_style(border_style)
                .title(" Types ")
                .title_style(Style::default().fg(if is_focused { FG_PRIMARY } else { FG_DIM })),
        )
        .highlight_style(Style::default().bg(SELECTION_BG).fg(FG_PRIMARY));

    f.render_stateful_widget(list, area, &mut app.list_state);

    f.render_stateful_widget(
        Scrollbar::default()
            .orientation(ScrollbarOrientation::VerticalRight)
            .begin_symbol(Some("↑"))
            .end_symbol(Some("↓"))
            .style(Style::default().fg(if is_focused { BORDER_FOCUSED } else { BORDER_UNFOCUSED })),
        area,
        &mut app.list_scroll_state,
    );
}

fn highlight_idl(text: &str) -> Text<'static> {
    let keywords = [
        "library", "interface", "coclass", "enum", "dispinterface", "module", "typedef",
        "importlib", "import", "HRESULT", "void", "struct", "union",
    ];
    let known_types = [
        "BSTR", "VARIANT_BOOL", "VARIANT", "SAFEARRAY", "GUID", "IUnknown", "IDispatch",
        "oleautomation", "dual", "nonextensible",
    ];
    let lines: Vec<Line> = text
        .lines()
        .map(|line| {
            let trimmed = line.trim_start();
            let indent = &line[..line.len() - trimmed.len()];
            let mut spans = Vec::new();
            if !indent.is_empty() {
                spans.push(Span::styled(indent.to_string(), Style::default().fg(FG_DIM)));
            }

            if trimmed.starts_with("//") {
                spans.push(Span::styled(
                    trimmed.to_string(),
                    Style::default().fg(FG_DIM).add_modifier(Modifier::ITALIC),
                ));
                return Line::from(spans);
            }

            let mut rest = trimmed;
            while !rest.is_empty() {
                // GUIDs
                if let Some(start) = rest.find('{') {
                    if let Some(end) = rest[start..].find('}') {
                        let end = start + end;
                        if start > 0 {
                            spans.push(Span::styled(
                                rest[..start].to_string(),
                                Style::default().fg(FG_PRIMARY),
                            ));
                        }
                        spans.push(Span::styled(
                            rest[start..=end].to_string(),
                            Style::default().fg(ACCENT_GREEN),
                        ));
                        rest = &rest[end + 1..];
                        continue;
                    }
                }
                // Strings
                if let Some(q) = rest.find('"') {
                    if q > 0 {
                        spans.push(Span::styled(
                            rest[..q].to_string(),
                            Style::default().fg(FG_PRIMARY),
                        ));
                    }
                    let end = rest[q + 1..].find('"').map(|e| q + 1 + e + 1).unwrap_or(rest.len());
                    spans.push(Span::styled(
                        rest[q..end].to_string(),
                        Style::default().fg(ACCENT_GREEN),
                    ));
                    rest = &rest[end..];
                    continue;
                }
                // Comments
                if let Some(c) = rest.find("//") {
                    if c > 0 {
                        spans.push(Span::styled(
                            rest[..c].to_string(),
                            Style::default().fg(FG_PRIMARY),
                        ));
                    }
                    spans.push(Span::styled(
                        rest[c..].to_string(),
                        Style::default().fg(FG_DIM).add_modifier(Modifier::ITALIC),
                    ));
                    rest = "";
                    continue;
                }
                // Numbers
                let rest_clone = rest;
                let word_end = rest_clone
                    .find(|c: char| !c.is_alphanumeric() && c != '_' && c != '.')
                    .unwrap_or(rest_clone.len());
                if word_end > 0 {
                    let word = &rest[..word_end];
                    if word.chars().all(|c| c.is_ascii_digit() || c == '.') {
                        spans.push(Span::styled(
                            word.to_string(),
                            Style::default().fg(ACCENT_YELLOW),
                        ));
                    } else if keywords.contains(&word) || (word.to_lowercase() == word && keywords.contains(&&word.to_lowercase().as_str())) {
                        spans.push(Span::styled(
                            word.to_string(),
                            Style::default()
                                .fg(ACCENT_MAGENTA)
                                .add_modifier(Modifier::BOLD),
                        ));
                    } else if known_types.contains(&word) {
                        spans.push(Span::styled(
                            word.to_string(),
                            Style::default().fg(ACCENT_CYAN),
                        ));
                    } else {
                        spans.push(Span::styled(
                            word.to_string(),
                            Style::default().fg(FG_PRIMARY),
                        ));
                    }
                    rest = &rest[word_end..];
                } else {
                    spans.push(Span::styled(
                        rest[..1].to_string(),
                        Style::default().fg(FG_PRIMARY),
                    ));
                    rest = &rest[1..];
                }
            }
            Line::from(spans)
        })
        .collect();
    Text::from(lines)
}

fn render_idl_view(f: &mut ratatui::Frame, app: &mut App, area: Rect) {
    let is_focused = app.focus == Focus::IdlView;
    let border_style = if is_focused {
        Style::default().fg(BORDER_FOCUSED)
    } else {
        Style::default().fg(BORDER_UNFOCUSED)
    };

    let idl_text = highlight_idl(&app.current_idl);
    let idl_paragraph = Paragraph::new(idl_text)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .border_type(BorderType::Rounded)
                .border_style(border_style)
                .title(" IDL Preview ")
                .title_style(Style::default().fg(if is_focused { FG_PRIMARY } else { FG_DIM })),
        )
        .wrap(Wrap { trim: false })
        .scroll((app.idl_scroll_offset, 0));
    f.render_widget(idl_paragraph, area);

    let line_count = app.current_idl.lines().count();
    app.idl_scroll_state = app.idl_scroll_state.content_length(line_count);
    app.idl_scroll_state = app
        .idl_scroll_state
        .position(app.idl_scroll_offset as usize);

    f.render_stateful_widget(
        Scrollbar::default()
            .orientation(ScrollbarOrientation::VerticalRight)
            .begin_symbol(Some("↑"))
            .end_symbol(Some("↓"))
            .style(Style::default().fg(if is_focused { BORDER_FOCUSED } else { BORDER_UNFOCUSED })),
        area,
        &mut app.idl_scroll_state,
    );
}

fn render_structured_view(f: &mut ratatui::Frame, app: &mut App, area: Rect) {
    if !app.current_methods.is_empty() {
        let chunks = Layout::default()
            .direction(Direction::Horizontal)
            .constraints([Constraint::Percentage(40), Constraint::Percentage(60)].as_ref())
            .split(area);
        render_method_list(f, app, chunks[0]);
        render_details(f, app, chunks[1]);
    } else if !app.current_enums.is_empty() {
        render_enum_table(f, app, area);
    } else {
        let idl_text = highlight_idl(&app.current_idl);
        let p = Paragraph::new(idl_text)
            .block(
                Block::default()
                    .borders(Borders::ALL)
                    .border_type(BorderType::Rounded)
                    .border_style(Style::default().fg(BORDER_UNFOCUSED))
                    .title(" IDL Preview (No structured data) ")
                    .title_style(Style::default().fg(FG_DIM)),
            )
            .wrap(Wrap { trim: false });
        f.render_widget(p, area);
    }
}

fn render_method_list(f: &mut ratatui::Frame, app: &mut App, area: Rect) {
    let is_focused = app.focus == Focus::MethodList;
    let border_style = if is_focused {
        Style::default().fg(BORDER_FOCUSED)
    } else {
        Style::default().fg(BORDER_UNFOCUSED)
    };
    let fg_text = if is_focused { FG_PRIMARY } else { FG_DIM };

    let member_query = app.member_search_query.to_lowercase();
    let filtered_methods: Vec<&MethodInfo> = app
        .current_methods
        .iter()
        .filter(|m| m.name.to_lowercase().contains(&member_query))
        .collect();

    let method_items: Vec<ListItem> = filtered_methods
        .iter()
        .map(|m| {
            ListItem::new(Line::from(vec![
                Span::styled(
                    "ƒ ",
                    Style::default().fg(if is_focused { ACCENT_MAGENTA } else { FG_DIM }),
                ),
                Span::styled(&m.name, Style::default().fg(fg_text)),
            ]))
        })
        .collect();

    let method_list = List::new(method_items)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .border_type(BorderType::Rounded)
                .border_style(border_style)
                .title(" Functions ")
                .title_style(Style::default().fg(if is_focused { FG_PRIMARY } else { FG_DIM })),
        )
        .highlight_style(Style::default().bg(SELECTION_BG).fg(FG_PRIMARY));

    f.render_stateful_widget(method_list, area, &mut app.method_list_state);

    app.method_list_scroll_state = app
        .method_list_scroll_state
        .content_length(filtered_methods.len());
    if let Some(i) = app.method_list_state.selected() {
        app.method_list_scroll_state = app.method_list_scroll_state.position(i);
    }
    f.render_stateful_widget(
        Scrollbar::default()
            .orientation(ScrollbarOrientation::VerticalRight)
            .begin_symbol(Some("↑"))
            .end_symbol(Some("↓"))
            .style(Style::default().fg(if is_focused { BORDER_FOCUSED } else { BORDER_UNFOCUSED })),
        area,
        &mut app.method_list_scroll_state,
    );
}

fn render_details(f: &mut ratatui::Frame, app: &mut App, area: Rect) {
    let is_focused = app.focus == Focus::Details;
    let border_style = if is_focused {
        Style::default().fg(BORDER_FOCUSED)
    } else {
        Style::default().fg(BORDER_UNFOCUSED)
    };

    let details_block = Block::default()
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(border_style)
        .title(" Details ")
        .title_style(Style::default().fg(if is_focused { FG_PRIMARY } else { FG_DIM }));

    let inner_details_area = details_block.inner(area);
    f.render_widget(details_block, area);

    let member_query = app.member_search_query.to_lowercase();
    let filtered_methods: Vec<&MethodInfo> = app
        .current_methods
        .iter()
        .filter(|m| m.name.to_lowercase().contains(&member_query))
        .collect();

    if let Some(selected_idx) = app.method_list_state.selected() {
        if let Some(method) = filtered_methods.get(selected_idx) {
            let details_layout = Layout::default()
                .direction(Direction::Vertical)
                .constraints([Constraint::Min(0)].as_ref())
                .split(inner_details_area);

            let mut lines: Vec<Line> = Vec::new();

            // Signature line
            lines.push(Line::from(vec![
                Span::styled("ƒ ", Style::default().fg(ACCENT_MAGENTA)),
                Span::styled(
                    &method.name,
                    Style::default()
                        .fg(FG_PRIMARY)
                        .add_modifier(Modifier::BOLD),
                ),
                Span::styled(" (", Style::default().fg(FG_SECONDARY)),
            ]));

            // Parameters
            for param in &method.params {
                let mut param_spans = Vec::new();
                param_spans.push(Span::styled("    ", Style::default().fg(FG_DIM)));
                if param.flags.contains(&"in".to_string()) {
                    param_spans.push(Span::styled("↓ ", Style::default().fg(ACCENT_GREEN)));
                }
                if param.flags.contains(&"out".to_string()) {
                    param_spans.push(Span::styled("↑ ", Style::default().fg(ACCENT_RED)));
                }
                if let Some(default_val) = &param.default_value {
                    param_spans.push(Span::styled(
                        format!("= {} ", default_val),
                        Style::default().fg(ACCENT_BLUE),
                    ));
                } else if param.flags.contains(&"defaultvalue".to_string()) {
                    param_spans.push(Span::styled("* ", Style::default().fg(ACCENT_BLUE)));
                }
                if param.flags.contains(&"optional".to_string()) {
                    param_spans.push(Span::styled(
                        "? ",
                        Style::default().fg(ACCENT_YELLOW),
                    ));
                }
                param_spans.push(Span::styled(
                    format!("{} ", param.type_name),
                    Style::default().fg(if is_focused { FG_PRIMARY } else { FG_DIM }),
                ));
                param_spans.push(Span::styled(
                    &param.name,
                    Style::default().fg(FG_SECONDARY),
                ));
                param_spans.push(Span::styled(",", Style::default().fg(FG_DIM)));
                lines.push(Line::from(param_spans));
            }

            // Return type
            lines.push(Line::from(vec![
                Span::styled("  ) → ", Style::default().fg(FG_SECONDARY)),
                Span::styled(
                    &method.ret_type,
                    Style::default().fg(ACCENT_GREEN).add_modifier(Modifier::BOLD),
                ),
            ]));
            lines.push(Line::from(""));

            // Documentation
            if let Some(provider) = &app.doc_provider {
                if let Some(doc) = provider.get_doc(&method.name) {
                    lines.push(Line::from(Span::styled(
                        "Description:",
                        Style::default()
                            .fg(ACCENT_CYAN)
                            .add_modifier(Modifier::BOLD | Modifier::UNDERLINED),
                    )));
                    lines.push(Line::from(Span::styled(
                        doc.description.clone(),
                        Style::default().fg(FG_SECONDARY),
                    )));
                    lines.push(Line::from(""));

                    if !doc.parameters.is_empty() {
                        lines.push(Line::from(Span::styled(
                            "Parameters:",
                            Style::default()
                                .fg(ACCENT_CYAN)
                                .add_modifier(Modifier::BOLD | Modifier::UNDERLINED),
                        )));
                        for param in &doc.parameters {
                            lines.push(Line::from(vec![
                                Span::styled(
                                    format!("- {}: ", param.name),
                                    Style::default().fg(ACCENT_YELLOW),
                                ),
                                Span::styled(
                                    param.description.clone(),
                                    Style::default().fg(FG_SECONDARY),
                                ),
                            ]));
                        }
                    }
                } else {
                    lines.push(Line::from(Span::styled(
                        "No documentation found.",
                        Style::default().fg(FG_DIM),
                    )));
                }
            } else {
                lines.push(Line::from(Span::styled(
                    "Documentation provider not available.",
                    Style::default().fg(FG_DIM),
                )));
            }

            let total_lines = lines.len();
            let paragraph = Paragraph::new(lines)
                .wrap(Wrap { trim: false })
                .scroll((app.details_scroll_offset, 0));

            f.render_widget(paragraph, details_layout[0]);

            app.details_scroll_state = app.details_scroll_state.content_length(total_lines);
            app.details_scroll_state = app
                .details_scroll_state
                .position(app.details_scroll_offset as usize);

            f.render_stateful_widget(
                Scrollbar::default()
                    .orientation(ScrollbarOrientation::VerticalRight)
                    .begin_symbol(Some("↑"))
                    .end_symbol(Some("↓"))
                    .style(Style::default().fg(if is_focused { BORDER_FOCUSED } else { BORDER_UNFOCUSED })),
                details_layout[0],
                &mut app.details_scroll_state,
            );
        }
    }
}

fn render_enum_table(f: &mut ratatui::Frame, app: &mut App, area: Rect) {
    let is_focused = app.focus == Focus::MethodList;
    let border_style = if is_focused {
        Style::default().fg(BORDER_FOCUSED)
    } else {
        Style::default().fg(BORDER_UNFOCUSED)
    };

    let member_query = app.member_search_query.to_lowercase();
    let header_cells = ["Name", "Value"]
        .iter()
        .map(|h| Cell::from(*h).style(Style::default().fg(FG_PRIMARY)));
    let header = Row::new(header_cells)
        .style(Style::default().bg(SELECTION_BG))
        .height(1);

    let filtered_enums: Vec<&EnumItemInfo> = app
        .current_enums
        .iter()
        .filter(|e| e.name.to_lowercase().contains(&member_query))
        .collect();

    let rows = filtered_enums.iter().map(|item| {
        Row::new(vec![
            Cell::from(Span::styled(
                &item.name,
                Style::default().fg(if is_focused { ACCENT_CYAN } else { FG_DIM }),
            )),
            Cell::from(Span::styled(
                &item.value,
                Style::default().fg(if is_focused { FG_PRIMARY } else { FG_DIM }),
            )),
        ])
    });

    let table = Table::new(
        rows,
        [Constraint::Percentage(70), Constraint::Percentage(30)],
    )
    .header(header)
    .block(
        Block::default()
            .borders(Borders::ALL)
            .border_type(BorderType::Rounded)
            .border_style(border_style)
            .title(" Enum Values ")
            .title_style(Style::default().fg(if is_focused { FG_PRIMARY } else { FG_DIM })),
    )
    .row_highlight_style(Style::default().bg(SELECTION_BG).fg(FG_PRIMARY));

    f.render_stateful_widget(table, area, &mut app.content_table_state);

    app.content_scroll_state = app
        .content_scroll_state
        .content_length(filtered_enums.len());
    if let Some(i) = app.content_table_state.selected() {
        app.content_scroll_state = app.content_scroll_state.position(i);
    }
    f.render_stateful_widget(
        Scrollbar::default()
            .orientation(ScrollbarOrientation::VerticalRight)
            .begin_symbol(Some("↑"))
            .end_symbol(Some("↓"))
            .style(Style::default().fg(if is_focused { BORDER_FOCUSED } else { BORDER_UNFOCUSED })),
        area,
        &mut app.content_scroll_state,
    );
}

fn render_global_search_popup(f: &mut ratatui::Frame, app: &mut App) {
    let area = centered_rect(60, 50, f.area());
    f.render_widget(Clear, area);

    let block = Block::default()
        .title(" Global Search (Esc to close) ")
        .title_style(Style::default().fg(ACCENT_CYAN))
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(BORDER_FOCUSED))
        .style(Style::default().bg(BG_PRIMARY));

    let inner_area = block.inner(area);
    f.render_widget(block, area);

    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints([Constraint::Length(3), Constraint::Min(0)].as_ref())
        .split(inner_area);

    let search_paragraph = Paragraph::new(app.global_search_query.as_str())
        .block(
            Block::default()
                .borders(Borders::ALL)
                .border_type(BorderType::Rounded)
                .border_style(Style::default().fg(BORDER_UNFOCUSED))
                .title(" Query ")
                .title_style(Style::default().fg(FG_DIM)),
        )
        .style(Style::default().fg(ACCENT_CYAN));
    f.render_widget(search_paragraph, chunks[0]);

    let items: Vec<ListItem> = app
        .global_search_results
        .iter()
        .map(|&idx| {
            if let Some(item) = app.all_search_items.get(idx) {
                ListItem::new(Line::from(vec![
                    Span::styled(
                        format!("{:<10}", item.kind),
                        Style::default().fg(ACCENT_YELLOW),
                    ),
                    Span::styled(
                        format!("{}::", item.type_name),
                        Style::default().fg(FG_SECONDARY),
                    ),
                    Span::styled(&item.member_name, Style::default().fg(ACCENT_CYAN)),
                ]))
            } else {
                ListItem::new("Invalid Item")
            }
        })
        .collect();

    let list = List::new(items)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .border_type(BorderType::Rounded)
                .border_style(Style::default().fg(BORDER_UNFOCUSED))
                .title(" Results ")
                .title_style(Style::default().fg(FG_DIM)),
        )
        .highlight_style(Style::default().bg(SELECTION_BG).fg(FG_PRIMARY));

    f.render_stateful_widget(list, chunks[1], &mut app.global_search_state);

    app.global_search_scroll_state = app
        .global_search_scroll_state
        .content_length(app.global_search_results.len());
    if let Some(i) = app.global_search_state.selected() {
        app.global_search_scroll_state = app.global_search_scroll_state.position(i);
    }
    f.render_stateful_widget(
        Scrollbar::default()
            .orientation(ScrollbarOrientation::VerticalRight)
            .begin_symbol(Some("↑"))
            .end_symbol(Some("↓"))
            .style(Style::default().fg(BORDER_FOCUSED)),
        chunks[1],
        &mut app.global_search_scroll_state,
    );
}

fn render_exit_confirmation(f: &mut ratatui::Frame) {
    let area = centered_rect(60, 20, f.area());
    f.render_widget(Clear, area);

    let block = Block::default()
        .title(" Exit Confirmation ")
        .title_style(Style::default().fg(ACCENT_RED).add_modifier(Modifier::BOLD))
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(ACCENT_RED))
        .style(Style::default().bg(BG_PRIMARY));

    let inner_area = block.inner(area);
    f.render_widget(block, area);

    let text = vec![
        Line::from(Span::styled(
            "Are you sure you want to exit?",
            Style::default().fg(FG_PRIMARY),
        )),
        Line::from(""),
        Line::from(vec![
            Span::styled(" Y ", Style::default().bg(ACCENT_RED).fg(BG_PRIMARY)),
            Span::styled(" / ", Style::default().fg(FG_DIM)),
            Span::styled(
                " Enter ",
                Style::default()
                    .fg(ACCENT_RED)
                    .add_modifier(Modifier::BOLD),
            ),
            Span::styled(": Confirm", Style::default().fg(FG_SECONDARY)),
        ]),
        Line::from(vec![
            Span::styled(" N ", Style::default().bg(FG_DIM).fg(BG_PRIMARY)),
            Span::styled(" / ", Style::default().fg(FG_DIM)),
            Span::styled(
                " Esc ",
                Style::default()
                    .fg(ACCENT_RED)
                    .add_modifier(Modifier::BOLD),
            ),
            Span::styled(": Cancel", Style::default().fg(FG_SECONDARY)),
        ]),
    ];

    let paragraph = Paragraph::new(text)
        .alignment(ratatui::layout::Alignment::Center)
        .wrap(Wrap { trim: true });

    // Vertically center the text
    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints(
            [
                Constraint::Percentage(25),
                Constraint::Percentage(50),
                Constraint::Percentage(25),
            ]
            .as_ref(),
        )
        .split(inner_area);

    f.render_widget(paragraph, chunks[1]);
}

fn render_help_popup(f: &mut ratatui::Frame) {
    let area = centered_rect(55, 60, f.area());
    f.render_widget(Clear, area);

    let block = Block::default()
        .title(" Help - Keyboard Shortcuts ")
        .title_style(Style::default().fg(ACCENT_CYAN).add_modifier(Modifier::BOLD))
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(BORDER_FOCUSED))
        .style(Style::default().bg(BG_PRIMARY));

    let inner_area = block.inner(area);
    f.render_widget(block, area);

    let shortcuts = vec![
        Line::from(vec![
            Span::styled(" ↑/↓ ", Style::default().fg(ACCENT_CYAN)),
            Span::styled("Navigate items", Style::default().fg(FG_PRIMARY)),
        ]),
        Line::from(vec![
            Span::styled(" PgUp/PgDn ", Style::default().fg(ACCENT_CYAN)),
            Span::styled("Jump 10 items", Style::default().fg(FG_PRIMARY)),
        ]),
        Line::from(vec![
            Span::styled(" ←/→ ", Style::default().fg(ACCENT_CYAN)),
            Span::styled("Change focus panel", Style::default().fg(FG_PRIMARY)),
        ]),
        Line::from(vec![
            Span::styled(" Tab/V ", Style::default().fg(ACCENT_CYAN)),
            Span::styled("Toggle IDL/Structured view", Style::default().fg(FG_PRIMARY)),
        ]),
        Line::from(vec![
            Span::styled(" / ", Style::default().fg(ACCENT_CYAN)),
            Span::styled(
                "Type search in current target",
                Style::default().fg(FG_PRIMARY),
            ),
        ]),
        Line::from(vec![
            Span::styled(" Ctrl+F ", Style::default().fg(ACCENT_CYAN)),
            Span::styled(
                "Toggle search target (Types/Members)",
                Style::default().fg(FG_PRIMARY),
            ),
        ]),
        Line::from(vec![
            Span::styled(" Ctrl+P ", Style::default().fg(ACCENT_CYAN)),
            Span::styled("Open global search", Style::default().fg(FG_PRIMARY)),
        ]),
        Line::from(vec![
            Span::styled(" Ctrl+C ", Style::default().fg(ACCENT_CYAN)),
            Span::styled("Exit immediately", Style::default().fg(FG_PRIMARY)),
        ]),
        Line::from(vec![
            Span::styled(" Esc ", Style::default().fg(ACCENT_CYAN)),
            Span::styled(
                "Exit confirmation / Close popup",
                Style::default().fg(FG_PRIMARY),
            ),
        ]),
        Line::from(vec![
            Span::styled(" q ", Style::default().fg(ACCENT_CYAN)),
            Span::styled(
                "Exit confirmation (when search is empty)",
                Style::default().fg(FG_PRIMARY),
            ),
        ]),
        Line::from(vec![
            Span::styled(" ? / F1 ", Style::default().fg(ACCENT_CYAN)),
            Span::styled("Toggle this help screen", Style::default().fg(FG_PRIMARY)),
        ]),
    ];

    let text = Text::from(shortcuts);

    let paragraph = Paragraph::new(text)
        .alignment(ratatui::layout::Alignment::Left);

    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints(
            [
                Constraint::Percentage(15),
                Constraint::Percentage(70),
                Constraint::Percentage(15),
            ]
            .as_ref(),
        )
        .split(inner_area);

    f.render_widget(paragraph, chunks[1]);
}
