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
        Block, Borders, Cell, Clear, List, ListItem, ListState, Paragraph, Row, Scrollbar,
        ScrollbarOrientation, ScrollbarState, Table, TableState, Wrap,
    },
};
use std::{error::Error, io, path::PathBuf};

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
                Constraint::Length(3),
                Constraint::Min(0),
                Constraint::Length(1),
            ]
            .as_ref(),
        )
        .split(f.area());

    render_search_bar(f, app, main_chunks[0]);
    render_footer(f, main_chunks[2]);

    let content_chunks = Layout::default()
        .direction(Direction::Horizontal)
        .constraints([Constraint::Percentage(30), Constraint::Percentage(70)].as_ref())
        .split(main_chunks[1]);

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

fn render_search_bar(f: &mut ratatui::Frame, app: &App, area: Rect) {
    let search_title = match app.search_target {
        SearchTarget::Types => "Search Types (Ctrl+F to switch, Ctrl+P for Global Search)",
        SearchTarget::Members => "Search Members (Ctrl+F to switch, Ctrl+P for Global Search)",
    };
    let search_text = match app.search_target {
        SearchTarget::Types => &app.search_query,
        SearchTarget::Members => &app.member_search_query,
    };
    let search_color = match app.search_target {
        SearchTarget::Types => Color::White,
        SearchTarget::Members => Color::Cyan,
    };

    let search_paragraph = Paragraph::new(search_text.as_str())
        .block(Block::default().borders(Borders::ALL).title(search_title))
        .style(Style::default().fg(search_color));
    f.render_widget(search_paragraph, area);
}

fn render_footer(f: &mut ratatui::Frame, area: Rect) {
    let footer_text = Line::from(vec![
        Span::styled(
            " Legend: ",
            Style::default().bg(Color::White).fg(Color::Black),
        ),
        Span::styled(" ↓ ", Style::default().fg(Color::Green)),
        Span::raw("In "),
        Span::styled("↑ ", Style::default().fg(Color::Red)),
        Span::raw("Out "),
        Span::styled("? ", Style::default().fg(Color::Yellow)),
        Span::raw("Optional "),
        Span::styled("= ", Style::default().fg(Color::Blue)),
        Span::raw("Default "),
        Span::styled(" | ", Style::default().fg(Color::DarkGray)),
        Span::styled(
            " Keys: ",
            Style::default().bg(Color::White).fg(Color::Black),
        ),
        Span::styled(" Tab/V ", Style::default().fg(Color::Cyan)),
        Span::raw("Toggle View "),
        Span::styled(" Ctrl+F ", Style::default().fg(Color::Cyan)),
        Span::raw("Switch Search "),
        Span::styled(" Ctrl+P ", Style::default().fg(Color::Cyan)),
        Span::raw("Global Search "),
        Span::styled(" Esc ", Style::default().fg(Color::Cyan)),
        Span::raw("Exit "),
        Span::styled(" ? ", Style::default().fg(Color::Cyan)),
        Span::raw("Help "),
    ]);
    let footer = Paragraph::new(footer_text).style(Style::default().bg(Color::DarkGray));
    f.render_widget(footer, area);
}

fn render_type_list(f: &mut ratatui::Frame, app: &mut App, area: Rect) {
    let items: Vec<ListItem> = app
        .filtered_types
        .iter()
        .map(|(_, name, kind)| {
            let content = Line::from(vec![
                Span::styled(format!("{:<10}", kind), Style::default().fg(Color::Yellow)),
                Span::raw(name),
            ]);
            ListItem::new(content)
        })
        .collect();

    let type_list_border_style = if app.focus == Focus::TypeList {
        Style::default().fg(Color::Yellow)
    } else {
        Style::default()
    };

    let list = List::new(items)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .border_style(type_list_border_style)
                .title("Types"),
        )
        .highlight_style(Style::default().bg(Color::Blue).fg(Color::White));

    f.render_stateful_widget(list, area, &mut app.list_state);

    f.render_stateful_widget(
        Scrollbar::default()
            .orientation(ScrollbarOrientation::VerticalRight)
            .begin_symbol(Some("↑"))
            .end_symbol(Some("↓")),
        area,
        &mut app.list_scroll_state,
    );
}

fn render_idl_view(f: &mut ratatui::Frame, app: &mut App, area: Rect) {
    let idl_border_style = if app.focus == Focus::IdlView {
        Style::default().fg(Color::Yellow)
    } else {
        Style::default()
    };

    let idl_paragraph = Paragraph::new(app.current_idl.as_str())
        .block(
            Block::default()
                .borders(Borders::ALL)
                .border_style(idl_border_style)
                .title("IDL Preview"),
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
            .end_symbol(Some("↓")),
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
        let p = Paragraph::new(app.current_idl.as_str())
            .block(
                Block::default()
                    .borders(Borders::ALL)
                    .title("IDL Preview (No structured data)"),
            )
            .wrap(Wrap { trim: false });
        f.render_widget(p, area);
    }
}

fn render_method_list(f: &mut ratatui::Frame, app: &mut App, area: Rect) {
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
                Span::styled("ƒ ", Style::default().fg(Color::Magenta)),
                Span::raw(&m.name),
            ]))
        })
        .collect();

    let method_border_style = if app.focus == Focus::MethodList {
        Style::default().fg(Color::Yellow)
    } else {
        Style::default()
    };

    let method_list = List::new(method_items)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .border_style(method_border_style)
                .title("Functions"),
        )
        .highlight_style(Style::default().bg(Color::Blue).fg(Color::White));

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
            .end_symbol(Some("↓")),
        area,
        &mut app.method_list_scroll_state,
    );
}

fn render_details(f: &mut ratatui::Frame, app: &mut App, area: Rect) {
    let details_border_style = if app.focus == Focus::Details {
        Style::default().fg(Color::Yellow)
    } else {
        Style::default()
    };

    let details_block = Block::default()
        .borders(Borders::ALL)
        .border_style(details_border_style)
        .title("Details");

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

            let mut lines = Vec::new();

            // Signature
            lines.push(Line::from(vec![
                Span::styled("ƒ ", Style::default().fg(Color::Magenta)),
                Span::styled(
                    &method.name,
                    Style::default()
                        .fg(Color::White)
                        .add_modifier(Modifier::BOLD),
                ),
                Span::raw(" ("),
            ]));

            for param in &method.params {
                let mut param_spans = Vec::new();
                param_spans.push(Span::raw("    "));
                if param.flags.contains(&"in".to_string()) {
                    param_spans.push(Span::styled("↓ ", Style::default().fg(Color::Green)));
                }
                if param.flags.contains(&"out".to_string()) {
                    param_spans.push(Span::styled("↑ ", Style::default().fg(Color::Red)));
                }
                if let Some(default_val) = &param.default_value {
                    param_spans.push(Span::styled(
                        format!("= {} ", default_val),
                        Style::default().fg(Color::Blue),
                    ));
                } else if param.flags.contains(&"defaultvalue".to_string()) {
                    param_spans.push(Span::styled("* ", Style::default().fg(Color::Blue)));
                }
                if param.flags.contains(&"optional".to_string()) {
                    param_spans.push(Span::styled("? ", Style::default().fg(Color::Yellow)));
                }
                param_spans.push(Span::styled(
                    format!("{} ", param.type_name),
                    Style::default().fg(Color::White),
                ));
                param_spans.push(Span::raw(&param.name));
                param_spans.push(Span::raw(","));
                lines.push(Line::from(param_spans));
            }

            lines.push(Line::from(vec![
                Span::raw("  ) -> "),
                Span::styled(&method.ret_type, Style::default().fg(Color::Green)),
            ]));
            lines.push(Line::from(""));

            // Documentation
            if let Some(provider) = &app.doc_provider {
                if let Some(doc) = provider.get_doc(&method.name) {
                    lines.push(Line::from(Span::styled(
                        "Description:",
                        Style::default().add_modifier(Modifier::BOLD | Modifier::UNDERLINED),
                    )));
                    lines.push(Line::from(doc.description.clone()));
                    lines.push(Line::from(""));

                    if !doc.parameters.is_empty() {
                        lines.push(Line::from(Span::styled(
                            "Parameters:",
                            Style::default().add_modifier(Modifier::BOLD | Modifier::UNDERLINED),
                        )));
                        for param in &doc.parameters {
                            lines.push(Line::from(vec![
                                Span::styled(
                                    format!("- {}: ", param.name),
                                    Style::default().fg(Color::Cyan),
                                ),
                                Span::raw(param.description.clone()),
                            ]));
                        }
                    }
                } else {
                    lines.push(Line::from(Span::styled(
                        "No documentation found.",
                        Style::default().fg(Color::DarkGray),
                    )));
                }
            } else {
                lines.push(Line::from(Span::styled(
                    "Documentation provider not available.",
                    Style::default().fg(Color::DarkGray),
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
                    .end_symbol(Some("↓")),
                details_layout[0],
                &mut app.details_scroll_state,
            );
        }
    }
}

fn render_enum_table(f: &mut ratatui::Frame, app: &mut App, area: Rect) {
    let member_query = app.member_search_query.to_lowercase();
    let header_cells = ["Name", "Value"]
        .iter()
        .map(|h| Cell::from(*h).style(Style::default().fg(Color::White)));
    let header = Row::new(header_cells)
        .style(Style::default().bg(Color::Blue))
        .height(1);

    let filtered_enums: Vec<&EnumItemInfo> = app
        .current_enums
        .iter()
        .filter(|e| e.name.to_lowercase().contains(&member_query))
        .collect();

    let rows = filtered_enums.iter().map(|item| {
        Row::new(vec![
            Cell::from(Span::styled(&item.name, Style::default().fg(Color::Cyan))),
            Cell::from(Span::styled(&item.value, Style::default().fg(Color::White))),
        ])
    });

    let content_border_style = if app.focus == Focus::MethodList {
        Style::default().fg(Color::Yellow)
    } else {
        Style::default()
    };

    let table = Table::new(
        rows,
        [Constraint::Percentage(70), Constraint::Percentage(30)],
    )
    .header(header)
    .block(
        Block::default()
            .borders(Borders::ALL)
            .border_style(content_border_style)
            .title("Enum Values"),
    )
    .row_highlight_style(Style::default().bg(Color::Blue));

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
            .end_symbol(Some("↓")),
        area,
        &mut app.content_scroll_state,
    );
}

fn render_global_search_popup(f: &mut ratatui::Frame, app: &mut App) {
    let area = centered_rect(60, 50, f.area());
    f.render_widget(Clear, area);

    let block = Block::default()
        .title("Global Search (Esc to close)")
        .borders(Borders::ALL)
        .style(Style::default().bg(Color::Black));

    let inner_area = block.inner(area);
    f.render_widget(block, area);

    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints([Constraint::Length(3), Constraint::Min(0)].as_ref())
        .split(inner_area);

    let search_paragraph = Paragraph::new(app.global_search_query.as_str())
        .block(Block::default().borders(Borders::ALL).title("Query"))
        .style(Style::default().fg(Color::Cyan));
    f.render_widget(search_paragraph, chunks[0]);

    let items: Vec<ListItem> = app
        .global_search_results
        .iter()
        .map(|&idx| {
            if let Some(item) = app.all_search_items.get(idx) {
                ListItem::new(Line::from(vec![
                    Span::styled(
                        format!("{:<10}", item.kind),
                        Style::default().fg(Color::Yellow),
                    ),
                    Span::raw(format!("{}::", item.type_name)),
                    Span::styled(&item.member_name, Style::default().fg(Color::Cyan)),
                ]))
            } else {
                ListItem::new("Invalid Item")
            }
        })
        .collect();

    let list = List::new(items)
        .block(Block::default().borders(Borders::ALL).title("Results"))
        .highlight_style(Style::default().bg(Color::Blue).fg(Color::White));

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
            .end_symbol(Some("↓")),
        chunks[1],
        &mut app.global_search_scroll_state,
    );
}

fn render_exit_confirmation(f: &mut ratatui::Frame) {
    let area = centered_rect(60, 20, f.area());
    f.render_widget(Clear, area);

    let block = Block::default()
        .title("Exit Confirmation")
        .borders(Borders::ALL)
        .style(Style::default().bg(Color::Red).fg(Color::White));

    let inner_area = block.inner(area);
    f.render_widget(block, area);

    let text = vec![
        Line::from("Are you sure you want to exit?"),
        Line::from(""),
        Line::from(vec![
            Span::styled("Y", Style::default().add_modifier(Modifier::BOLD)),
            Span::raw(" / "),
            Span::styled("Enter", Style::default().add_modifier(Modifier::BOLD)),
            Span::raw(": Confirm"),
        ]),
        Line::from(vec![
            Span::styled("N", Style::default().add_modifier(Modifier::BOLD)),
            Span::raw(" / "),
            Span::styled("Esc", Style::default().add_modifier(Modifier::BOLD)),
            Span::raw(": Cancel"),
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
        .title("Help - Keyboard Shortcuts")
        .borders(Borders::ALL)
        .style(Style::default().bg(Color::Black).fg(Color::White));

    let inner_area = block.inner(area);
    f.render_widget(block, area);

    let shortcuts = vec![
        Line::from(vec![
            Span::styled(" ↑/↓ ", Style::default().fg(Color::Cyan)),
            Span::raw("Navigate items"),
        ]),
        Line::from(vec![
            Span::styled(" PgUp/PgDn ", Style::default().fg(Color::Cyan)),
            Span::raw("Jump 10 items"),
        ]),
        Line::from(vec![
            Span::styled(" ←/→ ", Style::default().fg(Color::Cyan)),
            Span::raw("Change focus panel"),
        ]),
        Line::from(vec![
            Span::styled(" Tab/V ", Style::default().fg(Color::Cyan)),
            Span::raw("Toggle IDL/Structured view"),
        ]),
        Line::from(vec![
            Span::styled(" / ", Style::default().fg(Color::Cyan)),
            Span::raw("Type search in current target"),
        ]),
        Line::from(vec![
            Span::styled(" Ctrl+F ", Style::default().fg(Color::Cyan)),
            Span::raw("Toggle search target (Types/Members)"),
        ]),
        Line::from(vec![
            Span::styled(" Ctrl+P ", Style::default().fg(Color::Cyan)),
            Span::raw("Open global search"),
        ]),
        Line::from(vec![
            Span::styled(" Ctrl+C ", Style::default().fg(Color::Cyan)),
            Span::raw("Exit immediately"),
        ]),
        Line::from(vec![
            Span::styled(" Esc ", Style::default().fg(Color::Cyan)),
            Span::raw("Exit confirmation / Close popup"),
        ]),
        Line::from(vec![
            Span::styled(" q ", Style::default().fg(Color::Cyan)),
            Span::raw("Exit confirmation (when search is empty)"),
        ]),
        Line::from(vec![
            Span::styled(" ? / F1 ", Style::default().fg(Color::Cyan)),
            Span::raw("Toggle this help screen"),
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
