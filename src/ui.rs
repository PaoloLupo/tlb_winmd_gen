use crate::chm_doc::ChmDocumentationProvider;
use crate::idlgen::{EnumItemInfo, MethodInfo, TypeLibInfo};
use crossterm::{
    event::{
        self, DisableMouseCapture, EnableMouseCapture, Event, KeyCode, KeyEventKind, KeyModifiers,
    },
    execute,
    terminal::{
        DisableLineWrap, EnableLineWrap, EnterAlternateScreen, LeaveAlternateScreen,
        disable_raw_mode, enable_raw_mode,
    },
};
use ratatui::{
    Terminal,
    backend::{Backend, CrosstermBackend},
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
    Content,
}

struct SearchItem {
    type_index: usize,
    type_name: String,
    member_name: String,
    kind: String, // "Method" or "Enum"
}

struct App {
    tlb_path: PathBuf,
    type_lib_info: TypeLibInfo,
    doc_provider: Option<ChmDocumentationProvider>, // New field
    show_doc: bool,                                 // New field
    types: Vec<(String, String)>,                   // Name, Kind
    filtered_types: Vec<(usize, String, String)>,   // Original Index, Name, Kind
    list_state: ListState,
    list_scroll_state: ScrollbarState, // Scrollbar for Type List
    current_idl: String,
    current_methods: Vec<MethodInfo>,
    current_enums: Vec<EnumItemInfo>,
    search_query: String,
    member_search_query: String,
    view_mode: ViewMode,
    search_target: SearchTarget,
    focus: Focus,
    content_table_state: TableState,
    content_scroll_state: ScrollbarState,
    // Global Search
    all_search_items: Vec<SearchItem>,
    show_global_search: bool,
    global_search_query: String,
    global_search_results: Vec<usize>, // Indices into all_search_items
    global_search_state: ListState,
    global_search_scroll_state: ScrollbarState, // Scrollbar for Global Search
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
                types.push((name.clone(), kind));

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
                            kind: "Enum".to_string(),
                        });
                    }
                }
            }
        }

        let doc_provider = if let Some(path) = chm_path {
            match ChmDocumentationProvider::new(&path) {
                Ok(provider) => Some(provider),
                Err(_e) => {
                    // Log error or just ignore? For TUI, maybe just print to stderr before starting?
                    // Or just ignore and have no docs.
                    // eprintln!("Failed to load CHM: {}", e);
                    None
                }
            }
        } else {
            None
        };

        let mut app = App {
            tlb_path,
            type_lib_info,
            doc_provider,
            show_doc: false,
            types,
            filtered_types: Vec::new(),
            list_state: ListState::default(),
            list_scroll_state: ScrollbarState::default(),
            current_idl: String::new(),
            current_methods: Vec::new(),
            current_enums: Vec::new(),
            search_query: String::new(),
            member_search_query: String::new(),
            view_mode: ViewMode::Structured,
            search_target: SearchTarget::Types,
            focus: Focus::TypeList,
            content_table_state: TableState::default(),
            content_scroll_state: ScrollbarState::default(),
            all_search_items,
            show_global_search: false,
            global_search_query: String::new(),
            global_search_results: Vec::new(),
            global_search_state: ListState::default(),
            global_search_scroll_state: ScrollbarState::default(),
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
                self.content_table_state.select(None);
                self.content_scroll_state = ScrollbarState::default();
                if !self.current_methods.is_empty() || !self.current_enums.is_empty() {
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
            Focus::Content => {
                let len = if !self.current_methods.is_empty() {
                    self.current_methods.len()
                } else {
                    self.current_enums.len()
                };
                if len > 0 {
                    let i = match self.content_table_state.selected() {
                        Some(i) => {
                            if i >= len - 1 {
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
            Focus::Content => {
                let len = if !self.current_methods.is_empty() {
                    self.current_methods.len()
                } else {
                    self.current_enums.len()
                };
                if len > 0 {
                    let i = match self.content_table_state.selected() {
                        Some(i) => {
                            if i == 0 {
                                len - 1
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
                let (type_index, member_name) =
                    if let Some(item) = self.all_search_items.get(item_idx) {
                        (item.type_index, item.member_name.clone())
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

                // Select the member in the content table
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
                        self.content_table_state.select(Some(pos));
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
            }
        }
    }

    fn get_selected_name(&self) -> Option<String> {
        match self.focus {
            Focus::TypeList => {
                if let Some(idx) = self.list_state.selected() {
                    if idx < self.filtered_types.len() {
                        return Some(self.filtered_types[idx].1.clone());
                    }
                }
            }
            Focus::Content => {
                if let Some(idx) = self.content_table_state.selected() {
                    let member_query = self.member_search_query.to_lowercase();
                    if !self.current_methods.is_empty() {
                        let filtered_methods: Vec<&MethodInfo> = self
                            .current_methods
                            .iter()
                            .filter(|m| m.name.to_lowercase().contains(&member_query))
                            .collect();
                        if idx < filtered_methods.len() {
                            return Some(filtered_methods[idx].name.clone());
                        }
                    } else if !self.current_enums.is_empty() {
                        let filtered_enums: Vec<&EnumItemInfo> = self
                            .current_enums
                            .iter()
                            .filter(|e| e.name.to_lowercase().contains(&member_query))
                            .collect();
                        if idx < filtered_enums.len() {
                            return Some(filtered_enums[idx].name.clone());
                        }
                    }
                }
            }
        }
        None
    }

    fn is_method_selected(&self) -> bool {
        self.focus == Focus::Content && !self.current_methods.is_empty()
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

fn run_app<B: Backend>(terminal: &mut Terminal<B>, mut app: App) -> io::Result<()> {
    loop {
        terminal.draw(|f| ui(f, &mut app))?;

        if let Event::Key(key) = event::read()? {
            if key.kind == KeyEventKind::Press {
                if app.show_global_search {
                    match key.code {
                        KeyCode::Esc => app.show_global_search = false,
                        KeyCode::Down => app.next_global_result(),
                        KeyCode::Up => app.previous_global_result(),
                        KeyCode::Enter => app.select_global_result(),
                        KeyCode::Char(c) => {
                            app.global_search_query.push(c);
                            app.update_global_search();
                        }
                        KeyCode::Backspace => {
                            app.global_search_query.pop();
                            app.update_global_search();
                        }
                        _ => {}
                    }
                } else {
                    match key.code {
                        KeyCode::Char('q')
                            if app.search_query.is_empty()
                                && app.member_search_query.is_empty() =>
                        {
                            return Ok(());
                        }
                        KeyCode::Esc => return Ok(()),
                        KeyCode::Down => app.next(),
                        KeyCode::Up => app.previous(),
                        KeyCode::Right => app.focus = Focus::Content,
                        KeyCode::Left => app.focus = Focus::TypeList,
                        KeyCode::Tab => app.toggle_view(),
                        KeyCode::Char('p') if key.modifiers.contains(KeyModifiers::CONTROL) => {
                            app.show_global_search = true;
                            app.global_search_query.clear();
                            app.update_global_search();
                        }
                        KeyCode::Char('f') if key.modifiers.contains(KeyModifiers::CONTROL) => {
                            app.toggle_search_target();
                        }
                        KeyCode::Enter => {
                            if app.is_method_selected() {
                                app.show_doc = !app.show_doc;
                            }
                        }
                        KeyCode::Char(c) => match app.search_target {
                            SearchTarget::Types => {
                                app.search_query.push(c);
                                app.update_filter();
                            }
                            SearchTarget::Members => {
                                app.member_search_query.push(c);
                            }
                        },
                        KeyCode::Backspace => match app.search_target {
                            SearchTarget::Types => {
                                app.search_query.pop();
                                app.update_filter();
                            }
                            SearchTarget::Members => {
                                app.member_search_query.pop();
                            }
                        },
                        _ => {}
                    }
                }
            }
        }
    }
}

fn ui(f: &mut ratatui::Frame, app: &mut App) {
    let main_chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints([Constraint::Length(3), Constraint::Min(0)].as_ref())
        .split(f.area());

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
    f.render_widget(search_paragraph, main_chunks[0]);

    let content_chunks = Layout::default()
        .direction(Direction::Horizontal)
        .constraints([Constraint::Percentage(30), Constraint::Percentage(70)].as_ref())
        .split(main_chunks[1]);

    // Left panel: List of types
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

    let list = List::new(items)
        .block(Block::default().borders(Borders::ALL).title("Types"))
        .highlight_style(Style::default().bg(Color::Blue).fg(Color::White));

    f.render_stateful_widget(list, content_chunks[0], &mut app.list_state);

    // Render Scrollbar for Types
    f.render_stateful_widget(
        Scrollbar::default()
            .orientation(ScrollbarOrientation::VerticalRight)
            .begin_symbol(Some("↑"))
            .end_symbol(Some("↓")),
        content_chunks[0],
        &mut app.list_scroll_state,
    );

    // Right panel: Content (IDL or Structured)
    let content_border_style = if app.focus == Focus::Content {
        Style::default().fg(Color::Yellow)
    } else {
        Style::default()
    };

    match app.view_mode {
        ViewMode::Idl => {
            let idl_paragraph = Paragraph::new(app.current_idl.as_str())
                .block(
                    Block::default()
                        .borders(Borders::ALL)
                        .border_style(content_border_style)
                        .title("IDL Preview"),
                )
                .wrap(Wrap { trim: false });
            f.render_widget(idl_paragraph, content_chunks[1]);
        }
        ViewMode::Structured => {
            let member_query = app.member_search_query.to_lowercase();

            if !app.current_methods.is_empty() {
                let header_cells = ["Method Signature"]
                    .iter()
                    .map(|h| Cell::from(*h).style(Style::default().fg(Color::White))); // White text for header
                let header = Row::new(header_cells)
                    .style(Style::default().bg(Color::Blue)) // Unified Blue for header
                    .height(1);

                let filtered_methods: Vec<&MethodInfo> = app
                    .current_methods
                    .iter()
                    .filter(|m| m.name.to_lowercase().contains(&member_query))
                    .collect();

                let rows = filtered_methods.iter().map(|method| {
                    // Format:
                    // ƒ Name
                    //     ↓ Type Param
                    //     ↑ Type Param
                    //     -> ReturnType
                    let mut lines = Vec::new();

                    // Line 1: Function Name
                    lines.push(Line::from(vec![
                        Span::styled("ƒ ", Style::default().fg(Color::Magenta)),
                        Span::styled(
                            &method.name,
                            Style::default()
                                .fg(Color::Gray)
                                .add_modifier(Modifier::BOLD),
                        ),
                        Span::raw(" ("),
                    ]));

                    // Params (indented)
                    for param in &method.params {
                        let mut param_spans = Vec::new();
                        param_spans.push(Span::raw("    ")); // Indentation

                        // Icons for flags
                        if param.flags.contains(&"in".to_string()) {
                            param_spans.push(Span::styled("↓ ", Style::default().fg(Color::Green)));
                        }
                        if param.flags.contains(&"out".to_string()) {
                            param_spans.push(Span::styled("↑ ", Style::default().fg(Color::Red)));
                        }
                        if param.flags.contains(&"defaultvalue".to_string()) {
                            param_spans.push(Span::styled("* ", Style::default().fg(Color::Blue)));
                        }
                        if param.flags.contains(&"optional".to_string()) {
                            param_spans
                                .push(Span::styled("? ", Style::default().fg(Color::Yellow)));
                        }

                        param_spans.push(Span::styled(
                            format!("{} ", param.type_name),
                            Style::default().fg(Color::White),
                        ));
                        param_spans.push(Span::raw(&param.name));
                        param_spans.push(Span::raw(","));

                        lines.push(Line::from(param_spans));
                    }

                    // Return type
                    lines.push(Line::from(vec![
                        Span::raw("  ) -> "),
                        Span::styled(&method.ret_type, Style::default().fg(Color::Green)),
                    ]));

                    // Add an empty line separator
                    lines.push(Line::from(""));

                    let height = lines.len() as u16;
                    Row::new(vec![Cell::from(Text::from(lines))]).height(height)
                });

                let table = Table::new(rows, [Constraint::Percentage(100)])
                    .header(header)
                    .block(
                        Block::default()
                            .borders(Borders::ALL)
                            .border_style(content_border_style)
                            .title("Methods"),
                    )
                    .row_highlight_style(Style::default().bg(Color::Blue)); // Unified Blue

                f.render_stateful_widget(table, content_chunks[1], &mut app.content_table_state);

                // Render Scrollbar for Methods
                app.content_scroll_state = app
                    .content_scroll_state
                    .content_length(filtered_methods.len());
                if let Some(i) = app.content_table_state.selected() {
                    app.content_scroll_state = app.content_scroll_state.position(i);
                }
                f.render_stateful_widget(
                    Scrollbar::default()
                        .orientation(ScrollbarOrientation::VerticalRight)
                        .begin_symbol(Some("↑"))
                        .end_symbol(Some("↓")),
                    content_chunks[1],
                    &mut app.content_scroll_state,
                );
            } else if !app.current_enums.is_empty() {
                let header_cells = ["Name", "Value"]
                    .iter()
                    .map(|h| Cell::from(*h).style(Style::default().fg(Color::White))); // White text for header
                let header = Row::new(header_cells)
                    .style(Style::default().bg(Color::Blue)) // Unified Blue for header
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
                .row_highlight_style(Style::default().bg(Color::Blue)); // Unified Blue

                f.render_stateful_widget(table, content_chunks[1], &mut app.content_table_state);

                // Render Scrollbar for Enums
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
                    content_chunks[1],
                    &mut app.content_scroll_state,
                );
            } else {
                // Fallback to IDL if no structured data available
                let idl_paragraph = Paragraph::new(app.current_idl.as_str())
                    .block(
                        Block::default()
                            .borders(Borders::ALL)
                            .border_style(content_border_style)
                            .title("IDL Preview (No structured data)"),
                    )
                    .wrap(Wrap { trim: false });
                f.render_widget(idl_paragraph, content_chunks[1]);
            }
        }
    }

    // Global Search Popup
    if app.show_global_search {
        let area = centered_rect(60, 50, f.area());
        f.render_widget(Clear, area); // Clear background

        let block = Block::default()
            .title("Global Search (Esc to close)")
            .borders(Borders::ALL)
            .style(Style::default().bg(Color::Black)); // Darker background

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

        // Render Scrollbar for Global Search
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

    // Documentation Popup
    if app.show_doc {
        let area = centered_rect(90, 80, f.area());
        f.render_widget(Clear, area);

        let block = Block::default()
            .title(" Documentation (Enter to close) ")
            .borders(Borders::ALL)
            .style(Style::default().bg(Color::Black));

        let inner_area = block.inner(area);
        f.render_widget(block, area);

        // Layout:
        // Top: Description (20%)
        // Bottom: Split into Left (Signature) and Right (Parameters) (80%)
        let main_chunks = Layout::default()
            .direction(Direction::Vertical)
            .constraints([Constraint::Percentage(20), Constraint::Percentage(80)].as_ref())
            .split(inner_area);

        // --- Top: Description ---
        if let Some(provider) = &app.doc_provider {
            if let Some(name) = app.get_selected_name() {
                if let Some(doc) = provider.get_doc(&name) {
                    let desc_paragraph = Paragraph::new(doc.description)
                        .block(
                            Block::default()
                                .borders(Borders::BOTTOM)
                                .title(" Description "),
                        )
                        .wrap(Wrap { trim: true });
                    f.render_widget(desc_paragraph, main_chunks[0]);

                    // --- Bottom Split ---
                    // Left: Signature (30%)
                    // Right: Parameters (70%)
                    let bottom_chunks = Layout::default()
                        .direction(Direction::Horizontal)
                        .constraints(
                            [Constraint::Percentage(30), Constraint::Percentage(70)].as_ref(),
                        )
                        .split(main_chunks[1]);

                    // --- Bottom Left: Function Signature ---
                    let mut signature_lines = Vec::new();
                    let member_query = app.member_search_query.to_lowercase();
                    let filtered_methods: Vec<&MethodInfo> = app
                        .current_methods
                        .iter()
                        .filter(|m| m.name.to_lowercase().contains(&member_query))
                        .collect();

                    if let Some(idx) = app.content_table_state.selected() {
                        if let Some(method) = filtered_methods.get(idx) {
                            // Same styling as main view
                            let mut lines = Vec::new();
                            lines.push(Line::from(vec![
                                Span::styled("ƒ ", Style::default().fg(Color::Magenta)),
                                Span::styled(
                                    &method.name,
                                    Style::default()
                                        .fg(Color::Gray)
                                        .add_modifier(Modifier::BOLD),
                                ),
                                Span::raw(" ("),
                            ]));

                            for param in &method.params {
                                let mut param_spans = Vec::new();
                                param_spans.push(Span::raw("    "));
                                if param.flags.contains(&"in".to_string()) {
                                    param_spans.push(Span::styled(
                                        "↓ ",
                                        Style::default().fg(Color::Green),
                                    ));
                                }
                                if param.flags.contains(&"out".to_string()) {
                                    param_spans
                                        .push(Span::styled("↑ ", Style::default().fg(Color::Red)));
                                }
                                if param.flags.contains(&"defaultvalue".to_string()) {
                                    param_spans
                                        .push(Span::styled("* ", Style::default().fg(Color::Blue)));
                                }
                                if param.flags.contains(&"optional".to_string()) {
                                    param_spans.push(Span::styled(
                                        "? ",
                                        Style::default().fg(Color::Yellow),
                                    ));
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

                            signature_lines = lines;
                        }
                    }

                    let signature_paragraph = Paragraph::new(signature_lines)
                        .block(
                            Block::default()
                                .borders(Borders::RIGHT)
                                .title(" Signature "),
                        )
                        .wrap(Wrap { trim: false });
                    f.render_widget(signature_paragraph, bottom_chunks[0]);

                    // --- Bottom Right: Parameters Table ---
                    if !doc.parameters.is_empty() {
                        let header_cells = ["Name", "Type", "Description"]
                            .iter()
                            .map(|h| Cell::from(*h).style(Style::default().fg(Color::Yellow)));
                        let header = Row::new(header_cells).height(1).bottom_margin(1);

                        // Calculate column width for description
                        // Table width is bottom_chunks[1].width
                        // Constraints: 20%, 30%, 50%
                        // Description column is 50% of the table width
                        // We subtract 3 for borders/padding roughly
                        let table_width = bottom_chunks[1].width.saturating_sub(2);
                        let desc_col_width = (table_width as f32 * 0.5) as usize;
                        // Ensure at least some width
                        let desc_col_width = desc_col_width.max(10);

                        let rows = doc.parameters.iter().map(|param| {
                            // Use textwrap to calculate exact lines needed
                            let wrapped_desc = textwrap::wrap(&param.description, desc_col_width);
                            let height = (wrapped_desc.len() as u16).max(1);

                            let desc_text = wrapped_desc.join("\n");

                            let cells = vec![
                                Cell::from(Span::styled(
                                    &param.name,
                                    Style::default().add_modifier(Modifier::BOLD),
                                )),
                                Cell::from(Span::raw(&param.type_info)),
                                Cell::from(Text::from(desc_text)),
                            ];
                            Row::new(cells).height(height)
                        });

                        let table = Table::new(
                            rows,
                            [
                                Constraint::Percentage(20),
                                Constraint::Percentage(30),
                                Constraint::Percentage(50),
                            ],
                        )
                        .header(header)
                        .block(
                            Block::default()
                                .borders(Borders::LEFT)
                                .title(" Parameters "),
                        );

                        f.render_widget(table, bottom_chunks[1]);
                    } else {
                        f.render_widget(
                            Paragraph::new("No parameters documented.")
                                .block(Block::default().borders(Borders::LEFT)),
                            bottom_chunks[1],
                        );
                    }
                } else {
                    f.render_widget(Paragraph::new("No documentation found."), main_chunks[0]);
                }
            }
        } else {
            f.render_widget(
                Paragraph::new("Documentation provider not available."),
                main_chunks[0],
            );
        }
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
